import os
import shutil
import fnmatch
import logging
from pdfminer.high_level import extract_text
import docx
import concurrent.futures
from multiprocessing import cpu_count

# Define the keywords and file extensions to search for
keywords = set(["wetland", "dwg", "figure", "report", "final", "treatment", "design", "analysis", "analyses", "guidelines",
    "plan", "section", "fig", "elevation", "layout", "sketch", "memo", "schematic", "detail", "blueprint", "diagram", "map", 
    "plot", "assessment", "evaluation", "study", "summary", "manual", "guide", "standard", "protocol", "regulation", 
    "procedure", "method", "instruction", "steps", "criteria"])
file_extensions = [".dwg", ".docx", ".pdf", ".xlsx", ".prs", ".rpt", ".skt", ".dnt", ".dat"]

# Define the source and destination directories
source_directories = [
    r"\\ae.ca\data\projects\edm\20198403\00_Ft_McKay_Wetland"
]
destination_directory = r"\\ae.ca\data\working\gpr\2023-3668-00\_incoming_data\Background Information from Previous Wetland Projects"

# Set up logging to ensure progress updates are provided
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to search for keywords within PDF contents using pdfminer.six
def search_pdf(file_path, keywords):
    try:
        text = extract_text(file_path)
        return any(keyword in text.lower() for keyword in keywords)
    except Exception as e:
        logging.error(f"Error reading PDF file {file_path}: {e}")
        return False

# Function to search for keywords within Word document contents
def search_docx(file_path, keywords):
    try:
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            if any(keyword in para.text.lower() for keyword in keywords):
                return True
    except Exception as e:
        logging.error(f"Error reading Word file {file_path}: {e}")
    return False

# Function to search for files based on keywords and file extensions
def search_files(directory, keywords, file_extensions):
    matches = []
    skip_words = set(["meeting", "change order", "invoice", "pmp", "project management plan",
                      "cover letter", "record of meeting", "rom", "pjha", "psc", "dha", "email", "agr", "qaqc",
                      "aepea authorization","inspection report","receipt","timesheet", "agd", "pio", "pco",
                      "standard fee","IMG","contractor schedule","pdpl", "resume", "frm", "standard fee", "prp", "construction report", "contract", "ltr"])
    
    skip_folders = set(["email","pm","bid documents", "invoicing", "data collection", "cost_est", "request for information", "arc map",
                        "proposed change notices","om manuals", "contract_admin", "proposal", "purchase orders", "change orders",
                        "contract administration","bids received evaluation", "regulatory permitting", "regulatorypermitting",
                        "field orders","construction inspection", "safety", "resume", "field form", "contract admin"])
    
    for root, dirnames, filenames in os.walk(directory):
        # Skip folders based on the skip_folders set
        if any(skip_folder in root.lower() for skip_folder in skip_folders):
            continue
        
        # Skip folders mentioning invoices or photos
        if any(skip_folder in root.lower() for skip_folder in ["invoice","photos"]):
            continue
        
        for filename in filenames:
            # Skip temporary Word files
            if filename.startswith("~$"):
                continue
            
            # Skip emails and invoices
            if filename.lower().endswith(".msg") or "invoice" in filename.lower():
                continue
            
            # Skip documents mentioning specific words
            if any(skip_word in filename.lower() for skip_word in skip_words):
                continue
            
            # Skip files with .db and .txt extensions
            if filename.lower().endswith(( ".db",".txt",".eml",".png",".jpg",".json",".html",".pyc",".py",".js",
                                          ".yaml",".ipynb",".svg",".pyi",".tcl",".img",".fax",".rom",".ppc", ".tif",
                                          ".rfc",".dbf", ".ovr", ".pgwx", ".sbx",".shx",".cpg",".dbf",".mxd",".sbn",".shp", ".ecw", ".eww",
                                          ".rfi",".prj",".lyr", ".xml", ".gdb", ".lock", ".sr", ".jgwx", ".kmx", ".kmz", ".r", ".idx", ".jp2")):
                continue
            
            # Always include .dwg files
            if filename.lower().endswith(".dwg"):
                matches.append(os.path.join(root, filename))
                continue
            
            # Check file contents for keywords
            file_path = os.path.join(root, filename)
            if filename.lower().endswith(".pdf"):
                if search_pdf(file_path, keywords):
                    matches.append(file_path)
            elif filename.lower().endswith(".docx") and search_docx(file_path, keywords):
                matches.append(file_path)
            elif any(keyword in filename.lower() for keyword in keywords) or any(fnmatch.fnmatch(filename.lower(), f"*{ext}") for ext in file_extensions):
                matches.append(file_path)
    
    return matches

# Function to copy files to the destination directory
def copy_files(files, destination_directory, project_name):
    project_folder = os.path.join(destination_directory, project_name)
    os.makedirs(project_folder, exist_ok=True)
    
    for file in files:
        # Copy the file to the project subfolder
        try:
            shutil.copy(file, project_folder)
            logging.info(f"Copied file: {file}")
        except FileNotFoundError as e:
            logging.error(f"File not found: {file}. Skipping this file.")
        except Exception as e:
            logging.error(f"An error occurred while copying {file}: {e}")

# Example usage
project_name = "2019-8403-00 - Fort McKay Wetland Monitoring"
with concurrent.futures.ThreadPoolExecutor(max_workers=cpu_count()) as executor:
    futures = [executor.submit(search_files, source_directory, keywords, file_extensions) for source_directory in source_directories]
    
    for future in concurrent.futures.as_completed(futures):
        try:
            matched_files = future.result()
            copy_files(matched_files, destination_directory, project_name)
        except Exception as e:
            logging.error(f"An error occurred while processing: {e}")

logging.info("Files have been successfully copied to the destination directory.")
