
import subprocess
import time
import platform
import zipfile
import os



def split_exe(input_file_path, output_folder, chunk_size_mb=20):
    if not os.path.exists(input_file_path):
        print(f"Error: File not found - {input_file_path}")
        return
    if not os.path.exists(output_folder):                                                   # Create the output folder if it doesn't exist
        os.makedirs(output_folder)
    chunk_size_bytes = chunk_size_mb * 1024 * 1024                                          # Calculate the chunk size in bytes
    with open(input_file_path, 'rb') as input_file:                                         # Open the input file for reading in binary mode
        file_content = input_file.read()                                                    # Read the input file content
        total_size = len(file_content)                                                      # Get the total size of the file
        num_chunks = (total_size + chunk_size_bytes - 1) // chunk_size_bytes                # Calculate the number of chunks needed

        for i in range(num_chunks):                                                         # Split the file into chunks
            start_index = i * chunk_size_bytes                                              # Calculate the start and end indices for each chunk
            end_index = min((i + 1) * chunk_size_bytes, total_size)
            output_file_name = f"{os.path.basename(input_file_path)}_part_{i + 1}.zip"      # Create the output file name
            output_file_path = os.path.join(output_folder, output_file_name)
            with zipfile.ZipFile(output_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:  # Write the chunk to the output file
                zip_file.writestr(os.path.basename(input_file_path), file_content[start_index:end_index])
            print(f"Chunk {i + 1} created: {output_file_path}")
    print("Splitting complete.")

def combine_chunks(input_folder: str, output_file: str) -> None:
    # Get a list of all zip files in the input folder
    zip_files = [f for f in os.listdir(input_folder) if f.endswith(".zip")]
    # Sort the zip files based on part number
    zip_files.sort(key=lambda x: int(x.split('_')[-1].split('.')[0]))
    # Open the output file for writing in binary mode
    with open(output_file, 'wb') as ofile:
        # Iterate through the sorted zip files and append their content to the output file
        for zip_file_name in zip_files:
            zip_file_path = os.path.join(input_folder, zip_file_name)
            with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
                # Assume there is only one file in the zip archive (the original exe)
                file_content = zip_file.read(zip_file.namelist()[0])
                ofile.write(file_content)
            print(f"Chunk {zip_file_name} added to the output file.")
    print("Combining complete.")

def make_broken_zip():
    installer_origin_exe = "QBSDK160_x64.exe"
    installer_output_folder = os.getcwd()
    split_exe("QBSDK160_x64.exe", installer_output_folder)
    #combine_chunks(installer_output_folder, installer_origin_exe)




#===============================================================
# Installation Check
#===============================================================

def is_windows():
    return platform.system().lower() == 'windows'

def ensure_installation() -> None:
    """Make sure that the quickbooks installation is 
       complete and ready to use. """
    
    qb_sdk_path = "C:\\Program Files\\Intuit\\IDN\\QBSDK16.0"
    installer_home_exe = os.path.join(os.getcwd(),"QBSDK160_x64.exe")
    
    # check to see if the quickbooks sdk is available
    if not os.path.exists(qb_sdk_path):
        print("Running installer...")
        combine_chunks(os.getcwd(), installer_home_exe)
        installer_path = "QBSDK160_x64.exe" # create installer path
        install_process = subprocess.Popen([installer_path])                      # run the installer
        install_progress = install_process.poll() 
        while install_progress is None:                                           # wait for installer to finish
            install_progress = install_process.poll()
            time.sleep(0.50)
        print(install_process.returncode)

        qb_xml_requester_path = "C:\\Program Files (x86)\\Intuit\\IDN\\QBSDK16.0\\tools\\installers"
        for root, dirs, files in os.walk(qb_xml_requester_path):
            for file in files:
                print("installing ", file)
                exepath = os.path.join(root, file)
                installing_process = subprocess.Popen([exepath])
                installing_progress = installing_process.poll()
                while installing_progress is None:
                    installing_progress = installing_process.poll()
                    time.sleep(0.50)
                print(install_process.returncode)

        print("installation complete")
    else:
        print("installation checked and passed successfully") 

def precheck() -> None:
    if is_windows():
        print("Running on Windows.")
    else:
        print("Not running on Windows.")
        exit()
    installer_path = os.path.join(os.getcwd(), "QBSDK160_64.exe")
    if not os.path.exists(installer_path):
        combine_chunks(os.getcwd(), installer_path)

    try:
        ensure_installation()
    except Exception as e:
        print(e)
        print("installation failed to proceed or pass.")
        exit()


