
import os
import subprocess


if __name__ == "__main__":

        file_path = os.path.abspath("writerbird-word-reg.txt")

        #print(file_path)
        
        # Read the entire file into memory
        with open(file_path, "r") as f:
            content = f.readlines()

        # Update the "Url" line with correcly formatted shared path variable
        for i, line in enumerate(content):
            if line.startswith('"Url"='):
                content[i] = line.replace('\\', '\\\\')  # Replace all single backslashes with double backslashes
                #print(content[i])
                break           
            
        # Write the updated content back to the file
        with open(file_path, "w") as f:
            f.writelines(content)        

        # Change the file extension to ".reg"
        os.rename(file_path, os.path.splitext(file_path)[0] + ".reg")

        # Set the path to the .reg file. This is the file that adds writerbird to the MS Word add-ins
        reg_file_path = os.path.abspath("writerbird-word-reg.reg")

        # Use subprocess to execute the reg.exe command
        subprocess.call(["reg.exe", "import", reg_file_path]) 

        # Rename back to .txt
        os.rename(reg_file_path, os.path.splitext(reg_file_path)[0] + ".txt")
        
        print('Installer complete. You can close the terminal')
