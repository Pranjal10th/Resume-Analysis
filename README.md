# Resume Shortlister Code Documentation
# Overview
This code is designed to shortlist candidates based on their resumes. It uses a graphical user interface (GUI) built with Tkinter to input criteria and select a folder containing resumes. The code then checks each resume against the entered criteria and copies shortlisted resumes to a new "Shortlisted" folder.

# Code Structure
The code is divided into several sections:

1. Importing Libraries: The code starts by importing the necessary libraries, including Tkinter for the GUI and os for file operations.
2. Function Definitions: The code defines several functions to perform specific tasks, such as getting the folder path, getting criteria, checking resumes, and shortlisting candidates.
3. GUI Creation: The code creates a GUI with fields for folder path, skills, experience, and education.
4. Shortlisting Logic: The code checks each resume against the entered criteria and copies shortlisted resumes to a new "Shortlisted" folder.

# Function Documentation
get_folder_path()
- Purpose: Get the folder path from the user.
- Parameters: None
- Returns: None
- Description: This function opens a file dialog for the user to select a folder. The selected folder path is then inserted into the folder entry field.

get_criteria()
- Purpose: Get the criteria from the user.
- Parameters: None
- Returns: A dictionary containing the criteria (skills, experience, education)
- Description: This function retrieves the values from the skills, experience, and education entry fields and returns them as a dictionary.

check_resume(resume_path, criteria)
- Purpose: Check a resume against the entered criteria.
- Parameters:
    - resume_path (str): The path to the resume file.
    - criteria (dict): A dictionary containing the criteria (skills, experience, education)
- Returns: A boolean indicating whether the resume matches the criteria
- Description: This function reads the resume file and checks if it contains the specified skills, experience, and education. It returns True if the resume matches the criteria and False otherwise.

shortlist_candidates()
- Purpose: Shortlist candidates based on their resumes.
- Parameters: None
- Returns: None
- Description: This function gets the folder path and criteria from the user, checks each resume in the folder against the criteria, and copies shortlisted resumes to a new "Shortlisted" folder.

# GUI Documentation
The GUI has the following fields:

- Folder Path: A field to enter the folder path containing the resumes.
- Skills: A field to enter the required skills (comma-separated).
- Experience: A field to enter the required experience (years).
- Education: A field to enter the required education.

The GUI also has a "Shortlist Candidates" button that triggers the shortlisting logic.

# Usage
To use this code, follow these steps:

1. Run the code to launch the GUI.
2. Select a folder containing resumes by clicking the "Browse" button.
3. Enter the required skills, experience, and education in the corresponding fields.
4. Click the "Shortlist Candidates" button to start the shortlisting process.
5. The shortlisted resumes will be copied to a new "Shortlisted" folder within the selected folder.
