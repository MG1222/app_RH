# Relance RH Application

## Overview
Relance RH is a mini-application designed for recording and creating Excel files with candidate information. It streamlines the process of managing recruitment data, making it easier for HR departments to handle candidate information efficiently.

## Supported CV Format

The application supports processing CVs in the format provided in the demo file `demo-fichier_entretien.xlsx`. It extracts and utilizes the following information:

- **Directory and Profile**: The application expects the CVs to be organized in directories named after the 
  candidate's profile, following the format `D:\Relance_RH\Ressource\31 - sourcing\2024\<Nom Prenom - Profil>`.
- **Personal Information**: Extracts name, surname, telephone number, and email address.
- **Availability**: Determines the candidate's availability.
- **Interview Dates and Managers**: Extracts up to three interview dates along with the corresponding managers.
- **Last Interview Calculation**: The application calculates the date of the last interview based on the provided dates.

Ensure that the CVs are formatted according to the `demo-fichier_entretien.xlsx` example for optimal processing.

## Features
- **Candidate Information Management:** Easily record and store detailed information about candidates.
- **Excel File Creation:** Automatically generate Excel files with candidate information for easy sharing and analysis.
- **Folder Selection:** Users can select specific folders to process Excel files, enhancing the application's flexibility.
- **Progress Tracking:** Visual progress bar to track the processing of files.
- **Data Validation:** Includes functionality to verify phone numbers, emails, and ensure the correct format of Excel files.

## Installation

### Prerequisites
- Python 3.12
- pip for installing dependencies

### Steps
1. Clone the repository to your local machine: `git clone https://github.com/MG1222/app_RH.git`
2. Navigate to the cloned directory: `cd app_RH`

### Setting Up a Virtual Environment

After forking the repository, it's recommended to set up a virtual environment for the project. This helps in managing dependencies and ensuring that the project runs smoothly on your machine.

1. **Navigate to the project directory**:
   - Open a terminal and change to the project directory with `cd path/to/project`.

2. **Create the virtual environment**:
   - Run `python -m venv env` to create a new virtual environment named `env`.

3. **Activate the virtual environment**:
   - On Windows, execute `.\env\Scripts\activate`.
   - On macOS and Linux, use `source env/bin/activate`.

4. **Install dependencies**:
   - With the virtual environment activated, install the project dependencies by running `pip install -r requirements.txt`.

## Ongoing Development

We are currently implementing the feature to allow users to customize the cell locations from which candidate information is extracted. This enhancement aims to provide greater flexibility in handling various CV formats and ensuring that our application can adapt to different document structures efficiently.

