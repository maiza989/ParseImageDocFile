# Word Automation Readme

## Overview
This C# program is designed to automate tasks related to processing Microsoft Word documents (.doc and .docx files) using the Spire.Doc library. It performs various operations on Word documents located within a specified folder, such as replacing images, detecting text boxes, and converting files to .docx format.

## Features
1. **Image Replacement**: The program identifies images within Word documents and replaces them with a designated image file.
2. **Text Box Detection**: It detects text boxes within documents and converts the entire document to .docx format if a text box is found.
3. **Error Handling**: Error handling mechanisms are in place to manage exceptions during document processing.
4. **File Management**: Files causing null reference errors are moved to a separate error folder for manual inspection. Additionally, documents with text boxes are moved to another folder after conversion.

## Usage
1. **Setup**: Ensure that the Spire.Doc library is properly installed and referenced in the project.
2. **Configuration**: Update the file paths (`imagePath`, `folderPath`, `errorTextboxFolderPath`, `errorNullRefFolderPath`) according to your environment.
3. **Execution**: Run the program. It will iterate through all .doc and .docx files in the specified folder, processing each document according to the defined operations.
4. **Output**: The program will print relevant information to the console, including file processing status, error messages, and actions taken.

## Dependencies
- Spire.Doc: This library is used for working with Word documents in the .NET environment. Ensure it is installed and referenced in the project.

## Notes
- Make sure to handle exceptions gracefully, as the program relies on external file operations and document processing, which may encounter errors if not set up correctly.
- Review the code comments for detailed explanations of each method and functionality.


