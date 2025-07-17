

/*
 * 
 *                                                                                          --Spire.doc Attempt--
 */
using System.Drawing;

using System.Threading.Tasks;

using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Document = Spire.Doc.Document;
using Section = Spire.Doc.Section;
using Paragraph = Spire.Doc.Documents.Paragraph;
using Field = Spire.Doc.Fields.Field;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office;
using Microsoft.Office.Interop.Word;

using Aspose.Words;
namespace WordAutomation
{
    class Program
    {
        private string imagePath = @"\\lmusrdata\User_Profiles\malghamgham\Desktop\My work - Maitham\GIF\Logo\Logo Final Small-Edited.jpg";                                             // Image Path                                               
        private string folderPath = @"\\lmusrdata\User_Profiles\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";                                                             // Original Word Documents Folder
       // private string backupPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder\Backup";                                                              // **Optional** Destination path for file backup 
        private string errorTextboxFolderPath = @"\\lmusrdata\User_Profiles\malghamgham\Desktop\My work - Maitham\Projects\TestFolder\ErrorTest\TextboxError";                          // **Optional** Destination path for file with Textbox in them
        private string errorNullRefFolderPath = @"\\lmusrdata\User_Profiles\malghamgham\Desktop\My work - Maitham\Projects\TestFolder\ErrorTest\NullReferenceError";                    // **Optional** Destination path for file throwing null reference error

        int imageCount = 0;                                                                                                                                                     // Image changed counter
        int fileCount = 0;                                                                                                                                                      // File processed counter
        int imageMaxWidth = 250;                                                                                                                                                // Variable store the MAXIMUM width the image you want to change has. 
        int imageMinWidth = 90;                                                                                                                                                 // Variable store the MINIMUM width the image you want to change has
        int imageMaxHeight = 100;                                                                                                                                               // Variable store the MAXIMUM height the image you want to change has
        int imageMinHeight = 40;                                                                                                                                                // Variable store the MINIMUM height the image you want to change has

        bool isMoved = false;                                                                                                                                                   // Boolean check if file is moved.
        bool isImageChanged = false;                                                                                                                                            // Boolean check if image change in file
        Document doc;                                                                                                                                                           // Document Object

        public Program() 
        {
            doc = new Document();
        }// end of Program construction

        /*
         * A method that make sure folder exist and files in that certain folder exists.
         */
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            try
            { 
                if (Directory.Exists(folderPath))                                                                                    // Check if Folder exist
                {
                    string[] files = Directory.GetFiles(folderPath, "*.doc*");                                                       // Get all .doc or .docx files in folder
                    if(files.Length > 0)
                    {
                        foreach (string filePath in files)
                        {
                            ProcessDocument(filePath);                                                                               // Process each files in the folder
                        }// end of foreach

                        /*  Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 2 }, filePath =>
                          {
                              // Create new instance per thread to avoid shared state
                              var programInstance = new Program();
                              programInstance.ProcessDocument(filePath);
                          });*/
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"No files found in folder: {Path.GetFileName(folderPath)}");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }
                }// end of if statement
                else
                {
                    Console.WriteLine("Folder not found.");
                }// end of if-else
            }// end of Try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Error finding folder: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of Catch
        }// end of Run
        /*
         * A method that start the file processing of each section of the word documents.
         */
        private void ProcessDocument(string filePath)
        {
            bool wasOriginallyDoc = false;
            string originalDocPath = filePath;
            string tempDocxPath = filePath;
            fileCount ++;
            Console.WriteLine($"\nProcessing File \"{fileCount}\": {Path.GetFileName(filePath)}");
            try
            {
                if( doc != null && doc.Sections != null)
                {

                    // Convert .doc to .docx before processing
                    if (filePath.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) &&
                        !filePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"\tConverting {Path.GetFileName(filePath)} to .docx for safe editing...");
                        Console.ForegroundColor = ConsoleColor.Gray;


                        //ConvertDocToDocxAspose(filePath); // Your helper
                        ConvertDocToDocxInterop(filePath);
                        filePath = Path.ChangeExtension(filePath, ".docx");
                        wasOriginallyDoc = true;
                    }

                    // Now safely load .docx using Spire
                    doc = new Document();
                    doc.LoadFromFile(filePath);
                    //CreateCopyOfFile(folderPath, backupPath);

                    foreach (Section section in doc.Sections)
                    {
                        ProcessSection(section, filePath);                                                                                                                  // Go through each section in side the Word document
                    }// end of  outter-outter foreach
                    if(imageCount == 0 && !isMoved)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"\tNo Picture Found in doc {Path.GetFileName(filePath)}");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of if-statement
                    if(!isMoved)                                                                                                                                            // If isMoved is true dont save file. This is a check for if file was moved an converted to .docx 
                    {
                        FileFormat fileFormat = filePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? FileFormat.Docx : FileFormat.Doc;                          // Determine the file format based on the file extension
                        doc.SaveToFile(filePath, fileFormat);
                        
                        
                    }// end of if-statement isMoved
                    if (wasOriginallyDoc)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"\tReverting edited file back to .doc...");
                        Console.ForegroundColor = ConsoleColor.Gray;
/*
                        originalDocPath = Path.ChangeExtension(filePath, ".doc"); // .docx -> .doc
                        doc.LoadFromFile(filePath); // Reload just in case
                        doc.SaveToFile(originalDocPath, FileFormat.Doc);
                        File.Delete(filePath); // Delete the temporary .docx*/
                        
                        ConvertDocxToDocInterop(filePath);
                    }

                    Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
                    imageCount = 0;
                    isMoved = false;
                    wasOriginallyDoc = false;                                                                                                                  
                }// end of If-Statement
            }// end of Try
            catch (NullReferenceException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"!!!Error processing file {Path.GetFileName(filePath)}: {ex.Message}!!!");
                Console.ForegroundColor = ConsoleColor.Gray;
                MoveFileToTargetFolder(filePath);
            }// end of catch
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"!!!Error processing file {Path.GetFileName(filePath)}: {ex.Message}!!!");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of ProcessDocument

        /*
         * A method that process each paragraph of the word documents. This method will detect images or textboxes and actions will be taken based on that. 
         */
        private void ProcessSection(Section section, string filePath)
        {
            try
             {
                for (int i = 0; i < section.Paragraphs.Count; i++) // Iterate using index to avoid collection modification issues
                {
                    Paragraph paragraph = section.Paragraphs[i];
                    for (int j = 0; j < paragraph.ChildObjects.Count; j++) // Iterate using index to avoid collection modification issues
                    {
                        DocumentObject docObj = paragraph.ChildObjects[j];
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture newPicture = docObj as DocPicture;                                                                           // Create a new DocPicture with the desired image                     
                            if (newPicture.Width < imageMaxWidth && newPicture.Width > imageMinWidth
                                && newPicture.Height < imageMaxHeight && newPicture.Height > imageMinHeight)                                        // Only change the image the width is bigger than 90px and smaller than 250px and height smaller than 100px and larger than 50px
                            {
                                isImageChanged = true;
                                imageCount++;                                                                                                       // Increment image count if a picture is found
                                newPicture.LoadImage(Image.FromFile(imagePath));                                                                    // Replace image found in document with new image
                                Console.ForegroundColor = ConsoleColor.DarkCyan;
                                Console.WriteLine($"\tChanged Image \"{imageCount}\" in file: {Path.GetFileName(filePath)}");
                                Console.ForegroundColor = ConsoleColor.Gray;
                                continue;                                                                                                           // Continue to the next iteration if the image is changed
                            }// end of inner if-statement 
                        }// end of if-statement
                        else if (docObj.DocumentObjectType == DocumentObjectType.TextBox)
                        {
                            Console.WriteLine($"\t\tTextbox detected in file: {Path.GetFileName(filePath)}.");
                            ConvertToDocx(filePath);                                                                                                // Convert the file to .docx
                            isMoved = true;
                            return;
                        }
                    }// end of inner for each
                }// end of outter foreach
                if(isImageChanged == true && imageCount > 0)
                {
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine($"\t\tTotal images found in {Path.GetFileName(filePath)}: {imageCount}");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    isImageChanged = false;
                }
            }// end of Try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\t\tError looking for a picture in file {Path.GetFileName(filePath)}: {ex.Message}\n\tMoving to \"Error\" folder");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of ProcessSection

        /*
         * This method was made due to some files giving null references for some reason. This method hand pick them and move them to a saparate folder for manual inspection
         */
        private void MoveFileToTargetFolder(string filePath)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                string targetFilePath = Path.Combine(errorNullRefFolderPath, fileName);

                File.Move(filePath, targetFilePath);                                                                                               // Move file from originl path to target path
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\tFile \"{fileName}\" moved to Error folder.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\tError moving file: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of MoveFileToTargetFolder

        /*
         * A method that move move a file to saparate folder and conver it the file to .docx format. 
         * This Method was created due to some files having textbox and when doc spire goes through the code and save the files as .doc file. It breaks the format. 
         * Saving the file as .docx seems to perserve the format of the templates.
         * Manual insepction might be required. 
         */
        private void ConvertToDocx(string filePath)
        {
            bool isMoved = false;
            try
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string targetFilePath = Path.Combine(errorTextboxFolderPath, $"{fileName}.docx");

                if (!isMoved)                                                                                                                      // Save the file with .docx extension
                {
                    File.Move(filePath, targetFilePath);                                                                                           // Move file from originl path to target path
                    doc.SaveToFile(targetFilePath, FileFormat.Docx);                                                                               // Covert original File to .docx format
                    isMoved = true;                                                                                                                 
                    imageCount = 0;
                }// end of if-statement not isMoved
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\t\tFile \"{Path.GetFileName(filePath)}\" converted to .docx and saved in: {Path.GetFileName(targetFilePath)}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\t\tError converting file: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of ConvertToDocx

        /*
         *                              *****NOT USED*****
         * A method to copy files being open and create a copy on the backup folder.
         */
        private void CreateCopyOfFile(string sourceFilePath, string destinationFolderPath)
        {
            try
            {
                string fileName = Path.GetFileName(sourceFilePath);
                string destinationFilePath = Path.Combine(destinationFolderPath, fileName);

                File.Copy(sourceFilePath, destinationFilePath, true); // Set overwrite to true to overwrite existing files
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"\tFile \"{fileName}\" copied to {Path.GetFileName(destinationFolderPath)}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error copying file: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of CreateCopyOfFile

        private bool IsBarcodeField(string fieldText)
        {
            fieldText = fieldText.Trim();

            // Print debugging info
            // Console.WriteLine($"\t\tChecking merge field: \"{fieldText}\"");

            // Check if both Fileno and LLCODE exist in the same field
            return fieldText.Equals("MERGEFIELD FILENO") && fieldText.Equals("MERGEFIELD LLCODE");
        }

        void ConvertDocToDocxInterop(string docPath)
        {
            var wordApp = new Word.Application();
            wordApp.Visible = false;

            var doc = wordApp.Documents.Open(docPath);
            string newPath = Path.ChangeExtension(docPath, ".docx");

            doc.SaveAs2(newPath, Word.WdSaveFormat.wdFormatXMLDocument);
            doc.Close(false);
            wordApp.Quit();

            File.Delete(docPath); // delete original .doc
        }

        void ConvertDocxToDocInterop(string docxPath)
        {
            var wordApp = new Word.Application();
            wordApp.Visible = false;

            var doc = wordApp.Documents.Open(docxPath);
            string originalDocPath = Path.ChangeExtension(docxPath, ".doc");

            doc.SaveAs2(originalDocPath, Word.WdSaveFormat.wdFormatDocument);
            doc.Close(false);
            wordApp.Quit();

            File.Delete(docxPath); // delete temporary .docx
        }
       /* void ConvertDocToDocxAspose(string docPath)
        {
            var doc = new Aspose.Words.Document(docPath);
            string newPath = System.IO.Path.ChangeExtension(docPath, ".docx");
            doc.Save(newPath, Aspose.Words.SaveFormat.Docx);
            System.IO.File.Delete(docPath);
        }

        void ConvertDocxToDocAspose(string docxPath)
        {
            var doc = new Aspose.Words.Document(docxPath);
            string originalDocPath = System.IO.Path.ChangeExtension(docxPath, ".doc");
            doc.Save(originalDocPath, Aspose.Words.SaveFormat.Doc);
            System.IO.File.Delete(docxPath);
        }*/

        static void Main(string[] args)
        {
            Program program = new Program();
            program.Run();
        }// end of main method
    }// end of Program Class
}// end of namespace



