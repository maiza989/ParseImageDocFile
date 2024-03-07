/*
 *                                                                                  --Microsoft.Office.Interop.Word Attempt--
 *                                        This code fucntion correctly, but some does not pick some images if the Warp Text format of the image is not "In line with Text"
 * 
 * using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace WordAutomation
{
    class Program
    {
        Application wordApp;
        Document doc;
        string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo2.jpg";
        string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";
        List<InlineShape> InlineShapesToDelete;                                                                                                   // List to hold inlines shapes to delete ( a picture, an OLE object, or an ActiveX control)

        public Program()
        {
            wordApp = new Application();
            InlineShapesToDelete = new List<InlineShape>();
        }// end of program construction
        public void Run()
        {
            try
            {
                if (Directory.Exists(folderPath))                                                                                           // Check if folder exists
                {
                    string[] files = Directory.GetFiles(folderPath, "*.doc");                                                               // Get all .doc files in the folder
                    foreach (string filePath in files)
                    {
                        ProcessDocument(filePath);                                                                                          // Call ProcessDocument to process each .doc file
                    }
                }
                else
                {
                    Console.WriteLine("Folder not found.");
                }
            }// end of outter try
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }// end of catch
            finally
            {
                CleanupApplication();
            }// end of finally
        }// end of Run 
        private void ProcessDocument(string filePath)
        {
            Console.WriteLine($"\nProcessing File: {Path.GetFileName(filePath)}");
            try
            {
                doc = wordApp.Documents.Open(filePath);                                                                                      // Open the Word document
                Console.WriteLine($"\tOpened File: {Path.GetFileName(filePath)}");
                InlineShapesToDelete.Clear();                                                                                                // Reset shapes to delete for each document

                foreach (Section section in doc.Sections)
                {
                    ProcessSection(section, filePath, imagePath);                                                                            // Iterate through all inline shapes in the section
                }

                if (InlineShapesToDelete.Count == 0)
                {
                    Console.WriteLine($"\t\tNo Picture Found in doc {Path.GetFileName(filePath)}");
                }

                foreach (InlineShape shapeToDelete in InlineShapesToDelete)
                {
                    shapeToDelete.Delete();                                                                                                   // Delete the old pictures after iterating through all shapes
                }
                doc.Save();                                                                                                                   // Save Documents
                Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
            }// end of inner try 
            catch (Exception ex)
            {
                Console.WriteLine($"Error: processing file {Path.GetFileName(filePath)}: {ex.Message}");

            }// end of catch
            finally
            {
                CleanupDocument();
            }// end of finally
        }// end of ProcessDocument
        private void ProcessSection(Section section, string filePath, string imagePath)
        {
            //bool replaced = false;

            int imageCount = 0;
            foreach (InlineShape shape in section.Range.InlineShapes)                                                                          // Iterate through all inline shapes in the section
            {
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)                                                                      // Check if the shape is a picture
                {
                    

                    shape.Select();
                    InlineShapesToDelete.Add(shape);                                                                                           // Add the shape to delete list
                    shape.Range.InlineShapes.AddPicture(imagePath);                                                                            // Add new image
                    imageCount++;
                    Console.WriteLine($"\t\tImage \"{imageCount}\" Changed in file: {Path.GetFileName(filePath)}");
                    
                    // replaced = true; 
                }
            }
        }// end of ProcessSection
       *//* private bool IsImageBehindText(InlineShape shape)
        {
            return shape.Anchor is ShapeRange;  // Check if the shape's anchor is a ShapeRange object
        }*//*
        private void CleanupDocument()
        {

            if (doc != null)
            {
                doc.Close();                                                                                                                  // Close Documents 
            }
        }// end of CleanupDocument
        private void CleanupApplication()
        {
            if (wordApp != null)
            {
                wordApp.Quit();                                                                                                               // Close Application
                Marshal.ReleaseComObject(wordApp);                                                                                            // Release COM Objects
            }
        }// end of CleanupApplication

        static void Main(string[] args)
        {
            Program program = new Program();
            program.Run();
        }
    }
}
*/

/*
 * 
 *                                                                                          --Spire.doc Attempt--
 */
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace WordAutomation
{
    class Program
    {
        private string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo2.jpg";
        private string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";
        int imageCount = 0;
        int fileCount = 0;
        public Program() 
        {
            
        }// end of Program construction
        public void Run()
        {
            try
            { 
                if (Directory.Exists(folderPath))                                                                               // Check if Folder exist
                {
                    string[] files = Directory.GetFiles(folderPath, "*.doc");                                                   // Get all .doc files in folder
                    foreach (string filePath in files)
                    {
                        ProcessDocument(filePath);                                                                              // Process each files in the folder
                    }// end of foreach
                }// end of if statement
                else
                {
                    Console.WriteLine("Folder not found.");
                }// end of if-else
            }// end of Try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Error: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of Catch
        }// end of Run

        private void ProcessDocument(string filePath)
        {
            fileCount ++;
            Console.WriteLine($"\nProcessing File \"{fileCount}\": {Path.GetFileName(filePath)}");
            try
            {
                Document doc = new Document();
                doc.LoadFromFile(filePath);                                                                                        // Load file into Document Object

                if( doc != null )
                {
                    foreach (Section section in doc.Sections)
                    {
                        ProcessSection(section, filePath);                                                                         // Go through each section in side the Word document
                    }// end of  outter-outter foreach
                    if(imageCount == 0)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"\t\t\tNo Picture Found in doc {Path.GetFileName(filePath)}");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of if-statement
                    doc.SaveToFile(filePath, FileFormat.Doc);
                    imageCount = 0;
                    Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
                }
            }// end of Try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"!!!Error: processing file {Path.GetFileName(filePath)}: {ex.Message}!!!");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of ProcessDocument

        private void ProcessSection(Section section, string filePath)
        {
            try
             {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    foreach (DocumentObject docObj in paragraph.ChildObjects)
                    {
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            imageCount++;                                                                                                       // Increment image count if a picture is found
                            DocPicture newPicture = docObj as DocPicture;                                                                       // Create a new DocPicture with the desired image                     
                            newPicture.LoadImage(Image.FromFile(imagePath));                                                                    // Replace image found in document with new image
                            Console.ForegroundColor = ConsoleColor.DarkCyan;
                            Console.WriteLine($"\tChanged Image \"{imageCount}\" in file: {Path.GetFileName(filePath)}");
                            Console.ForegroundColor = ConsoleColor.Gray;

                        }// end of if-statement
                    }// end of inner for each
                }// end of outter foreach
                if(imageCount > 0)
                {
                Console.ForegroundColor = ConsoleColor.DarkMagenta;
                Console.WriteLine($"\t\tTotal images found in {Path.GetFileName(filePath)}: {imageCount}");
                Console.ForegroundColor = ConsoleColor.Gray;
                }
            }// end of Try
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"Error looking for a picture in file {Path.GetFileName(filePath)}: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of ProcessSection

        static void Main(string[] args)
        {
            Program program = new Program();
            program.Run();
        }// end of main method
    }// end of Program Class
}// end of namespace



