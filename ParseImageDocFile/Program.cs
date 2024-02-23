using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Core.Shape;



namespace WordAutomation
{
    class Program
    {
        Application wordApp;
        Document doc;
        string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo.png";
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
                InlineShapesToDelete.Clear();                                                                                                      // Reset shapes to delete for each document

                foreach (Section section in doc.Sections)
                {
                    ProcessSection(section, filePath);                                                                                       // Iterate through all inline shapes in the section
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
        private void ProcessSection(Section section, string filePath)
        {
            //bool replaced = false;
            int imageCount = 0;
            foreach (InlineShape shape in section.Range.InlineShapes)                                                                         // Iterate through all inline shapes in the section
            {
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)                                                                     // Check if the shape is a picture
                {
                    shape.Select();
                    InlineShapesToDelete.Add(shape);                                                                                          // Add the shape to delete list
                    shape.Range.InlineShapes.AddPicture(imagePath);                                                                           // Add new image
                    imageCount++;
                    Console.WriteLine($"\t\tImage \"{imageCount}\" Changed in file: {Path.GetFileName(filePath)}");
                   // replaced = true; 
                }
            }
        }// end of ProcessSection
        
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
