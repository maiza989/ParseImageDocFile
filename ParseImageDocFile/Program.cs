/*using System.Runtime.InteropServices;
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
 *                                                                          --XML Attempt-- 
 *                                                                          
 *  Do not work since I am working with .doc files.                                                                          
 */

/*using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Document = Aspose.Words.Document;

namespace WordAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";
            string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo2.jpg";

            string folderDOCPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder\.doc";
            string folderDOCXpath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder\.docx";


           // ConvertDocToDocx(folderDOCPath, folderDOCXpath);
            ProcessDocuments(folderPath, imagePath);
        }

        static void ConvertDocToDocx(string docFilePath, string docxFilePath)
        {
            // Load the .doc file
            Document doc = new Document(docFilePath);

            // Save as .docx
            doc.Save(docxFilePath, SaveFormat.Docx);
        }

        static void ProcessDocuments(string folderPath, string imagePath)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("Folder not found.");
                return;
            }

            string[] files = Directory.GetFiles(folderPath, "*.doc");
            foreach (string filePath in files)
            {
                Console.WriteLine($"\nProcessing File: {Path.GetFileName(filePath)}");
                try
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                    {
                        IEnumerable<ImagePart> imageParts = doc.MainDocumentPart.ImageParts.ToList();
                        int imageCount = 0;
                        foreach (ImagePart imagePart in imageParts)
                        {
                            try
                            {
                                ReplaceImage(imagePart, imagePath);
                                imageCount++;
                                Console.WriteLine($"\tImage \"{imageCount}\" Changed in file: {Path.GetFileName(filePath)}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error replacing image in file {Path.GetFileName(filePath)}: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: processing file {Path.GetFileName(filePath)}: {ex.Message}");
                    Console.WriteLine($"Skipping the file: {Path.GetFileName(filePath)}");
                }

            }
        }

        static void ReplaceImage(ImagePart imagePart, string imagePath)
        {
            using (Stream imageStream = File.Open(imagePath, FileMode.Open))
            {
                imagePart.FeedData(imageStream);
            }
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
using System.Collections.Generic;
using System.Drawing;

namespace WordAutomation
{
    class Program
    {
        private string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo2.jpg";
        private string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";
        List<DocumentObject> objectsToRemove;
        List<DocPicture> pictures = new List<DocPicture>();


        public Program() 
        {
            objectsToRemove = new List<DocumentObject>();
        }


        public void Run()
        {
            try
            {
                if (Directory.Exists(folderPath))
                {
                    string[] files = Directory.GetFiles(folderPath, "*.doc");
                    foreach (string filePath in files)
                    {
                        ProcessDocument(filePath);
                    }
                }
                else
                {
                    Console.WriteLine("Folder not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        private void ProcessDocument(string filePath)
        {
            Console.WriteLine($"\nProcessing File: {Path.GetFileName(filePath)}");
            try
            {
                Document doc = new Document();
                doc.LoadFromFile(filePath);

                foreach (Section section in doc.Sections)
                {
                    ProcessSection(section, filePath);
                }
                if(objectsToRemove.Count == 0)
                {
                    Console.WriteLine($"\t\tNo Picture Found in doc {Path.GetFileName(filePath)}");
                }
                // Remove the objects after iterating through all paragraphs
                foreach (var objToRemove in objectsToRemove)
                {
                    // Remove the object from its parent collection
                    objToRemove.Owner.ChildObjects.Remove(objToRemove);
                }
                doc.SaveToFile(filePath, FileFormat.Doc);

                Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: processing file {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }

        private void ProcessSection(Section section, string filePath)
        {
            try
            {
                int imageCount = 0;
                
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    foreach (DocumentObject docObj in paragraph.ChildObjects)
                    {
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            imageCount++; // Increment image count if a picture is found

                            // Create a new DocPicture with the desired image
                            DocPicture newPicture = docObj as DocPicture;                         
                            newPicture.LoadImage(Image.FromFile(imagePath));
                            Console.WriteLine($"\t\tChange Image \"{imageCount}\" in file: {Path.GetFileName(filePath)}");
 
                        }
                    }
                }
                
                Console.WriteLine($"Total images found in {Path.GetFileName(filePath)}: {imageCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error looking for a picture in file {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }

        static void Main(string[] args)
        {
            Program program = new Program();

            program.Run();
        }
    }
}



