using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace WordAutomation
{
    class Program
    {
        Application wordApp = new Application();
        Document doc = null;
        string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo.png";
        string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";
        List<InlineShape> shapesToDelete = new List<InlineShape>();                                                                          // List to hold shapes to delete

        private void ParseDocFile()
        {
            try
            {

                if (Directory.Exists(folderPath))                                                                                           // Check if the folder exists
                {
                    string[] files = Directory.GetFiles(folderPath, "*.doc");                                                               // Get all .doc files in the folder
                                                                                                                                            // Process each .doc file
                    foreach (string filePath in files)
                    {
                        Console.WriteLine($"\nProcessing File: {Path.GetFileName(filePath)}");
                        try
                        {
                            doc = wordApp.Documents.Open(filePath);                                                                         // Open the Word document
                            Console.WriteLine($"\tOpened File: {Path.GetFileName(filePath)}");
                            shapesToDelete.Clear();                                                                                         // Reset shapes to delete for each document

                            foreach (Section section in doc.Sections)
                            {
                                Range range = section.Range;

                                foreach (InlineShape shape in section.Range.InlineShapes)                                                   // Iterate through all inline shapes in the section
                                {

                                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture)                                               // Check if the shape is a picture
                                    {
                                        shape.Select();
                                        shapesToDelete.Add(shape);                                                                          // Add the shape to delete list
                                        shape.Range.InlineShapes.AddPicture(imagePath);                                                     // Add new image
                                        Console.WriteLine($"\t\tImage Changed in file: {Path.GetFileName(filePath)}");
                                    }
                                }// end of inlineshape foreacy loop 
                            }// end of section foreach loop

                            if (shapesToDelete.Count == 0)
                            {
                                Console.WriteLine($"\t\tNo Picture Found in doc {Path.GetFileName(filePath)}");
                            }
                            foreach (InlineShape shapeToDelete in shapesToDelete)
                            {
                                shapeToDelete.Delete();                                                                                    // Delete the old pictures after iterating through all shapes
                            }
                            doc.Save();                                                                                                    // Save the changes
                            doc.Close();                                                                                                   // Close the document
                            Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");


                        }// end of inner try 
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error: processing file {Path.GetFileName(filePath)}: {ex.Message}");
                            doc.Save();
                            doc.Close();
                            Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
                        }// end of catch 
                        finally
                        {
                            if (doc != null)
                            {
                                Marshal.ReleaseComObject(doc);                                                                              // Release COM objects
                            }
                        }// end of finally
                    }// end of foreach loop itreating through files
                }// end of if statement to check folder exist
                else
                {
                    Console.WriteLine("Folder not found.");
                }
            }// end of outer try
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);                                                                                      // Release COM objects
                }
            }
        }// end of method 
        static void Main(string[] args)
        {
            Program program = new Program();
            program.ParseDocFile();
        }
    }
}
