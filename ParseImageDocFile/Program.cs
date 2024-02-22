using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace WordAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            Application wordApp = new Application();
            Document doc = null;
            string imagePath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo.png";
            string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\TestFolder";
            // List to hold shapes to delete
            List<InlineShape> shapesToDelete = new List<InlineShape>();

            try
            {
                                                                                                                                            // Check if the folder exists
                if (Directory.Exists(folderPath))
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
                                                                                                                                           // Iterate through all inline shapes in the section
                                foreach (InlineShape shape in section.Range.InlineShapes)
                                {
                                                                                                                                           // Check if the shape is a picture
                                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                                    {
                                        shape.Select();
                                        shapesToDelete.Add(shape);                                                                         // Add the shape to delete list
                                        shape.Range.InlineShapes.AddPicture(imagePath);
                                        
                                        Console.WriteLine($"\t\tImage Changed in file: {Path.GetFileName(filePath)}");
                                    }
                                }
                            }
                            if (shapesToDelete.Count == 0)
                            {
                                Console.WriteLine($"\t\tNo Picture Found in doc {Path.GetFileName(filePath)}");
                            }
                            foreach (InlineShape shapeToDelete in shapesToDelete)
                            {
                                shapeToDelete.Delete();
                            }
                            doc.Save();                                                                                                    // Save the changes
                            doc.Close();                                                                                                   // Close the document
                            Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
                                                                                                                                           // Delete the old pictures after iterating through all shapes
                            
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error: processing file {Path.GetFileName(filePath)}: {ex.Message}");
                            doc.Save();
                            doc.Close();
                            Console.WriteLine($"\t\t\tClosed File: {Path.GetFileName(filePath)}");
                        }
                        finally
                        {
                                                                                                                                            // Release COM objects
                            if (doc != null)
                            {
                                Marshal.ReleaseComObject(doc);
                            }
                            //Marshal.ReleaseComObject(doc);
                        }
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
            finally
            {
                // Release COM objects
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }
    }
}
