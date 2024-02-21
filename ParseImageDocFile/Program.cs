using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace WordAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            Application wordApp = new Application();
            Document doc = null;
            try
            {
                // Replace with your folder path
                string folderPath = @"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\Projects\PraseTestFolder";

                // Check if the folder exists
                if (Directory.Exists(folderPath))
                {
                    // Get all .doc files in the folder
                    string[] files = Directory.GetFiles(folderPath, "*.doc");

                    // Process each .doc file
                    foreach (string filePath in files)
                    {
                        Console.WriteLine($"\nProcessing File: {filePath}");

                        bool pictureFound = false; // Reset for each document
                        // Open the Word document
                        doc = wordApp.Documents.Open(filePath);
                        Console.WriteLine($"\tOpened File: {filePath}");

                        // List to hold shapes to delete
                        List<InlineShape> shapesToDelete = new List<InlineShape>();

                        // Iterate through all inline shapes in the document
                        foreach (InlineShape shape in doc.InlineShapes)
                        {
                            // Check if the shape is a picture
                            if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                            {
                                // Replace the picture with the new one
                                shape.Select();
                                // Add the new picture from file
                                InlineShape newShape = wordApp.Selection.InlineShapes.AddPicture(@"\\10.0.2.12\users\malghamgham\Desktop\My work - Maitham\GIF\Logo\logo.png");

                                // Add the shape to delete list
                                shapesToDelete.Add(shape);
                                Console.WriteLine($"\tImage Changed in file: {filePath}");
                                pictureFound = true;
                            }
                        }

                        if (!pictureFound)
                        {
                            Console.WriteLine($"\tNo Picture Found in doc {filePath}");
                        }

                        // Delete the old pictures after iterating through all shapes
                        foreach (InlineShape shapeToDelete in shapesToDelete)
                        {
                            shapeToDelete.Delete();
                        }

                        // Save the changes
                        doc.Save();
                        // Close the document
                        doc.Close();
                        Console.WriteLine($"\tClosed File: {filePath}");
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
                // Quit Word application
                wordApp?.Quit();
            }
        }
    }
}
