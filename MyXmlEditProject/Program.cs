using System;
using System.Linq;
using System.Xml.Linq;
using OfficeOpenXml;


class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set to NonCommercial if appropriate

        // Specify the path to your XML file
        string xmlFilePath = "/Users/connieleighton/Documents/XMLEdit/OriginalXML.xml";
        string spreadsheetPath = "/Users/connieleighton/Downloads/XMLConvert.xlsx";

        using (var packageA = new ExcelPackage(new FileInfo(spreadsheetPath)))
        {

            var Sheet1 = packageA.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "Sheet1");
            int Sheet1Rows = Sheet1.Dimension.Rows;

            for (int rowA = 2; rowA <= Sheet1Rows; rowA++) 
            {
                string careRecipientID = Sheet1.Cells[rowA, 3].Value?.ToString();

                try
                {
                    // Load the XML document from the file
                    XDocument xmlDoc = XDocument.Load(xmlFilePath);

                    // Find and remove the careRecipientPayment node where careRecipientID is "0414835747"
                    XElement nodeToRemove = xmlDoc.Descendants("careRecipientPayment")
                        .FirstOrDefault(node => node.Element("careRecipientDetails")?.Element("careRecipientID")?.Value == ("0" + careRecipientID));

                    if (nodeToRemove != null)
                    {
                        nodeToRemove.Remove();

                        // Save the modified XML document back to the file
                        xmlDoc.Save(xmlFilePath);

                        Console.WriteLine("Node removed successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Node not found.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }

                 try
                {
                    // Load the XML document from the file
                    XDocument xmlDoc = XDocument.Load(xmlFilePath);

                    // Find and remove the careRecipientPayment node where careRecipientID is "0414835747"
                    XElement nodeToRemove = xmlDoc.Descendants("careRecipientPayment")
                        .FirstOrDefault(node => node.Element("careRecipientID")?.Value == ("0" + careRecipientID));

                    if (nodeToRemove != null)
                    {
                        nodeToRemove.Remove();

                        // Save the modified XML document back to the file
                        xmlDoc.Save(xmlFilePath);

                        Console.WriteLine("Node removed successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Node not found.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }
    }
}
