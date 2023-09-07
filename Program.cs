using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

public class MyDataFormat
{
    public string? deliveryStoreNumber { get; set; }
    public string? customerReference { get; set; }
    public string? customerName { get; set; }
    public string? deliveryFormattedAddress { get; set; }
}

class Program
{
    static void Main(string[] args)
    {
        string excelFilePath = "/Users/tangmay/Documents/1. SOLBOX/Test files/testExcelApp.xlsx";

        if (File.Exists(excelFilePath))
        {
            List<MyDataFormat> data = ImportExcelData(excelFilePath);

            // Now you can work with the 'data' list, which contains the mapped Excel data.
            foreach (var item in data)
            {
                Console.WriteLine($"Delivery Store Number: {item.deliveryStoreNumber}");
                Console.WriteLine($"Customer Reference: {item.customerReference}");
                Console.WriteLine($"Customer Name: {item.customerName}");
                Console.WriteLine($"Delivery Formatted Address: {item.deliveryFormattedAddress}");
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("The specified Excel file does not exist.");
        }
    }

    static List<MyDataFormat> ImportExcelData(string filePath)
    {
        List<MyDataFormat> mappedData = new List<MyDataFormat>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            // Skip the header row and start from row 2
            for (int row = 2; row <= rowCount; row++)
            {
                MyDataFormat mappedItem = new MyDataFormat();

                mappedItem.deliveryStoreNumber = worksheet.Cells[row, 1].Text;
                mappedItem.customerReference = worksheet.Cells[row, 2].Text;
                mappedItem.customerName = worksheet.Cells[row, 3].Text;
                mappedItem.deliveryFormattedAddress = worksheet.Cells[row, 4].Text;

                mappedData.Add(mappedItem);
            }
        }

        return mappedData;
    }
}
