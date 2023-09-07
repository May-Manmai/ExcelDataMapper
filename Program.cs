using System;
using System.IO;
using OfficeOpenXml;

public class MyDataFormat
{
    public string? deliveryStoreNumber { get; set; }
    public string? customerReference { get; set; }
    public string? customerName { get; set; }
    public string? deliveryFormattedAddress { get; set; }

    // Add other properties as needed to match your desired format

}

class Program
{
    static void Main(string[] args)
    {
        string excelFilePath = "/Users/tangmay/Documents/1. SOLBOX/Import spreadsheets/Drillcut/testExcelApp.xlsx";

        if (File.Exists(excelFilePath))
        {
            ImportExcelData(excelFilePath);
        }
        else
        {
            Console.WriteLine("The specified Excel file does not exist.");
        }
    }

    static void ImportExcelData(string filePath)
    {
        // Your Excel data import code here
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    Console.Write($"{worksheet.Cells[row, col].Text}\t");
                }
                Console.WriteLine();
            }
        }
    }

    static List<MyDataFormat> MapExcelDataToMyFormat(ExcelWorksheet worksheet)
    {
        List<MyDataFormat> mappedData = new List<MyDataFormat>();
        for (int row = 2; row <= worksheet.Dimension.Rows; row++)
        {
            MyDataFormat mappedItem = new MyDataFormat();

            mappedItem.deliveryStoreNumber = worksheet.Cells[row, 1].Text;
            mappedItem.customerReference = worksheet.Cells[row, 1].Text;
            mappedItem.customerName = worksheet.Cells[row, 1].Text;
            mappedItem.deliveryFormattedAddress = worksheet.Cells[row, 1].Text;

            mappedData.Add(mappedItem);

        }
        return mappedData;
    }
}
