using ExcelToXML.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Xml;
using System.Xml.Linq;
namespace ExcelToXML.Controllers;

public class FinmoFeesController : Controller
{
    [HttpGet]
    public IActionResult Convert()
    {
        var convertViewModel = new ConvertViewModel();
        return View(convertViewModel);
    }

    [HttpPost]
    public IActionResult Convert(IFormFile file, string fileName)
    {
        if (file == null || file.Length == 0)
        {
            ModelState.AddModelError("File", "No file uploaded.");
            var convertViewModel = new ConvertViewModel { 
                FileName = fileName,
                File = file
            };
            return View(convertViewModel);
        }

        // Process the file and convert it to XML
        var xmlDocument = ConvertExcelToXml(file);

        if (xmlDocument == null)
        {
            ModelState.AddModelError("File", "Error converting the file to XML.");
            var convertViewModel = new ConvertViewModel { FileName = fileName ,
                File =file
            };
            return View(convertViewModel);
        }

        // Save the XML file
        var outputPath = Path.Combine("C:\\Users\\acer\\Downloads", fileName + ".xml");
        xmlDocument.Save(outputPath);

        return RedirectToAction("Result", new { fileName });
    }

    public IActionResult Result(string fileName)
    {
        var filePath = Path.Combine("C:\\Users\\acer\\Downloads", fileName + ".xml");
        var fileBytes = System.IO.File.ReadAllBytes(filePath);
        return File(fileBytes, "application/xml", fileName + ".xml");
    }

    private XDocument ConvertExcelToXml(IFormFile file)
    {
        try
        {
            using (var stream = file.OpenReadStream())
            {
                var workbook = new XSSFWorkbook(stream);
                var worksheet = workbook.GetSheetAt(0); // Assuming the data is in the first worksheet

                var rows = worksheet.PhysicalNumberOfRows;
                var excelData = new ExcelData { Rows = new List<ExcelRows>() };

                for (int row = 1; row < rows; row++) // Start from row 1 (excluding header row)
                {
                    // Extract data from the Excel fil
                        var excelRow = new ExcelRows();

                        var dateCell = worksheet.GetRow(row).GetCell(8);
                        if (dateCell != null && dateCell.CellType == CellType.Numeric)
                        {
                            excelRow.Date = dateCell.DateCellValue;
                        }

                        var typeCell = worksheet.GetRow(row).GetCell(5);
                        if (typeCell != null && typeCell.CellType == CellType.String)
                        {
                            excelRow.Type = typeCell.StringCellValue;
                        }

                        var eventDescriptionCell = worksheet.GetRow(row).GetCell(2);
                        if (eventDescriptionCell != null && eventDescriptionCell.CellType == CellType.String)
                        {
                            excelRow.EventDescription = eventDescriptionCell.StringCellValue;
                        }

                        var amountCell = worksheet.GetRow(row).GetCell(7);
                        if (amountCell != null && amountCell.CellType == CellType.Numeric)
                        {
                            excelRow.Amount = (decimal)amountCell.NumericCellValue;
                        }

                        var currencyCell = worksheet.GetRow(row).GetCell(12);
                        if (currencyCell != null && currencyCell.CellType == CellType.String)
                        {
                            excelRow.Currency = currencyCell.StringCellValue;
                        }

                        var createdAtCell = worksheet.GetRow(row).GetCell(13);
                        if (createdAtCell != null && createdAtCell.CellType == CellType.Numeric)
                        {
                            excelRow.CreatedAt = createdAtCell.DateCellValue;
                        }

                 

                    excelData.Rows.Add(excelRow);
                }

                // Convert the ExcelData to XML format
                var xmlDocument = new XDocument(new XElement("Data",
                    excelData.Rows.Select(row => new XElement("Row",
                        new XElement("Date", row.Date),
                        new XElement("Type", row.Type),
                        new XElement("EventDescription", row.EventDescription),
                        new XElement("Amount", row.Amount),
                        new XElement("Currency", row.Currency),
                        new XElement("CreatedAt", row.CreatedAt)
                    ))));

                return xmlDocument;
            }
        }
        catch (Exception ex)
        {
            // Handle any exceptions that occur during conversion
            // You can log the error or handle it in a way that suits your needs
            return null;
        }
    }

}
//public class FinmoFeesController : ControllerBase
//{

//    [HttpPost]
//    public IActionResult ExcelToXML()
//    {


//        var excelFilePath = "C:\\Users\\acer\\Downloads\\Employee Sample Data.xlsx";

//        // Assuming you have the Excel file downloaded and its path is stored in the 'excelFilePath' variable

//        using (var fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
//        {
//            var workbook = new XSSFWorkbook(fileStream);
//            var worksheet = workbook.GetSheetAt(0); // Assuming the data is in the first worksheet

//            var rows = worksheet.PhysicalNumberOfRows;
//            var excelData = new ExcelData { Rows = new List<ExcelRows>() };

//            for (int row = 1; row < rows; row++) // Start from row 1 (excluding header row)
//            {
//                var excelRow = new ExcelRows();

//                var dateCell = worksheet.GetRow(row).GetCell(8);
//                if (dateCell != null && dateCell.CellType == CellType.Numeric)
//                {
//                    excelRow.Date = dateCell.DateCellValue;
//                }

//                var typeCell = worksheet.GetRow(row).GetCell(5);
//                if (typeCell != null && typeCell.CellType == CellType.String)
//                {
//                    excelRow.Type = typeCell.StringCellValue;
//                }

//                var eventDescriptionCell = worksheet.GetRow(row).GetCell(2);
//                if (eventDescriptionCell != null && eventDescriptionCell.CellType == CellType.String)
//                {
//                    excelRow.EventDescription = eventDescriptionCell.StringCellValue;
//                }

//                var amountCell = worksheet.GetRow(row).GetCell(7);
//                if (amountCell != null && amountCell.CellType == CellType.Numeric)
//                {
//                    excelRow.Amount = (decimal)amountCell.NumericCellValue;
//                }

//                var currencyCell = worksheet.GetRow(row).GetCell(12);
//                if (currencyCell != null && currencyCell.CellType == CellType.String)
//                {
//                    excelRow.Currency = currencyCell.StringCellValue;
//                }

//                var createdAtCell = worksheet.GetRow(row).GetCell(13);
//                if (createdAtCell != null && createdAtCell.CellType == CellType.Numeric)
//                {
//                    excelRow.CreatedAt = createdAtCell.DateCellValue;
//                }

//                excelData.Rows.Add(excelRow);
//            }

//            // Convert the ExcelData to XML format
//            var xmlDocument = new XDocument(new XElement("Data",
//                    excelData.Rows.Select(row => new XElement("Row",
//                        new XElement("Date", row.Date),
//                        new XElement("Type", row.Type),
//                        new XElement("EventDescription", row.EventDescription),
//                        new XElement("Amount", row.Amount),
//                        new XElement("Currency", row.Currency),
//                        new XElement("CreatedAt", row.CreatedAt)
//                    ))));

//                // Save the XML document to a file or use it for further processing in Tally
//                xmlDocument.Save("C:\\Users\\acer\\Downloads\\Output.xml");
//            }




//        return Ok();
//    }
//}

