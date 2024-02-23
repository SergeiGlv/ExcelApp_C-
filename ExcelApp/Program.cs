using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Пример");
          
            var rngTable = worksheet.Range("A1:G" + 10);
            rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Columns().AdjustToContents();
       
            worksheet.Cell("B2").Value = "Транзакции";

            worksheet.Cell("B3").Value = "Имя";
            worksheet.Cell("B4").Value = "Иван";
            worksheet.Cell("B5").Value = "Петр";
            worksheet.Cell("B6").SetValue("Илья"); 

            worksheet.Cell("C3").Value = "Фамилия";
            worksheet.Cell("C4").Value = "Иванов";
            worksheet.Cell("C5").Value = "Петров";
            worksheet.Cell("C6").SetValue("Сидоров"); 

            worksheet.Cell("D3").Value = "В базе";
            worksheet.Cell("D4").Value = true;
            worksheet.Cell("D5").Value = false;
            worksheet.Cell("D6").SetValue(false); 

            
            worksheet.Cell("E3").Value = "Дата";
            worksheet.Cell("E4").Value = new DateTime(1919, 1, 21);
            worksheet.Cell("E5").Value = new DateTime(1907, 3, 4);
            worksheet.Cell("E6").SetValue(new DateTime(1921, 12, 15)); 

         
            worksheet.Cell("F3").Value = "Приход";
            worksheet.Cell("F4").Value = 2000;
            worksheet.Cell("F5").Value = 40000;
            worksheet.Cell("F6").SetValue(10000); 
            worksheet.Cell("D4").Value = true;
            worksheet.Cell("D5").Value = false;
            worksheet.Cell("D6").SetValue(false); 

         

            rngTable = worksheet.Range("B2:F6");
                     
            var rngDates = rngTable.Range("D3:D5"); 
            var rngNumbers = rngTable.Range("F4:F6"); 
            rngDates.Style.NumberFormat.NumberFormatId = 15;

            rngNumbers.Style.NumberFormat.Format = "$ #,##0";
            rngTable.FirstCell().Style
            .Font.SetBold()
            .Fill.SetBackgroundColor(XLColor.CornflowerBlue)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            rngTable.FirstRow().Merge(); 
            var rngData = worksheet.Range("B3:F6");
            var excelTable = rngData.CreateTable();

            excelTable.ShowTotalsRow = true;
   
            excelTable.Field("Приход").TotalsRowFunction = XLTotalsRowFunction.Average;
            excelTable.Field("Дата").TotalsRowLabel = "Среднее:";
            worksheet.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs("Example.xlsx");

        }
    }
}
