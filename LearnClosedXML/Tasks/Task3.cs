using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LearnClosedXML
{
    internal partial class Program
    {
        static void Task3()
        {
            //-----------------------------------------------------------
            //Задача №3 : Сделать тект жирным и добавить фон
            //-----------------------------------------------------------

            // Открываем существующий файл
            var workbook3 = new XLWorkbook("output.xlsx");
            //Получаем первый лист
            var worksheet3 = workbook3.Worksheet(1);

            for (int i = 1; i < 3; i++)
            {
                worksheet3.Cell(1, i).Style.Font.SetBold(true);
                worksheet3.Cell(1, i).Style.Font.FontColor = XLColor.White;
                worksheet3.Cell(1, i).Style.Font.FontSize = 16;

                //С выравниванием тут похоже намудрили
                worksheet3.Cell(1, i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet3.Cell(1, i).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                worksheet3.Cell(1, i).Style.Fill.BackgroundColor = XLColor.Blue;
            }
            workbook3.Save();
        }
    }
}
