using ClosedXML;
using ClosedXML.Excel;

namespace LearnClosedXML
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Создаем новую рабочую книгу
            var workbook = new XLWorkbook();

            // Добавляем новый лист
            var worksheet = workbook.Worksheets.Add("First Sheet");

            // Заполняем ячейки данными
            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Age";

            worksheet.Cell(2, 1).Value = "Bill";
            worksheet.Cell(2, 2).Value = 54;

            //Сохраняем файл
            workbook.SaveAs("output.xlsx");
        }
    }
}
