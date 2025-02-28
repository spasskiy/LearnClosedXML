using ClosedXML;
using ClosedXML.Excel;

namespace LearnClosedXML
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //------------------------------------------------------
            //Задача №1 : Создать файл с нашими данными
            //------------------------------------------------------

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

            //---------------------------------------------
            //Задача №2 : Считать данные из файла
            //---------------------------------------------

            // Открываем существующий файл
            var workbook2 = new XLWorkbook("output.xlsx");
            //Получаем первый лист
            var worksheet2 = workbook2.Worksheet(1);

            //Считываем нужные нам значения
            string name = worksheet2.Cell(2, 1).Value.ToString();

            //Если уверены в данных, то можно конечно сразу считать значение из ячейки, но через Try это более безопасный вариант
            int age = 0;
            if (worksheet2.Cell(2, 2).TryGetValue<int>(out int ageValue))
                age = ageValue;
            else
                throw new Exception("Wrong data in Cell");
            //int age = worksheet2.Cell(2, 2).GetValue<int>();

            Console.WriteLine($"Name : {name}, Age : {age}");
        }
    }
}
