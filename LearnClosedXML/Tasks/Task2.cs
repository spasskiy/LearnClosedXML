using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LearnClosedXML
{
    internal partial class Program
    {
        static void Task2()
        {
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
