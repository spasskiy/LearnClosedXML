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
        static void Task1()
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
            //Внимание! Тут можно влететь в эксепшен если файл будет уже открыт в чём то и к нему не будет доступа
            //так что лучше всё взаимодействие с файлом в рабочих проектах стоит оборачивать в try/catch
            workbook.SaveAs("output.xlsx");
        }
    }
}
