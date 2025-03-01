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
        static void Task4()
        {
            //------------------------------------------------------------------------
            //Задача №4 : Добавить в существующий файл новые столбцы
            //и заполнить файл новыми данными
            //------------------------------------------------------------------------
            var workbook4 = new XLWorkbook("output.xlsx");
            var worksheet4 = workbook4.Worksheet(1);

            //Список всех столбцов которые должны присутствовать на листе
            var currentHeaders = new List<string>() { "Name", "Age", "Address", "Status", "Salary" };

            //Получаем существующие заголовки
            var existingHeaders = new List<string>();
            foreach (var cell in worksheet4.Row(1).Cells())
            {
                existingHeaders.Add(cell.GetString().Trim());
            }
            //Из всех что должны быть, вычитаем те что уже есть
            var missingHeaders = currentHeaders.Except(existingHeaders).ToList();

            // Если есть недостающие заголовки, добавить их в конец
            if (missingHeaders.Any())
            {
                int lastColumnIndex = worksheet4.Row(1).CellsUsed().Count(); // Последний столбец в первой строке

                for (int i = 0; i < missingHeaders.Count; i++)
                {

                    // Установить значение ячейки
                    var newCell = worksheet4.Cell(1, lastColumnIndex + i + 1);
                    newCell.Value = missingHeaders[i];

                    // Скопировать форматирование из ячейки (1, 1) НО! Ширина тоже будет скопирована
                    newCell.Style = worksheet4.Cell(1, 1).Style;

                    // Задаем ширину столбца под размер содержимого ячейки
                    /*
                                Тут следующая заморочка. Если в столбце уже что-то есть, то растянется под размер
                                максимального содержимого. Если же это пустой столбец, то как раз получим растяжение по ширине
                                нашего добовляемого заголовка. Если нужно, то у Column есть свойство Width, так что при желании
                                можно задать его точное значение руками. Но какой-то адекватной формулы для расчёта ширины под
                                длину строки я не нашёл. Задаётся примерно на глаз.
                             */
                    worksheet4.Column(lastColumnIndex + i + 1).AdjustToContents();
                }
            }

            //Теперь заполним значением NULL все пустые поля в непустых строчках
            //Для начала найдём номер последней не пустой строчки

            // Укажите номер столбца, для которого нужно найти последнюю заполненную строку
            int columnNumber = 1; // Например, столбец A

            // Найдем последнюю заполненную ячейку в указанном столбце
            var lastCell = worksheet4.Column(columnNumber).LastCellUsed();

            // Получаем номер строки последней заполненной ячейки
            int lastRowIndexInColumn = lastCell.Address.RowNumber;

            //Т.к. в первой строке у нас заголовки, то идём начиная со второй строки и до последней заполненной
            for (int rowIndex = 2; rowIndex <= lastRowIndexInColumn; rowIndex++)
            {
                //По строке идём от первой ячейки до последнего столбца над которым есть заголовок
                for (int colIndex = 1; colIndex <= currentHeaders.Count; colIndex++)
                {
                    var cell = worksheet4.Cell(rowIndex, colIndex);
                    //Если в ячейке пусто
                    if (string.IsNullOrWhiteSpace(cell.GetString()))
                    {
                        //Вписываем заполнитель
                        cell.Value = "NULL";
                    }
                }
            }

            //Теперь добавим к документу сгенерированные данные
            //Генерируем данные
            List<User> users = new ContentGenerator().GenerateUsers();

            //Получаем индексы полей. Вдруг они перемешаны в нашем файле
            int nameIndex = -1;
            int ageIndex = -1;
            int addressIndex = -1;
            int statusIndex = -1;
            int salaryIndex = -1;

            // Перебираем все столбцы в первой строке
            for (int colIndex = 1; colIndex <= worksheet4.ColumnCount(); colIndex++)
            {
                // Получаем значение ячейки в первой строке и текущем столбце
                string headerValue = worksheet4.Cell(1, colIndex).GetString();

                // Используем switch для определения типа заголовка
                switch (headerValue)
                {
                    case "Name":
                        nameIndex = colIndex;
                        break;

                    case "Age":
                        ageIndex = colIndex;
                        break;

                    case "Address":
                        addressIndex = colIndex;
                        break;

                    case "Status":
                        statusIndex = colIndex;
                        break;

                    case "Salary":
                        salaryIndex = colIndex;
                        break;
                }
            }

            //Получаем индекс первой пустой строки. Тут мы уверены, что как минимум строка заголовка уже существует
            int currentRowIndex = worksheet4.LastRowUsed().RowNumber() + 1;

            foreach (var user in users)
            {
                //Заполняем строку
                worksheet4.Cell(currentRowIndex, nameIndex).Value = user.Name;
                worksheet4.Cell(currentRowIndex, ageIndex).Value = user.Age;
                worksheet4.Cell(currentRowIndex, addressIndex).Value = user.Address;

                //С bool внимание на локализацию! Может и на русском вписать, так что лучше в явном виде указать что вписывать
                worksheet4.Cell(currentRowIndex, statusIndex).Value = user.Status ? "true" : "false";
                worksheet4.Cell(currentRowIndex, salaryIndex).Value = user.Salary;

                currentRowIndex++;
            }

            workbook4.Save();
        }
    }
}
