﻿using ClosedXML;
using ClosedXML.Excel;

namespace LearnClosedXML
{
    /*
     *  Решение учебное. В боевом проекте не стоит писать настолько громоздкие методы.
     *  Тут так сделано для удобства восприятия.
     */
    internal partial class Program
    {
        static void Main(string[] args)
        {
            Task1(); //Задача №1 : Создать файл с нашими данными
            Task2(); //Задача №2 : Считать данные из файла
            Task3(); //Задача №3 : Сделать тект жирным и добавить фон
            Task4(); //Задача №4 : Добавить в существующий файл новые столбцы и заполнить файл новыми данными
        }
    }
}
