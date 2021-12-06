using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Channels;
using System.Threading.Tasks;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace CardCheck
{
    class Program
    {
        static async Task Main(string[] args)
        {
            //TODO: Добавить проверку вводимых данных.
            int CardBIN;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Card> Cards = new List<Card>();
            using (var package = new ExcelPackage("bins.xlsx"))
            {
                    var Sheet = package.Workbook.Worksheets["Grid Results"];
                    int LastRow = Sheet.Dimension.End.Row;
                    while (Sheet.Cells[LastRow, 1].Value == null)
                    {
                        LastRow--;
                    }

                    #region Read from xlsx

                    for (int i = 2; i <= LastRow; i++)
                    {
                        Cards.Add(new Card
                        {
                            ID = int.Parse(Sheet.Cells[i, 1].Value.ToString()),
                            BIN = int.Parse(Sheet.Cells[i, 2].Value.ToString()),
                            Brand = Sheet.Cells[i, 3].Value.ToString(),
                            Bank = Sheet.Cells[i, 4].Value.ToString(),
                            BINType = Sheet.Cells[i, 5].Value.ToString(),
                            BINLevel = Sheet.Cells[i, 6].Value.ToString(),
                            Country = Sheet.Cells[i, 7].Value.ToString(),
                            Withdrawal = true,
                            Put = true
                        });
                    }

                    #endregion

                    Console.WriteLine("Здравствуйте. Навигация в меню осуществляется ввода цифр, обозначающих пункты меню");
                    Console.WriteLine("1. Проверка карты по BIN \n "+
                                      "2. Создание правила");
                    switch (Console.ReadLine())
                    {
                        case "1":
                            Console.WriteLine("Введите BIN карты");
                            CardBIN= int.Parse(Console.ReadLine());
                            Card FindCard = Cards.Find((x) => x.BIN == CardBIN);
                           
                            #region Check Operations
                            if (FindCard.Withdrawal == true)
                            {
                                Console.WriteLine($"Списание с карты {FindCard.BIN} доступно");
                            }
                            else
                            {
                                Console.WriteLine($"Списание с карты {FindCard.BIN} недоступно \n Причина :");
                            }
                            if (FindCard.Put == true)
                            {
                                Console.WriteLine($"Пополнение карты {FindCard.BIN} доступно");
                            }
                            else
                            {
                                Console.WriteLine($"Пополнение карты {FindCard.BIN} недоступно \n Причина :");
                            }
                            #endregion
                            break;
                        case "2":
                            //TODO: Расширить возможности создания правил
                            Console.WriteLine("Введите BIN карты для создания правила");
                            CardBIN = int.Parse(Console.ReadLine()); 
                            Card FixCard = Cards.Find((x) => x.BIN == CardBIN);
                            Console.WriteLine("Карта найдена. Выберите операцию\n" +
                                              "1. Разрешить все операции\n" +
                                              "2. Запретить все операции\n" +
                                              "3. Разрешить только пополнение\n" +
                                              "4. Разрешить только снятие");
                            switch (Console.ReadLine())
                            {
                                case "1":
                                    FixCard.Withdrawal = true;
                                    FixCard.Put = true;
                                    break;
                                case "2":
                                    FixCard.Withdrawal = false;
                                    FixCard.Put = false;
                                    break;
                                case "3":
                                    FixCard.Withdrawal = false;
                                    FixCard.Put = true;
                                    break;
                                case "4":
                                    FixCard.Withdrawal = true;
                                    FixCard.Put = false;
                                        break;
                            }
                            Console.WriteLine("Правило добавлено");
                            break;
                    }
                       using (FileStream fs = new FileStream("result.json", FileMode.OpenOrCreate))
                    {
                        await JsonSerializer.SerializeAsync(fs, Cards);
                        Console.WriteLine($"Данные по картам  сохранены");
                    }
                    Console.ReadLine();
            }
            

        }
    }
}