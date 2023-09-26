using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"D:\WRK\Volgograd\vg3\test_vg3.xlsx";
            string outputFolder = @"D:\WRK\Volgograd\vg3";
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


            // Используем библиотеку ExcelDataReader для чтения данных из Excel файла
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Пропускаем заголовок
                    reader.Read();

                    // Создаем словарь для хранения данных
                    var data = new Dictionary<string, List<string[]>>();

                    while (reader.Read())
                    {
                        string valueInFirstColumn = reader.GetString(0);
                        string valueInSecondColumn = reader.GetString(1);
                        string valueInThirdColumn = reader.GetString(2);
                        string valueInFourthColumn = Convert.ToString(reader[3]);

                        // Если значение в первом столбце уже есть в словаре, добавляем значения в список
                        if (data.ContainsKey(valueInFirstColumn))
                        {
                            data[valueInFirstColumn].Add(new string[] { valueInSecondColumn, valueInThirdColumn, valueInFourthColumn });
                        }
                        // Если значение в первом столбце новое, создаем новую запись в словаре
                        else
                        {
                            data[valueInFirstColumn] = new List<string[]> { new string[] { valueInSecondColumn, valueInThirdColumn, valueInFourthColumn } };
                        }
                    }

                    // Создаем новые файлы на основе данных в словаре
                    foreach (var key in data.Keys)
                    {
                        string p_strPath = Path.Combine(outputFolder, $"{key}.xlsx");

                        using (var excel = new ExcelPackage())
                        {
                            var workSheet = excel.Workbook.Worksheets.Add(key);

                            workSheet.Cells[1, 1].Value = "Район";
                            workSheet.Cells[1, 2].Value = "Номер поля";
                            workSheet.Cells[1, 3].Value = "Площадь поля, га";
                            workSheet.Cells[1, 4].Value = "Данные из комментария";

                            int recordIndex = 2;
                           
                                foreach (var values in data[key])
                                {
                                workSheet.Cells[recordIndex, 1].Value = key;
                                workSheet.Cells[recordIndex, 2].Value = values[0];
                                workSheet.Cells[recordIndex, 3].Value = values[1];
                                workSheet.Cells[recordIndex, 4].Value = values[2];
                                recordIndex++;
                                }

                            workSheet.Column(1).AutoFit();
                            workSheet.Column(2).AutoFit();
                            workSheet.Column(3).AutoFit();
                            workSheet.Column(4).AutoFit();

                            if (File.Exists(p_strPath))
                                File.Delete(p_strPath);

                            // Create excel file on physical disk 
                            FileStream objFileStrm = File.Create(p_strPath);
                            objFileStrm.Close();

                            // Write content to excel file 
                            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
                            //Close Excel package
                            objFileStrm.Dispose();
                        }
                    }
                }
            }

            Console.WriteLine("Программа успешно выполнена!");
            Console.ReadLine();
        }
    }
}
