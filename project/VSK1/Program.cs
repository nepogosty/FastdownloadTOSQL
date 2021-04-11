using OfficeOpenXml;
using System;
using System.IO;
using RandomNameGeneratorLibrary;
using System.Data;
using Microsoft.Data.SqlClient;


namespace VSK1
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //Наполнение данных 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Random rng = new Random();
            PersonNameGenerator nameGenerator = new PersonNameGenerator();     
            PlaceNameGenerator placeNameGenerator = new PlaceNameGenerator();  

            //1 часть 
            string filepath =createexcel(rng, nameGenerator, placeNameGenerator); //создает xlsx-файл в пути bin\debug\netcoreapp3.1
            
            //2 часть

            var dt = DateTime.Now; //Счетчик выполнения программы
            WritetoDB(filepath); //записывает данные из созданной ранее таблицы в mssql(часть 2).
            var diff = DateTime.Now - dt;
            Console.WriteLine("Время выполнения 2-ой части:"+ diff.ToString()); //Вывод времени

        }
        public static string createexcel(Random rng, PersonNameGenerator nameGenerator, PlaceNameGenerator placeNameGenerator)
        {
            string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, Path.GetRandomFileName() + ".xlsx"); //путь текущей директории
            
            //создание excel-файла
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Information"); //Название листа
                worksheet.Cells[1, 1].Value = "Id"; //Название колонок
                worksheet.Cells[1, 2].Value = "Фамилия";
                worksheet.Cells[1, 3].Value = "Имя";
                worksheet.Cells[1, 4].Value = "Отчество";
                worksheet.Cells[1, 5].Value = "Телефон";
                worksheet.Cells[1, 6].Value = "Адрес";
                worksheet.Column(1).Width = 20; //Размер ячеек
                worksheet.Column(2).Width = 20;
                worksheet.Column(3).Width = 20;
                worksheet.Column(4).Width = 20;
                worksheet.Column(5).Width = 20;
                worksheet.Column(6).Width = 20;
                for (int i = 2; i < 200002; i++)
                {
                    worksheet.Cells[i, 1].Value = i-1; //наполнение данных 
                    worksheet.Cells[i, 2].Value = nameGenerator.GenerateRandomLastName();
                    worksheet.Cells[i, 3].Value = nameGenerator.GenerateRandomFirstName();
                    worksheet.Cells[i, 4].Value = nameGenerator.GenerateRandomFirstName();
                    worksheet.Cells[i, 5].Value = "+79" + Convert.ToString(rng.Next(100000000, 999999999));
                    worksheet.Cells[i, 6].Value = placeNameGenerator.GenerateRandomPlaceName();
                }
                excelPackage.Save(); //сохранение 
            }
            return (filePath);
        }
        public static void WritetoDB(string filepath1)
        {
            string filepath = filepath1;
            //Перемещение данных в datatable 
            FileInfo file = new FileInfo(filepath);
            ExcelPackage excelPackage = new ExcelPackage(file);
            DataTable dt = new DataTable();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];         
            int currentColumn = 1;
            // зацикливаем все столбцы на листе и добавляем их в таблицу данных
            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();
                dt.Columns.Add(columnName);
                currentColumn++;
            }
            // начинаем добавлять содержимое файла Excel в таблицу данных
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(newRow);
            }
            // подключение к БД
            string consString = String.Format("Server=(localdb)\\mssqllocaldb;Database=vskstrahovanie;Trusted_Connection=True;");
            using (SqlConnection con = new SqlConnection(consString))
            {
                //массивное копирование с помощью BulkCoy
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    sqlBulkCopy.DestinationTableName = "dbo.Persons";
                    con.Open();
                    sqlBulkCopy.WriteToServer(dt);
                    con.Close();
                }
            }

        }
    }
    //Класс, соответствующий БД 
    class Person
    {
        public int Id { get; set; }
        public string Surname { get; set; }
        public string Name { get; set; }
        public string Middlename { get; set; }
        public string Phone{ get; set; }
        public string Adress { get; set; }

    }
}
