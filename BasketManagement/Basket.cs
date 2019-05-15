using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace BasketManagement
{
    public class Basket
    {
        public List<Product> ProductList = new List<Product>();
        public List<string> tempList = new List<string>();
        public Basket()
        {
        }

        public Basket(List<Product> productList)
        {
            ProductList = productList;
        }

        public bool GiveDataToList()
        {
            var p1 = new Product(1, "Milk", 1.5, "Dairy Products");
            var p2 = new Product(2, "Eggs", 3.5, "Food");
            var p3 = new Product(3, "Tomatoes", 0.80, "Vegetables");
            var p4 = new Product(4, "shampoo", 5, "sanitary ware");
            var p5 = new Product(5, "chlorine", 3, "detergents");

            ProductList.Add(p1);
            ProductList.Add(p2);
            ProductList.Add(p3);
            ProductList.Add(p4);
            ProductList.Add(p5);

            return true;
        }
        public void SaveText()
        {
            using (TextWriter tw = new StreamWriter(@"C:\Users\Name\BasketToText.txt"))
            {
                foreach (var item in ProductList)
                {
                    tw.WriteLine(item.ToString());
                }
            }


        }

        public void SaveToaJason()
        {
            string jsonData = JsonConvert.SerializeObject(ProductList.ToArray());
            Console.WriteLine(jsonData);
            File.WriteAllText(@"C:\Users\Name\Jsondata.json", jsonData);
        }

        public void SaveToExcel()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Products"); //ISheet is an interface that creates a new sheet within an excel file.
            IRow r = sheet.CreateRow(0);               //IRow is an interface that creates a new row within a sheet.
            r.CreateCell(0).SetCellValue("Id");
            r.CreateCell(1).SetCellValue("Name");
            r.CreateCell(2).SetCellValue("Price");
            r.CreateCell(3).SetCellValue("Category");

            for (int i = 0; i < ProductList.Count; i++)
            {
                r = sheet.CreateRow(i + 1);
                r.CreateCell(0).SetCellValue(ProductList[i].Id);
                r.CreateCell(1).SetCellValue(ProductList[i].Name);
                r.CreateCell(2).SetCellValue(ProductList[i].Price);
                r.CreateCell(3).SetCellValue(ProductList[i].Category);

            }
            using (var fs = new FileStream(@"C:\Users\Name\studentData.xlsx", FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }

        }

        public void LoadfromText()
        {
            String path = @"C:\Users\Koron\source\TextLoad.txt";

            using (StreamReader sr = File.OpenText(path))
            {
                String temp = "";

                while ((temp = sr.ReadLine()) != null)
                {
                    tempList.Add(temp);
                }
            }

        }

        public void LoadfromJson()
        {

            string data = File.ReadAllText(@"C:\Users\Name\JsondataforInsert.json");
            var tempData = JsonConvert.DeserializeObject<List<Product>>(data);

            foreach (Product tempProduct in tempData)
            {
                ProductList.Add(tempProduct);
            }
        }
        public bool LoadfromExcel()
        {
            ProductList.Clear();
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(@"C:\Users\Name\ExcelToRead.xlsx", FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet("ProductList");
            for (int row = 1; row <= sheet.LastRowNum; row++) //zero line contains headers
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells

                {
                    int Id = int.Parse(sheet.GetRow(row).GetCell(0).ToString());
                    string Name = sheet.GetRow(row).GetCell(1).ToString();
                    double Price = double.Parse(sheet.GetRow(row).GetCell(2).ToString());
                    string Category = sheet.GetRow(row).GetCell(3).ToString();

                    Product temp = new Product(Id, Name, Price, Category);
                    ProductList.Add(temp);
                }
            }
            return true;
        }
    }
}
