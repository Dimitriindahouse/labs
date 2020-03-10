

using System;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Program pro = new Program();
       Console.WriteLine("Выберите любое целое число от 1 до 3");

 string n = Console.ReadLine();
            if ("1" == n)
               pro.AddWord(@"C:\Users\ZJack\source\repos\ConsoleApp4\1.jpg");
         else if ("2" == n)
                           pro.AddWord(@"C:\Users\ZJack\source\repos\ConsoleApp4\2.jpg");
                         else if ("3" == n)
                               pro.AddWord(@"C:\Users\ZJack\source\repos\ConsoleApp4\3.jpg");
                          else
                  Console.WriteLine("Не верный ввод!");
            
              Console.WriteLine("Нажмите любую клавишу для продолжения...");
                          Console.ReadKey();
            
              Console.WriteLine();
                          Console.WriteLine("Введите число для EXCEL: ");
                          pro.AddExcel(Console.ReadLine());
                          Console.WriteLine("Нажмите любую клавишу для выхода...");
            
              Console.ReadKey();
                      }
	
	        public void AddExcel(string number)
	        {
	            Microsoft.Office.Interop.Excel.Workbook workbook = null;
	            
	            
	            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
	            app.Visible = false;
	
	            string source = @"C:\Users\ZJack\source\repos\ConsoleApp4\excel.xlsx";
	
	            workbook = app.Workbooks.Open(source);
	
	            app.Cells.NumberFormat = "0";
	            app.Cells[1, 1] = number;
	            app.Visible = true;
	
	
	
	        }
	
	        public void AddWord(string picture)
	        {
	            //Создаем объект документа 
	            Microsoft.Office.Interop.Word.Document doc = null;
	            try
	            {
	                //Создаем объект приложения 
	                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
	                app.Visible = false;
	
	                //Путь до шаблона документа
	                string source = @"C:\Users\ZJack\source\repos\ConsoleApp4\word.docx";
	
	                //Открываем 
	                doc = app.Documents.Open(source);
	                doc.Activate();                
	
	                //Добавляем картинку
	                Microsoft.Office.Interop.Word.Range range;
	                range = doc.Content;
	                range.InlineShapes.AddPicture(picture);
	                app.Visible = true;
	                   
	            }catch (Exception ex)
	            {
	                doc.Close();
	                doc = null;
	                Console.WriteLine("Во время выполнения произошла ошибка!");
	                Console.ReadLine();
	            }



    }
}
}
