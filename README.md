# PixActivities_Activities.ExcelPassword
Активность зашифровывает указанный файл excel паролем

string pathFile;

// Объявление приложения
Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();

//Добавить рабочую книгу
Microsoft.Office.Interop.Excel.Workbook workBook = appExcel.Workbooks.Open(pathFile);

workBook.Password = securePassword;
appExcel.DisplayAlerts = false;          

//Закрыть книгу с сохранением
workBook.Save();
workBook.Close();

// Закрыть приложение
appExcel.Quit();