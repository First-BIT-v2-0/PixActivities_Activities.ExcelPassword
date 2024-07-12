using System;
using System.Security;
using Activities.ExcelPassword.Properties;
using BR.Core;
using BR.Core.Attributes;
using Microsoft.Office.Interop.Excel;

namespace Activities.ExcelPassword
{
    [LocalizableScreenName("ExcelPassword_ScreenName", typeof(Resources))] // Имя активности, отображаемое в списке активностей и в заголовке шага
    [LocalizablePath("PathActivities", typeof(Resources))] // Путь к активности в панели "Активности"
    [LocalizableDescription("Activities_Description", typeof(Resources))] // описание активности

    [Image(typeof(ExcelPassword), "Activities.ExcelPassword.password.png")] //Иконка активности

    public class ExcelPassword : Activity
    {
        [LocalizableScreenName("PathFile_ScreenName", typeof(Resources))]
        [LocalizableDescription("PathFile_Description", typeof(Resources))]
        [IsRequired]
        [IsFilePathChooser]
        public string pathFile { get; set; }

        [LocalizableScreenName("Password_ScreenName", typeof(Resources))]
        [LocalizableDescription("Password_Description", typeof(Resources))]
        [IsRequired]
        public string securePassword { get; set; }

        public override void Execute(int? optionID)
        {
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
        }
    }
}
