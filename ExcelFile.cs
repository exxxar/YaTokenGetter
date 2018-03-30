using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace RegAccYandex
{
    public class ExcelFile
    {
        protected Excel.Application xlApp ;
        protected Excel.Workbook xlWorkbook ;
        protected Excel._Worksheet xlWorksheet ;
        protected Excel.Range xlRange ;

        public ExcelFile()
        {
           
        }

        public void createExcel(string filePath)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                throw new Exception("Excel is not properly installed!!");
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Имя";
            xlWorkSheet.Cells[1, 2] = "Фамилия";
            xlWorkSheet.Cells[1, 3] = "Логин";
            xlWorkSheet.Cells[1, 4] = "Пароль";
            xlWorkSheet.Cells[1, 5] = "Вопрос";
            xlWorkSheet.Cells[1, 6] = "Ответ";
            xlWorkSheet.Cells[1, 7] = "Дата регистрации";
            xlWorkSheet.Cells[1, 8] = "ID приложения";
            xlWorkSheet.Cells[1, 9] = "Пароль приложение";
            xlWorkSheet.Cells[1, 10] = "Callback url";

            xlWorkBook.SaveAs(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + filePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);

            this.closeExcel();
        }
      
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        Microsoft.Office.Interop.Excel.Sheets xlBigSheet;
        Microsoft.Office.Interop.Excel.Sheets xlSheet;
        object misValue;

        public void append(string filePath,int index2, User user)
        {

            int index = 3;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0,
                        false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                         "", true, false, 0, true, false, false);

            xlBigSheet = xlWorkBook.Worksheets;
            string x = "Sheet1";
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBigSheet.get_Item(1);

            xlWorksheet.Cells[index, 1] = user.name;
            xlWorksheet.Cells[index, 2] = user.tname;
            xlWorksheet.Cells[index, 3] = user.login;
            xlWorksheet.Cells[index, 4] = user.pass;
            xlWorksheet.Cells[index, 5] = user.question;
            xlWorksheet.Cells[index, 6] = user.answer;
            xlWorksheet.Cells[index, 7] = user.date_reg;
            xlWorksheet.Cells[index, 8] = user.prog_id;
            xlWorksheet.Cells[index, 9] = user.prog_pass;
            xlWorksheet.Cells[index, 10] = user.callback_url;

            xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                    misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                    misValue, misValue, misValue,
                    misValue, misValue);

            xlWorkBook.Close(misValue, misValue, misValue);
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            //Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            //if (xlApp == null)
            //{
            //    throw new Exception("Excel is not properly installed!!");
            //}

            //xlApp = new Excel.Application();
            //xlWorkbook = xlApp.Workbooks.Open(filePath);
            //xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            //xlRange = xlWorksheet.UsedRange;
            //if (xlWorksheet.Cells[index, 3].Value2.ToString().Trim() != "")
            //    index++;

            //xlWorksheet.Cells[index, 1] = user.name;
            //xlWorksheet.Cells[index, 2] = user.tname;
            //xlWorksheet.Cells[index, 3] = user.login;
            //xlWorksheet.Cells[index, 4] = user.pass;
            //xlWorksheet.Cells[index, 5] = user.question;
            //xlWorksheet.Cells[index, 6] = user.answer;
            //xlWorksheet.Cells[index, 7] = user.date_reg;
            //xlWorksheet.Cells[index, 8] = user.prog_id;
            //xlWorksheet.Cells[index, 9] = user.prog_pass;
            //xlWorksheet.Cells[index, 10] = user.callback_url;

            //xlWorkbook.Save();
            //xlWorkbook.Close(true);

        }

        public List<User> readExecel(string filePath)
        {
            List<User> users = new List<User>();

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(filePath);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            if (rowCount < 2)
                return null;


            for (int i = 2; i <= rowCount; i++)
            {

                User u = new User();
                u.login = "" + xlRange.Cells[i, 3].Value2.ToString();
                u.pass = "" + xlRange.Cells[i, 4].Value2.ToString();
                u.prog_id = "" + xlRange.Cells[i, 8].Value2.ToString();
                u.prog_pass = "" + xlRange.Cells[i, 9].Value2.ToString();

                users.Add(u);
            }


            this.closeExcel();

            return users;
        }

        public void closeExcel()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
           
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
