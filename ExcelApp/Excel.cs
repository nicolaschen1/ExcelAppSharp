/*
ExcelAppSharp
VERSION: 1.0

Input: Excel File

Output: Make some processings in the Excel file

Description: This software tool allows to make processings in the Excel file.

Developer: Nicolas CHEN
*/

using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelApp
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        
        public Excel()
        {

        }

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path, ReadOnly: false);
            ws = wb.Worksheets[Sheet];            
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];           
        }

        public void CreateNewSheet()
        {
            Worksheet temptsheet = wb.Worksheets.Add(After: ws);
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            if (ws.Cells[i, j].Value2 != null) //value2 is the value stored int the cell
                return ws.Cells[i, j].Value2;
            else
                return "";             
        }

        /// <summary>
        /// Read a table
        /// </summary>
        /// <param name="starti">first row</param>
        /// <param name="starty">first column</param>
        /// <param name="endi">last row</param>
        /// <param name="endy">last column</param>
        /// <returns></returns>
        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti + 1, endy - starty + 1];
            for (int p = 1; p <= endi - starti + 1; p++)
            {
                for (int q = 1; q <= endy - starty + 1; q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();
                }
            }
            return returnstring;
        }

        public void WriteRange(int starti, int starty, int endi, int endy, string[,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }

        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void SelectWorksheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }

        public void DeleteWorksheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }

        public void ProtectSheet()
        {
            ws.Protect();
        }

        public void ProtectSheet(string Password)
        {
            ws.Protect(Password);
        }

        public void UnprotectSheet()
        {
            ws.Unprotect();
        }

        public void UnprotectSheet(string Password)
        {
            ws.Unprotect(Password);
        }

        public void Close()
        {
            wb.Close();
            //excel.Quit();
        }
    }
}
