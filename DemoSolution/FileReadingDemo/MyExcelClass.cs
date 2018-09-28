using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel=Microsoft.Office.Interop.Excel;

namespace FileReadingDemo
{
    /// <summary>
    /// Class with custom methods for interaction with the .xlsx file
    /// </summary>
    public class MyExcelClass
    {
        private _Application excel = new Excel.Application();
        private _Workbook workbook;
        private _Worksheet worksheet;
        
        /// <summary>
        /// Open the desired workbook
        /// </summary>
        /// <param name="path">Path of the Excel workbook to open </param>
        /// <param name="ws">Worksheet to load, defaul is 1</param>
        /// <returns></returns>
        public void OpenWorkbook(string path, int ws=1)
        {
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Sheets[ws];
        }
        /// <summary>
        /// Create a new workbook in the specified path
        /// </summary>
        /// <param name="path"></param>
        public static void CreateNewWorkbook(string path)
        {
            Excel.Application newExcelApp = new Excel.Application();
            newExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            newExcelApp.Workbooks[1].SaveAs(path);
            newExcelApp.Workbooks.Close();
            newExcelApp.Quit();
        }
        /// <summary>
        /// Save changes made in the current workbook
        /// </summary>
        public void Save()
        {
            workbook.Save();
        }
        /// <summary>
        /// Save changes and close the current workbook
        /// </summary>
        public void SaveAndClose()
        {
            workbook.Save();
            workbook.Close();
            excel.Quit();
        }
        /// <summary>
        /// Save the current workbook in a new specified location
        /// </summary>
        /// <param name="newpath"></param>
        public void SaveAs(string newpath)
        {
            workbook.SaveAs(newpath);
        }

        /// <summary>
        /// Close the current opened workbook
        /// </summary>
        public void CloseWorkbook()
        {
            workbook.Close();
            excel.Quit();
        }
        /// <summary>
        /// Change between sheets of the same workbook, if the sheet is not created it will
        /// redirect you to the latest added sheet 
        /// </summary>
        /// <param name="ws">New sheet to open, default is 1</param>
        public void ChangeWorksheet(int ws=1)
        {
            if (workbook.Sheets.Count<ws)
            {
                ws = workbook.Sheets.Count;
                Console.WriteLine("Not enough sheets in the current workbook, redirected to the last sheet in the document");
            }
            worksheet = workbook.ActiveSheet[ws];
        }
        /// <summary>
        /// Get the value of a single cell in the current worksheet and return it as a string
        /// </summary>
        /// <param name="row">Should start from 1</param>
        /// <param name="column">Should start from 1</param>
        /// <returns>Value of the selected cell, if row or column are
        /// less than 1, both get assign to 1</returns>
        public string ReadCell(int row,int column)
        {
            if (row<1||column<1)
            {
                row = column = 1;
            }
            var value = worksheet.Cells[row, column] != null ? worksheet.Cells[row, column].Value2.ToString():string.Empty;
            return value;
        }
        /// <summary>
        /// Find all the places in the column in wich the value is found and returns them as a
        /// list of int[],if row is not especified it will start from 1
        /// </summary>
        /// <param name="value"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public List<int[]> FindValueInColumn(string value, int column,int row=1)
        {
            List<int[]> result = new List<int[]>();
            bool isOver = false;
            string castedValue;
            while (!isOver)
            {
                var cellValue = worksheet.Cells[row, column].Value2;

                if (cellValue!=null)
                {
                    castedValue = cellValue.ToString();
                    if (castedValue.Equals(value))
                    {
                        result.Add(new int[] { row, column });
                    }
                   
                }
                else
                {
                    isOver = true;
                }
                row++;
            }
            return result;
        }

        /// <summary>
        /// Search in a specific range of cells in the current workshet, and return the values as strings
        /// </summary>
        /// <param name="initialRow"></param>
        /// <param name="initialColumn"></param>
        /// <param name="finalRow"></param>
        /// <param name="finalColumn"></param>
        /// <returns></returns>
        public string[,] ReadRange(int initialRow, int initialColumn,int finalRow,int finalColumn)
        {
            Excel.Range range = worksheet.Range[worksheet.Cells[initialRow,initialColumn],worksheet.Cells[finalRow,finalColumn]];
            object[,] rangeValues = range.Value2;
            string[,] returnValues = new string[finalRow-initialRow+1,finalColumn-initialColumn+1];
            for (int row=0; row < rangeValues.GetLength(0);row++)
            { 
                for (int column = 0; column < rangeValues.GetLength(1); column++)
                {
                    //Careful is needed here, Excel ranges always start from 1
                    returnValues[row, column] = rangeValues[row+1, column+1]!=null?rangeValues[row + 1, column + 1].ToString():string.Empty;
                }
            }
            return returnValues;
             
        }

       

        /// <summary>
        /// Write to a single specified cell, if the row or column index is less than 1, the default cell is [1,1] 
        /// </summary>
        /// <param name="row">Row index</param>
        /// <param name="column">Column index</param>
        /// <param name="value">Value to write in the cell</param>
        public void WriteToCell(int row, int column,string value)
        {
            if (row<1||column<1)
            {
                row = column = 1;
            }
            worksheet.Cells[row,column]=value;
        }

        public void WriteToRange(int initialRow,int initialColumn,int finalRow,int finalColumn, string[,]range)
        {
            int rows = finalRow - initialRow + 1;
            int columns = finalColumn - initialColumn + 1;
            if (range.GetLength(0)==rows&&range.GetLength(1)==columns)
            {
                Excel.Range rangeToWrite = worksheet.Range[worksheet.Cells[initialRow, initialColumn], worksheet.Cells[finalRow, finalColumn]];
                rangeToWrite.Value2=range;
            }
           
        }

       
    }
}
