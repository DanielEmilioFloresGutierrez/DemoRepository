using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileReadingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            MyExcelClass myExcelClass = new MyExcelClass();
            //The next line retrieves the current project path
            string root = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory()))+"\\";
            string filename = "TestBook.xlsx";
            try
            {
                //We can open the specified workbook
                myExcelClass.OpenWorkbook(root + filename, 1);
                Console.WriteLine("Workbook from path " + root + filename + " opened\n");

                //MyExcelClass.CreateNewWorkbook("C:/Documents/NewWorkbook.xlsx");       *Or we can create a new one*

                //We can read single cells
                Console.WriteLine("Single cell reading: " + myExcelClass.ReadCell(2, 1) + "\n");

                //Or we can read an entire range
                Console.WriteLine("Range reading: \n");
                string[,] range = (myExcelClass.ReadRange(2, 1, 20, 1));
                for (int row = 0; row < range.GetLength(0); row++)
                {
                    for (int column = 0; column < range.GetLength(1); column++)
                    {
                        Console.WriteLine(range[row, column].ToString() + "\t");
                    }
                }
                Console.WriteLine();


                //We can find values in a selected column and optionally start the search from a specific row
                List<int[]> searchResult=myExcelClass.FindValueInColumn("Find me",1);
                Console.WriteLine("The text \"Find me\" was located in cells:\n");
                foreach (var cell in searchResult)
                {
                    string row = cell[0].ToString();
                    string column = cell[1].ToString();
                    Console.WriteLine($"[{row},{column}]");
                }

                
               
                
                //Values can be writed and saved in the same file or a new one
                string[,] values = GenerateRangeValues(19, 4);//<===== Generation of random numbers to simulate data
                Console.WriteLine("Values to store, enter to proceed.\n\n");
                for (int row = 0; row < values.GetLength(0); row++)
                {
                    for (int column = 0; column < values.GetLength(1); column++)
                    {
                        Console.Write(values[row, column].ToString() + "\t");
                    }
                    Console.WriteLine();
                }
                Console.ReadLine();
                myExcelClass.WriteToRange(2,2,20,5,values);
                myExcelClass.Save();//<=====Save changes in the same workbook
                myExcelClass.SaveAs(root+"NewWorkbook.xlsx");//<=====Save changes in new workbook
                myExcelClass.CloseWorkbook();
              
            }
            catch (Exception e)
            {
                myExcelClass.CloseWorkbook();
                throw e;
            }
          

            string[,] GenerateRangeValues(int rows, int columns)
            {
                Random rd = new Random();
                string[,] values= new string[rows,columns];
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        values[i, j] = rd.Next(1, 1000).ToString();
                    }
                }
                return values;
            }
           
           
            
        }
    }
}
