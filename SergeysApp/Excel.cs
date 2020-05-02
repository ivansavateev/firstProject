using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SergeysApp
{
    class Excel
    {
        public static String DEFAULT_PATH = "";
        public static String CurrentPath;
        public static int SheetsCount;
        public static int HWellNumber;
        public static Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
        public static Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        public static Dictionary<string, string> HwellDictionary = new Dictionary<string, string>();
        public static Dictionary<VolumeDiameter, string> VolDiamDictionary;
        public static Dictionary<int, Dictionary<VolumeDiameter, string>> VolDiamDictionaryList = new Dictionary<int, Dictionary<VolumeDiameter, string>>();

        public static void LoadWorkBook(string filePath)
        {
            ObjWorkBook = ObjWorkExcel.Workbooks.Open(filePath);
            SheetsCount = ObjWorkBook.Sheets.Count;
            ObjWorkExcel.ScreenUpdating = false;
            GetHWellNumber();
        }

        public static Microsoft.Office.Interop.Excel.Worksheet OpenWorkSheet(int number)
        {
            if (number <= SheetsCount) //потестить
            {
                return (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[number]; //возвращает целую таблицу?
            }
            else
            {
                return null;
            }
        }

        public static void LoadVolDiameter()
        {
            for (int n = 1; n < SheetsCount; n++)
            {
                if (n != GetMainCell())
                {
                    var workSheet = OpenWorkSheet(n);
                    int lastRow = GetLastRow(n);
                    int lastColumn = GetLastColumn(n);
                    VolDiamDictionary = new Dictionary<VolumeDiameter, string>();
                    for (int j = 2; j <= lastRow;)
                    {
                        for (int i = 3; i <= lastColumn; i++)
                        {
                            if (workSheet.Cells[j + 2, i].Text.ToString() != "")
                            {
                                VolumeDiameter temp = new VolumeDiameter(workSheet.Cells[j, 1].Text.ToString(), workSheet.Cells[1, i].Text.ToString());
                                temp.Length = workSheet.Cells[j, i].Text.ToString();
                                temp.Weight = workSheet.Cells[j + 1, i].Text.ToString();
                                VolDiamDictionary.Add(temp, workSheet.Cells[j + 2, i].Text.ToString());
                            }
                        }
                        if (workSheet.Cells[j + 3, 1].Text.ToString().Contains("Объем") ||
                            workSheet.Cells[j + 3, 1].Text.ToString().Contains("Производительность") ||
                            workSheet.Cells[j + 3, 1].Text.ToString().Contains("Высота"))
                        {
                            j += 4;
                        }
                        else
                        {
                            j += 3;
                        }
                    }
                    VolDiamDictionaryList.Add(n, VolDiamDictionary);
                }
            }
            
            
        }
        public static void LoadVolDiameter2(int numberList)
        {
            var workSheet = OpenWorkSheet(numberList);
            int lastRow = GetLastRow(numberList);
            int lastColumn = GetLastColumn(numberList);
            VolDiamDictionary = new Dictionary<VolumeDiameter, string>();
            for (int j = 2; j <= lastRow;)
            {
                for (int i = 3; i <= lastColumn; i++)
                {
                    if (workSheet.Cells[j + 6, i].Text.ToString() != "")    //сделал выбор более нижний строки +6
                    {
                        VolumeDiameter temp = new VolumeDiameter(workSheet.Cells[j, 1].Text.ToString(), workSheet.Cells[1, i].Text.ToString());
                        temp.Height = workSheet.Cells[j, i].Text.ToString();  //создал Height in VolumeDiameter
                        temp.H2 = workSheet.Cells[j + 1, i].Text.ToString();
                        temp.H3 = workSheet.Cells[j + 2, i].Text.ToString();
                        temp.H4 = workSheet.Cells[j + 3, i].Text.ToString();
                        temp.D2 = workSheet.Cells[j + 4, i].Text.ToString();
                        temp.Weight = workSheet.Cells[j + 5, i].Text.ToString();
                        VolDiamDictionary.Add(temp, workSheet.Cells[j + 6, i].Text.ToString());
                    }
                }
                if (workSheet.Cells[j + 7, 1].Text.ToString().Contains("Объем") ||  //Прыгает не на 3, а теперь на 7
                    workSheet.Cells[j + 7, 1].Text.ToString().Contains("Производительность") ||
                    workSheet.Cells[j + 7, 1].Text.ToString().Contains("Высота"))
                {
                    j += 8; //Прыгает не на 4, а теперь на 8
                }
                else
                {
                    j += 7; //Прыгает не на 3, а теперь на 7
                }

            }

        }

        public static void LoadHwell(int numberList)
        {
            var workSheet = OpenWorkSheet(numberList);
            int lastRow = GetLastRow(numberList);
            for (int i = 0; i < lastRow; i++)
            {
                HwellDictionary.Add(workSheet.Cells[i + 1, 1].Text.ToString(), workSheet.Cells[i + 1, 2].Text.ToString());
            }
        }



        public static void CloseExcel()
        {
            ObjWorkExcel.ScreenUpdating = true;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
        }

        public static int GetLastColumn(int number)
        {
            var workSheet = OpenWorkSheet(number);
            var lastCell = workSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell); //1, 2 ,3 , 4    5
            return (int)lastCell.Column;
        }

        public static int GetMainCell()
        {
            for (int n = 1; n < SheetsCount; n++)
            {
                var workSheet = OpenWorkSheet(n);
                int lastRow = GetLastRow(n);

                for (int j = 2; j <= 10; j++)
                {
                    if (workSheet.Cells[j, 2].Text.ToString().Contains("H2"))
                    {
                        return n;
                    }
                }
            }
            return -1;

        }

        public static int GetLastRow(int number)
        {
            var workSheet = OpenWorkSheet(number);
            var lastCell = workSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            return (int)lastCell.Row;
        }

        public static int GetSheetNumberByName(string name)
        {
            for (int i = 1; i <= SheetsCount; i++)
            {
                if (OpenWorkSheet(i).Name.Equals(name))
                {
                    return i;
                }
            }
            return 0;
        }

        private static int GetHWellNumber()
        {
            for (int i = 1; i <= SheetsCount; i++)
            {
                if (GetLastRow(i) < 30)
                {
                    HWellNumber = i;
                }
            }
            return 0;
        }


    }
}
