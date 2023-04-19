using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Appium.Utilities;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UIAutomation.Pages;

namespace UIAutomation.Utils_Misc
{
    class ExlParser : Base
    {
        // defining max column size
        public static int maxDataSets;
        public static XSSFWorkbook hssfwb;

        // Initializing constructor to read datafile
        public ExlParser(string dataFilePath)
        {
            // Reading data file into npoi hssf work book
            using (FileStream file = new FileStream(dataFilePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
                XSSFFormulaEvaluator.EvaluateAllFormulaCells(hssfwb); // Refreshes all formulas
                file.Close();
            }
        }

        // To return the data dictionary for the particular test case and environment
        public Dictionary<string, string> LoadDataDictionay(string testCaseName, string persona, string environment)
        {
            // Create a dictionary with parameters as keys and arrayList of values
            Dictionary<string, string> dict = new Dictionary<string, string>();
            IRow currentRow, headerRow;
            //ArrayList executeIndicator = new ArrayList();
            var envCol=0;
            ISheet dataSheet = hssfwb.GetSheet(persona);
            IEnumerator currentRows = dataSheet.GetRowEnumerator();
            // Get to the first row and set maxDataSets
            currentRows.MoveNext();
            headerRow = (XSSFRow)currentRows.Current;
            maxDataSets = headerRow.LastCellNum;
            // Finding the environment
            for (int i = 2; i < maxDataSets; i++)
            {
                if (GetCellValue(headerRow, i).Equals(environment, StringComparison.InvariantCultureIgnoreCase))
                {
                    envCol = i;
                }
            }
            if (envCol == 0)
            {
                throw new Exception($"Environment [{environment}] not found in Data File!");
            }
            // Finding the scenario
            while (currentRows.MoveNext())
            {
                headerRow = (XSSFRow)currentRows.Current;
                if (GetCellValue(headerRow, 0) == testCaseName)
                {
                    break;
                }
            }
            // Throw Exception if result not found
            if (GetCellValue(headerRow, 0) == "End")
            {
                throw new Exception($"Script [{testCaseName}] not found in Data File!");
            }
            // Looping through the parameters in the data set
            currentRow = headerRow;
            currentRows.MoveNext();
            currentRow = (XSSFRow)currentRows.Current;
            while (GetCellValue(currentRow, 0) == "")
            {
                dict.Add(GetCellValue(currentRow, 1), GetCellValue(currentRow, envCol));
                currentRows.MoveNext();
                currentRow = (XSSFRow)currentRows.Current;
            }
            return dict;
        }

        // Returns the string value of cell evaluating any fomulas if present
        string GetCellValue(IRow row, int column)
        {
            string str;
            switch (row.GetCell(column).CellType)
            {
                case CellType.Formula:
                    try
                    {
                        str = row.GetCell(column).StringCellValue;
                    }
                    catch (Exception e)
                    {
                        if (e.Message == "Cannot get a text value from a numeric formula cell")
                        {
                            str = row.GetCell(column).NumericCellValue.ToString();
                        }
                        else
                        {
                            throw e;
                        }
                    }
                    break;
                default:
                    str = row.GetCell(column).ToString();
                    break;
            }
            return str;
        }
    }
}
