using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebDriveProject_Selenium.Services
{
    public class ExcelUtility
    {
        public static List<Datacollection> dataCol = new List<Datacollection>();
        private static string FilePath;

        /// <summary>
        /// Excel to Data Table convertion 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string fileName, string sheetName)
        {
            //open file and returns as Stream
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                //Createopenxmlreader via ExcelReaderFactory
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    //Return as DataSet
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        //Set the First Row as Column Name
                        ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    //Get all the Tables
                    DataTableCollection table = result.Tables;
                    //Store it in DataTable
                    DataTable resultTable = table[sheetName];
                    //return
                    return resultTable;
                }
            }
        }

        /// <summary>
        /// Populate the data in collection
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        public static void DataPopulateInCollection(string fileName, string sheetName)
        {
            DataTable table = ExcelToDataTable(fileName, sheetName);
            dataCol.DefaultIfEmpty();
            //Iterate through the rows and columns of the Table
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    Datacollection dtTable = new Datacollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    //Add all the details for each row
                    dataCol.Add(dtTable);
                }
            }
        }

        /// <summary>
        /// Read data from Collection
        /// </summary>
        /// <param name="rowNumber"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static string ReadData(int rowNumber, string columnName)
        {
            try
            {
                //Retriving Data using LINQ to reduce much of iterations
                string data = (from colData in dataCol
                               where colData.colName == columnName && colData.rowNumber == rowNumber
                               select colData.colValue).SingleOrDefault();

                //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;
                return data.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        /// <summary>
        /// Get data rows count from Collection
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int DataCount(string fileName, string sheetName)
        {
            try
            {
                DataTable table = new DataTable();
                table = ExcelToDataTable(fileName, sheetName);
                return table.Rows.Count;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0;
            }
        }

    }

    public class Datacollection
    {
        public int rowNumber { get; set; }
        public string colName { get; set; }
        public string colValue { get; set; }
    }
}
