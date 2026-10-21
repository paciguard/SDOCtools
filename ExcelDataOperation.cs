using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Spire.Xls;

namespace AcAdAmComm
{
    /// <summary>
    /// File Name       ：ExcelDataOperation
    /// Description     ：Read excel template data
    /// Author          ：SY
    /// Create Date     ：2021/07/27 10:43
    /// Modifier        ：
    /// Modifier Date   ：
    /// Revision History：
    /// Version         ：V1.0
    public class ExcelDataOperation
    {
        #region Read a sheet page in Excel
        /// <summary>
        /// Read a sheet page in Excel
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>All data of sheet page</returns>
        public Excel._Worksheet ReadExcelSheet(string fileName)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;

            object missing = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlWorkbook = xlApp.Workbooks.Open(fileName);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlWorksheet.Activate();

            return xlWorksheet;
        }
        #endregion

        #region Read EXECL basic data into dataset
        /// <summary>
        /// Read EXECL basic data into dataset
        /// </summary>
        /// <param name="filePath">Specify EXECL file path</param>
        /// <param name="fileName">Specify EXECL file name</param>
        /// <param name="isFirstLineColumnName">Sets whether the first row is a column name</param>
        /// <returns>Returns a DataTable dataset</returns>
        public static DataSet ExcelToDataSet(string filePath, string fileName, bool isFirstLineColumnName)
        {
            //DataSet name，Is the first name of XML, defined as the file name
            DataSet dataSet = new DataSet(fileName);
            int startRow = 0;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook = null;
                    // If it is the excel version of 2007 +
                    if (filePath.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    // If it is the excel version of 2003 -
                    else if (filePath.IndexOf(".xls") > 0)
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    if (workbook != null)
                    {
                        //Each sheet in excel is read circularly,
                        //and each sheet page is converted into a DataTable and placed in the dataset
                        for (int p = 0; p < workbook.NumberOfSheets; p++)
                        {
                            ISheet sheet = workbook.GetSheetAt(p);
                            DataTable dataTable = new DataTable();
                            dataTable.TableName = sheet.SheetName + "TABLE";
                            //When sheet has data
                            if (sheet != null)
                            {
                                int rowCount = sheet.LastRowNum;//Get total rows
                                if (rowCount > 0)
                                {
                                    //SD basic data sheet read datable start
                                    if (sheet.SheetName.Substring(0, 2).Equals("SD"))
                                    {
                                        IRow firstRow31 = sheet.GetRow(30);//Get line 31
                                        int cellCount31 = firstRow31.LastCellNum;//Get the total number of columns
                                        for (int i = firstRow31.FirstCellNum; i < cellCount31; i++)
                                        {
                                            ICell cell = firstRow31.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    DataColumn column = new DataColumn(cell.StringCellValue);
                                                    dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                        //Fill row
                                        for (int i = 31; i <= rowCount - 31; ++i)
                                        {
                                            IRow row = sheet.GetRow(i);
                                            if (row == null)
                                            {
                                                continue;
                                            }

                                            DataRow dataRow = dataTable.NewRow();
                                            for (int j = row.FirstCellNum; j < cellCount31; j++)
                                            {
                                                ICell cell = row.GetCell(j);
                                                if (cell == null)
                                                {
                                                    dataRow[j] = "";
                                                }
                                                else
                                                {
                                                    //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                                    switch (cell.CellType)
                                                    {
                                                        case CellType.Blank:
                                                            dataRow[j] = "";
                                                            break;
                                                        case CellType.Numeric:
                                                            short format = cell.CellStyle.DataFormat;
                                                            //Time format（2015.12.5、2015/12/5、2015-12-5）
                                                            if (format == 14 || format == 31 || format == 57 || format == 58)
                                                                dataRow[j] = cell.DateCellValue;
                                                            else
                                                                dataRow[j] = cell.NumericCellValue.tostring();
                                                            break;
                                                        case CellType.String:
                                                            dataRow[j] = cell.StringCellValue;
                                                            break;
                                                    }
                                                }
                                            }
                                            dataTable.Rows.Add(dataRow);
                                        }

                                    }
                                    //SD basic data sheet read datable end
                                    //PCS basic data sheet read datable start
                                    else if (sheet.SheetName.Substring(0, 3).Equals("PCS"))
                                    {
                                        IRow firstRow3 = sheet.GetRow(2);//Get line 3
                                        int cellCount3 = firstRow3.LastCellNum;//Get the total number of columns
                                        for (int i = firstRow3.FirstCellNum; i < cellCount3; i++)
                                        {
                                            ICell cell = firstRow3.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    DataColumn column = new DataColumn(cell.StringCellValue);
                                                    dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                        //Fill row
                                        for (int i = 3; i < rowCount - 3; i++)
                                        {
                                            IRow row = sheet.GetRow(i);
                                            if (row == null)
                                            {
                                                continue;
                                            }

                                            DataRow dataRow = dataTable.NewRow();
                                            for (int j = row.FirstCellNum; j < cellCount3; j++)
                                            {
                                                ICell cell = row.GetCell(j);
                                                if (cell == null)
                                                {
                                                    dataRow[j] = "";
                                                }
                                                else
                                                {
                                                    //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                                    switch (cell.CellType)
                                                    {
                                                        case CellType.Blank:
                                                            dataRow[j] = "";
                                                            break;
                                                        case CellType.Numeric:
                                                            short format = cell.CellStyle.DataFormat;
                                                            //Time format（2015.12.5、2015/12/5、2015-12-5）
                                                            if (format == 14 || format == 31 || format == 57 || format == 58)
                                                                dataRow[j] = cell.DateCellValue;
                                                            else
                                                                dataRow[j] = cell.NumericCellValue.tostring();
                                                            break;
                                                        case CellType.String:
                                                            dataRow[j] = cell.StringCellValue;
                                                            break;
                                                    }
                                                }
                                            }
                                            dataTable.Rows.Add(dataRow);
                                        }
                                    }
                                    //PCS basic data sheet read datable end
                                    //Dropdownmenu data sheet read datable start
                                    else
                                    {
                                        IRow firstRow = sheet.GetRow(0);//Get first row
                                        int cellCount = firstRow.LastCellNum;//Get the total number of columns
                                        //Building columns of datatable
                                        if (isFirstLineColumnName)
                                        {
                                            //If the first row is a column name, it is read from the second row
                                            startRow = 1;
                                            for (int i = firstRow.FirstCellNum; i < cellCount; i++)
                                            {
                                                ICell cell = firstRow.GetCell(i);
                                                if (cell != null)
                                                {
                                                    if (cell.StringCellValue != null)
                                                    {
                                                        DataColumn column = new DataColumn(cell.StringCellValue);
                                                        dataTable.Columns.Add(column);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            for (int i = firstRow.FirstCellNum; i < cellCount; i++)
                                            {
                                                DataColumn column = new DataColumn("column" + (i + 1));
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                        //Fill row
                                        for (int i = startRow; i <= rowCount; i++)
                                        {
                                            IRow row = sheet.GetRow(i);
                                            if (row == null)
                                            {
                                                continue;
                                            }

                                            DataRow dataRow = dataTable.NewRow();
                                            for (int j = row.FirstCellNum; j < cellCount; j++)
                                            {
                                                ICell cell = row.GetCell(j);
                                                if (cell == null)
                                                {
                                                    dataRow[j] = "";
                                                }
                                                else
                                                {
                                                    //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                                    switch (cell.CellType)
                                                    {
                                                        case CellType.Blank:
                                                            dataRow[j] = "";
                                                            break;
                                                        case CellType.Numeric:
                                                            short format = cell.CellStyle.DataFormat;
                                                            //Time format（2015.12.5、2015/12/5、2015-12-5）
                                                            if (format == 14 || format == 31 || format == 57 || format == 58)
                                                                dataRow[j] = cell.DateCellValue;
                                                            else
                                                                dataRow[j] = cell.NumericCellValue.tostring();
                                                            break;
                                                        case CellType.String:
                                                            dataRow[j] = cell.StringCellValue;
                                                            break;
                                                    }
                                                }
                                            }
                                            dataTable.Rows.Add(dataRow);
                                        }
                                    }
                                    //Dropdownmenu data sheet read datable end
                                }
                            }
                            dataSet.Tables.Add(dataTable);
                        }
                    }
                } 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dataSet;
        }
        #endregion

        #region Read sumlist file to dataset
        /// <summary>
        /// Read sumlist file to dataset
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static DataSet SumListExcelToDataSat(string filePath, List<string> strList)
        {
            var strCellNamelist = strList;
            DataSet dataSet = new DataSet();
            DataTable dataTableBody = new DataTable();
            DataTable dataTableHead = new DataTable();
            int startRow = 22;
            int startCell = 7;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook = null;
                    workbook = new XSSFWorkbook(fs);
                    if (workbook != null)
                    {
                        ISheet sheet = workbook.GetSheetAt(0);
                        dataTableBody.TableName = "DataTableBody";
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//Get total rows
                            //Lastcellnum is the column number of the last column
                            //Gets the total number of columns in 23 rows
                            int cellCount = sheet.GetRow(startRow).LastCellNum;
                            //From line 23 to the last line, the data of the line is fetched circularly
                            for (int i = startRow; i < rowCount; i++)
                            {
                                //Get each row
                                IRow eachRow = sheet.GetRow(i);
                                //Gets the value of the specified column D（S1,S2,S3......）
                                ICell cell = eachRow.GetCell(3);
                                //Judge whether the column name exists in the list, and create a new column if it exists
                                if (strCellNamelist.Contains(cell.StringCellValue))
                                {
                                    DataColumn column = new DataColumn(cell.StringCellValue);
                                    dataTableBody.Columns.Add(column);
                                }
                            }
                            //Loop column, starting from column 7 to the last column
                            for (int j = startCell - 1; j < cellCount; j++)
                            {
                                //Add data to datatable
                                DataRow dataRowBody = dataTableBody.NewRow();
                                int indexD = 0;
                                //Loop line, starting from line 23 to the last line
                                for (int i = startRow; i < rowCount; i++)
                                {
                                    IRow eachRow = sheet.GetRow(i);//获取每行
                                    ICell cellD = eachRow.GetCell(3);
                                    if (strCellNamelist.Contains(cellD.StringCellValue))
                                    {
                                        //Strong rotation whether it is empty or not
                                        dataRowBody[indexD] = Convert.ToString(eachRow.GetCell(j));
                                        indexD++;
                                    }
                                }
                                dataTableBody.Rows.Add(dataRowBody);
                            }
                            //Read header data
                            dataTableHead.TableName = "DataTableHead";
                            //From row 13 to row 21, cycle the data of column D
                            for (int e = 12; e <= 20; e++)
                            {
                                //Get each row
                                IRow eachRow = sheet.GetRow(e);
                                //Gets the value of the specified column D（S2,S3,S4,S5,S6,S7）
                                ICell cellD = eachRow.GetCell(3);
                                DataColumn column = new DataColumn(cellD.StringCellValue);
                                dataTableHead.Columns.Add(column);
                            }
                            int indexE = 0;
                            DataRow dataRowHead = dataTableHead.NewRow();
                            //From row 13 to row 21, the data of column e is retrieved cyclically
                            for (int f = 12; f <= 20; f++)
                            {
                                //Get each row
                                IRow eachRow = sheet.GetRow(f);
                                dataRowHead[indexE] = eachRow.GetCell(4).StringCellValue;
                                indexE++;
                            }
                            //Add data to datatablehead, just one piece of data
                            dataTableHead.Rows.Add(dataRowHead);
                        }
                        dataSet.Tables.Add(dataTableHead);
                        dataSet.Tables.Add(dataTableBody);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dataSet;
        }
        #endregion

        #region Excel File OutPut
        /// <summary>
        /// Excel File OutPut
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        public static void SaveToExcelFile(XSSFWorkbook workbook, string path)
        {
            using (FileStream newfs = new FileStream(path, FileMode.Create))
            {
                workbook.Write(newfs);
            }
        }
        #endregion

        #region Pdf File OutPut
        /// <summary>
        /// Pdf File OutPut
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        public static void SaveToPdfFile(string excelPath, string pdfPath)
        {
            //Create a wordbook class object and load the excel document to be converted
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(excelPath, ExcelVersion.Version2010);
            //Save excel document as PDF
            workbook.SaveToFile(pdfPath, FileFormat.PDF);
        }
        #endregion
    }
}
