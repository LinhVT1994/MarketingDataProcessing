using MarketingDataProcessing.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using static MarketingDataProcessing.Attributes.AutoIncrementAttribute;
using static MarketingDataProcessing.Attributes.UniqueAttribute;
using static MarketingDataProcessing.Attributes.ExcelColumnAttribute;
using static MarketingDataProcessing.Attributes.DistinguishAttribute;
using static MarketingDataProcessing.Attributes.ForeignKeyAttribute;
using static MarketingDataProcessing.Attributes.PrimaryKeyAttribute;
using static MarketingDataProcessing.Attributes.PropertyInfoExtensions;
using System.Reflection;
using System.Data.SqlClient;
using MarketingDataProcessing.Models;
using Newtonsoft.Json;

namespace MarketingDataProcessing.Utilities
{
    class ExcelDataAccess
    {
        private  string Url = null;
        private Excel.Application _XlApplication = null;
        Excel.Workbook _XlWorkBook = null;
        private int _StartRowInExcel = 1;
        private int _MaxColumn = 0;
        private int _HeaderPosition = 1;
        private ExcelDataAccess(string url)
        {
            Url = url;
            try
            {
                _XlApplication = new Excel.Application();
                _XlApplication.Visible = false;
                _XlApplication.DisplayAlerts = false;
              
            }
            catch(Exception)
            {

            }
        }
        public static ExcelDataAccess  Execute(string url)
        {
            return new ExcelDataAccess(url);
        }
        public void Open<T>()
        {
            try
            {
                _XlWorkBook = _XlApplication.Workbooks.Open(Url);
                int i = 1;
                foreach (Excel.Worksheet xlworksheet in _XlWorkBook.Worksheets)
                {
                    ReadAWorkSheet<T>(i);
                    i++;
                }
                
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ExcelToSqlFullColumn<T>()
        {
            try
            {
                _XlWorkBook = _XlApplication.Workbooks.Open(Url);
                int i = 1;
                foreach (Excel.Worksheet xlworksheet in _XlWorkBook.Worksheets)
                {
                    ReadAWorkSheetWithFullColumn<T>(i);
                    i++;
                }

            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ReadAWorkSheetWithFullColumn<T>(int index)
        {
            int amountOfUsedRows = 0;
            int amountOfUsedColumns = 0;
            Excel.Range xlRange;
            Excel.Worksheet xlworkSheet = null;
            try
            {
                xlworkSheet = (Excel.Worksheet)_XlWorkBook.Sheets[index];
                xlworkSheet.Unprotect();
            }
            catch (Exception)
            {

                throw;
            }
            xlRange = xlworkSheet.UsedRange;
            amountOfUsedRows = xlRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
            amountOfUsedColumns = xlRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;
           
            string header = null;
            string value = null;
            for (int row = StartRowInExcel; row < amountOfUsedRows; row++)
            {
                Dictionary<string, string> valuesInARow = new Dictionary<string, string>();
                for (int column = 1; column < amountOfUsedColumns; column++)
                {
                        Excel.Range cell = xlworkSheet.Cells[row, column];
                        Excel.Range columnHeader = xlworkSheet.Cells[HeaderPosition, column];
                        if (cell != null && columnHeader != null)
                        {
                            if (cell.Value != null && columnHeader.Value != null)
                            {
                                header = xlworkSheet.Cells[HeaderPosition, column].Value.ToString();
                                value = xlworkSheet.Cells[row, column].Value.ToString();

                                if (valuesInARow.ContainsKey(header))
                                {
                                    valuesInARow[header] = valuesInARow[header] + ", " + value;
                                }
                                else
                                {
                                    valuesInARow.Add(header, value);
                                }
                             }
                         }
                }
                if (valuesInARow.Count <= 0)
                {
                    continue;
                }
                else
                {
                    Synthesis temp = new Synthesis();
                    temp.Json = JsonConvert.SerializeObject(valuesInARow);
                    //Update to sql
                    RequestToSql<Synthesis>(temp);
                }
            }
        }
        public void ReadAWorkSheet<T>(int index)
        {
            int amountOfUsedRows = 0;
            int amountOfUsedColumns = 0;
            Excel.Range xlRange;
            Excel.Worksheet xlworkSheet = null;
            bool hasALeastOneUnique = true;
            List<string> propertyNames = RequiredAttribute.GetRequiredPropertiesName(typeof(T));
            List<string> uniqueProperties = GetUniqueProperties(typeof(T));

            if (GetUniqueProperties(typeof(T)).Count <= 0)
            {
                hasALeastOneUnique = false;
            }

            try
            {
                xlworkSheet = (Excel.Worksheet)_XlWorkBook.Sheets[index];
                xlworkSheet.Unprotect();
            }
            catch (Exception)
            {

                throw;
            }
            Dictionary<string, T> dicResult = new Dictionary<string, T>();
            xlRange = xlworkSheet.UsedRange;
            amountOfUsedRows = xlRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
            amountOfUsedColumns = xlRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;
            _MaxColumn = amountOfUsedColumns;
            for (int row = StartRowInExcel; row < amountOfUsedRows; row++)
            {
                T newObj = (T)Activator.CreateInstance(typeof(T));
                string key = null;
                foreach (string pName in propertyNames)
                {
                    PropertyInfo propertyInfo = newObj.GetType().GetProperty(pName);
                    if (IsPrimaryKey(typeof(T), pName)) // is primany key.
                    {
                        if (IsAutoIncrement(typeof(T), pName))
                        {
                            if (!hasALeastOneUnique)
                            {
                                key = (dicResult.Count + 1).ToString();
                                propertyInfo.SetValue(newObj, dicResult.Count + 1);
                            }
                            
                        }
                    }
                    else if (IsUnique(typeof(T), pName))
                    {
                        try
                        {
                            string columnName = null;
                            key = HandleForUniqueKey(newObj, pName, xlworkSheet, row, out columnName);
                            if (string.IsNullOrWhiteSpace(key))
                            {
                                string message = string.Format("Error at: Cell[{0},{1}] Handled: {2} Message: {3}", row, columnName, "Ignore", "Can't get value on this cell.");
                                SetErrorInfoMarkForRow(xlworkSheet, row, amountOfUsedColumns);
                                LoggingHelper.WriteDown(message);
                                break;
                            }
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                    }
                    else
                    {
                        HandleForRequiredProperty(newObj, pName, xlworkSheet, row);
                    }
                }
                // if the key has not setted any value then ignore it.
                if (string.IsNullOrWhiteSpace(key))
                {
                    //
                    continue;
                }
                else
                {
                    if (!dicResult.ContainsKey(key))
                    {
                        dicResult.Add(key, newObj);
                        string json = JsonConvert.SerializeObject(newObj);
                        LoggingHelper.LogForSuccession(json);
                        RequestToSql<T>(newObj);

                    }
                }
            }
        }
        private int RequestToSql<T>(T parseTo)
        {
            List<string> requiredProperties = RequiredAttribute.GetRequiredPropertiesName(parseTo.GetType());
            string table = typeof(T).GetAttributeValue((SqlParameterAttribute dna) => dna.PropertyName);
            List<SqlParameter> parameters = new List<SqlParameter>();
            foreach (string property in requiredProperties)
            {
                if (IsAutoIncrement(typeof(T), property))
                {

                }
                else
                {
                    string paramName = SqlParameterAttribute.GetNameOfParameterInSql(parseTo.GetType(), property);
                    PropertyInfo propertyInfo = parseTo.GetType().GetProperty(property);
                    object result = propertyInfo.GetValue(parseTo);
                    if (result != null)
                    {
                        object paramValue = propertyInfo.GetValue(parseTo);
                        if (propertyInfo.PropertyType == typeof(string))
                        {
                            parameters.Add(new SqlParameter(paramName, paramValue));
                        }
                        else if (propertyInfo.PropertyType == typeof(int))
                        {
                            parameters.Add(new SqlParameter(paramName, paramValue));
                        }
                        else if (propertyInfo.PropertyType == typeof(double))
                        {
                            parameters.Add(new SqlParameter(paramName, paramValue));
                        }
                        else if (propertyInfo.PropertyType.BaseType == typeof(Element))
                        {
                            string refId = ForeignKeyAttribute.GetRefId(typeof(T), property);
                            object data = GetPrimaryKeyValue(paramValue);
                            parameters.Add(new SqlParameter(paramName, data));
                        }
                        else
                        {
                            throw new Exception("Code hasnot implemented");
                        }
                    }
                }

            }
            return CreateInsertQuery(table, parameters);
        }
        public int CreateInsertQuery(string table, List<SqlParameter> sqlParams)
        {
            if (sqlParams.Count <= 0)
            {
                return -1;
            }
            string sPropertyNames = "(";
            foreach (SqlParameter para in sqlParams)
            {
                sPropertyNames += "" + para.ParameterName + ",";
            }
            sPropertyNames = sPropertyNames.Remove(sPropertyNames.Length - 1);
            sPropertyNames += ")";

            string sValues = null;
            foreach (SqlParameter para in sqlParams)
            {
                sValues += "@" + para.ParameterName + ",";
            }
            sValues = sValues.Remove(sValues.Length - 1);

            StringBuilder sqlQuery = new StringBuilder();

            sqlQuery.AppendFormat("insert into {0}{1} values({2})", table, sPropertyNames, sValues);

            SqlDataAccess sqlDataAccess = new SqlDataAccess();
            return sqlDataAccess.ExecuteInsertOrUpdateQuery(sqlQuery.ToString(), sqlParams.ToArray());
        }
        private void HandleForRequiredProperty(object newObj, string property, Excel.Worksheet worksheet , int row)
        {
            Dictionary<string, string> columnMapping = ColumnNamesMapping(newObj.GetType());
            PropertyInfo propertyInfo = newObj.GetType().GetProperty(property);
            if (columnMapping.ContainsKey(property)) // if this property has value will get from in Excel file.
            {
                string columnName = null;
                if (true)
                {

                }
                string returnedValue = GetValueInCell(columnMapping, property, worksheet, row, out columnName);

                if (!string.IsNullOrWhiteSpace(returnedValue))
                {
                    propertyInfo.SetValueByDataType(newObj, returnedValue);
                }
                else
                {
                    propertyInfo.SetValueByDataType(newObj, null);
                }
            }
            else
            {
                propertyInfo.SetValueByDataType(newObj, null);
            }
        }
        private string HandleForUniqueKey(object newObj, string property, Excel.Worksheet workSheet , int row, out string rowName)
        {
            string key = null;
            Dictionary<string, string> columnMapping = ColumnNamesMapping(newObj.GetType());
            // read position in excel
            if (columnMapping.ContainsKey(property))
            {
                PropertyInfo propertyInfo = newObj.GetType().GetProperty(property);
                string columnName = null;
                string returnedValue = GetValueInCell(columnMapping, property, workSheet, row, out columnName);
                key = returnedValue;
                if (string.IsNullOrWhiteSpace(returnedValue))
                {
                    string message = string.Format("Error at: Cell[{0},{1}] Handled: {2} Message: {3}", row, columnName, "Ignore", "Can't get value on this cell.");
                    LoggingHelper.WriteDown(message);
                    rowName = columnName;
                    return null;
                }
                else
                {
                    propertyInfo.SetValue(newObj, returnedValue);
                }
                rowName = columnName;
                return key;
            }
            else
            {
                rowName = null;
                CloseExcelFile();
                throw new Exception("The mapping attribute of this property is not correct. : " + property);
            }
        }   
        private string GetValueInCell(Dictionary<string, string> columnMap, string property, Excel.Worksheet workSheet, int row, out string columnName)
        {
            if (columnMap.TryGetValue(property, out columnName))
            {
                string s = null;
                try
                {
                    if (columnName.Equals("all"))
                    {
                        for (int i = 2; i <= _MaxColumn; i++)
                        {
                            Excel.Range cell = workSheet.Cells[row, i];
                            if (i == _MaxColumn)
                            {
                                if (cell.Value != null)
                                {
                                    s += workSheet.Cells[row, i].Value.ToString();
                                }
                            }
                            else
                            {
                                if (cell.Value != null)
                                {
                                    s += workSheet.Cells[row, i].Value.ToString() + " || ";
                                }
                                
                            }
                        }
                        return s;
                    }
                    else
                    {
                        Excel.Range cell = workSheet.Cells[row, columnName];
                        if (cell.Value != null)
                        {
                            s = workSheet.Cells[row, columnName].Value.ToString();
                            return s;
                        }
                    }
                   
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return null;
        }
        public int StartRowInExcel
        {
            get
            {
                return _StartRowInExcel;
            }
            set
            {
                _StartRowInExcel = value;
            }
        }
        private void CloseExcelFile()
        {
            _XlWorkBook.Save();
            _XlWorkBook.Close();
            _XlApplication.Quit();
        }
        private void SetErrorInfoMarkForRow(Excel.Worksheet workSheet, int row, int toColumn)
        {
            Excel.Range range = workSheet.Range[workSheet.Cells[row, "A"], workSheet.Cells[row, toColumn]];
            range.Interior.Color = Excel.XlRgbColor.rgbRed;
        }
        public int HeaderPosition
        {
            get
            {
                return _HeaderPosition;
            }
            set
            {
                _HeaderPosition = value;
            }
        }
    }
}
