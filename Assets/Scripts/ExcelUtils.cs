using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using NPOI.XSSF.Extractor;
using OfficeOpenXml;
using UnityEngine;

namespace Chart3D.DataProvider.Data.Excel
{
    /// <summary>
    /// 读取保存Excel文件类
    /// </summary>
    public class ExcelUtils
    {
        /// <summary>
        /// 读取Excel文件中指定编号的Sheet
        /// </summary>
        /// <param name="filePath">Excel文件的路径</param>
        /// <param name="sheetNum">读取的sheet编号，默认为0</param>
        /// <param name="rowLimit">行数限制，小于0表示无限制</param>
        /// <param name="columnLimit">列数限制，小于0表示无限制</param>
        /// <returns>读取的DataTable</returns>
        public static DataTable ReadAsDataTable(string filePath, int sheetNum = 0, int rowLimit = -1, int columnLimit = -1)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;
            FileStream fs = null;
            try
            {
                IWorkbook workbook = null;
                using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 判断文件格式
                    if (Path.GetExtension(filePath).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else if (Path.GetExtension(filePath).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        if (sheetNum >= 0 && sheetNum < workbook.NumberOfSheets)
                        {
                            var sheet = workbook.GetSheetAt(sheetNum);
                            if (sheet != null)
                            {
                                return GetDataTableBySheet(sheet, rowLimit, columnLimit);
                            }
                        }
                    }
                    return null;
                }
            }
            catch (Exception e)
            {
                ;//throw new DataSourceException(string.Format("读取【{0}】文件第【{1}】个子表失败",filePath,sheetNum));
            }
            finally
            {
                fs?.Close();
               
            }
            return null;
        }
        public static DataTable ReadAsDataTable(string filePath, string sheetName, int rowLimit = -1, int columnLimit = -1)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;
            FileStream fs = null;
            try
            {
                IWorkbook workbook = null;
                using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 判断文件格式
                    if (Path.GetExtension(filePath).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else if (Path.GetExtension(filePath).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        if (!string.IsNullOrEmpty(sheetName))
                        {
                            var sheet = workbook.GetSheet(sheetName);
                            if (sheet != null)
                            {
                                return GetDataTableBySheet(sheet, rowLimit, columnLimit);
                            }
                            else
                            {
                                ;
                                //throw new DataSourceException("找不到sheet: " + sheetName);
                            }
                        }
                    }

                    return null;
                }
            }
            catch (Exception e)
            {
                ;// throw new DataSourceException("Excel读取失败： "+ filePath, e);
            }
            finally
            {
                fs?.Close();
                
            }
            return null;
        }
        /// <summary>
        /// 根据列名字符串获得新的DataTable
        /// 列名不可以重复
        /// </summary>
        /// <param name="filePath">完整路径</param>
        /// <param name="sheetName">Sheet名</param>
        /// <param name="columnNames">列名字符串</param>
        /// <returns></returns>
        public static DataTable ReadAsDataTable(string filePath, string sheetName, List<string> columnNames)
        {
            if (sheetName == null)
            {
                return ReadAsDataTable(filePath, 0);
            }

            var sheets = GetSheetsByWorkBook(filePath);
            if (sheets == null)
            {
                return null;
            }

            DataTable dataTable = GetDataTableBySheet(sheets.Find(s => s.SheetName == sheetName));
            if (dataTable == null)
            {
                return null;
            }
            if (columnNames == null)
            {
                return ReadAsDataTable(filePath, sheetName);
            }

            DataTable table = new DataTable();
            if (dataTable.Rows.Count <= 0) return table;

            DataView view = new DataView(dataTable);
            try
            {
                //table = view.ToTable(dataTable.TableName, false, columnNames.ToArray());
                table = Select(dataTable, columnNames, dataTable.Rows.Count);
            }
            catch
            {
                ;// throw new DataSourceException("DataTable中的某些列，在Excel中不存在");
            }

            return table;
        }

        private static DataColumn GetDataColumnAsDiffName(DataTable newDataTable, string columnName, Type columnType, int index = 0)
        {
            DataColumn dc = new DataColumn(columnName, columnType);
            if (newDataTable.Columns.Contains(columnName))
            {
                dc.ColumnName += index.ToString();
                return GetDataColumnAsDiffName(newDataTable, dc.ColumnName, columnType, index + 1);
            }
            return dc;
        }
        /// <summary>
        /// 获取表格中所有的Sheet信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<ISheet> GetSheetsByWorkBook(string filePath)
        {
            List<ISheet> sheetList = new List<ISheet>();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;
            FileStream fs = null;
            try
            {
                IWorkbook workbook = null;
                using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 判断文件格式
                    if (Path.GetExtension(filePath).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else if (Path.GetExtension(filePath).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            sheetList.Add(workbook.GetSheetAt(i));
                        }
                        return sheetList;
                    }
                    return null;
                }
            }
            catch (Exception e)
            {
                sheetList.Clear();
                //throw new DataSourceException("打开文件失败： " + filePath, e);
            }
            finally
            {
                fs?.Close();
            }
            return null;
        }
        public static int GetSheetCount(string filePath)
        {
            List<ISheet> sheetList = new List<ISheet>();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return -1;
            FileStream fs = null;
            try
            {
                IWorkbook workbook = null;
                using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 判断文件格式
                    if (Path.GetExtension(filePath).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else if (Path.GetExtension(filePath).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        return workbook.NumberOfSheets;
                    }
                    return -1;
                }
            }
            catch (Exception e)
            {
                ; // throw new DataSourceException("打开文件失败： " + filePath, e);
            }
            finally
            {
                fs?.Close();
            }
            return 0;
        }

        /// <summary>
        /// 读取Excel表格中所有的Sheet
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="rowLimit">读取行数限制</param>
        /// <param name="columnLimit">读取列数限制</param>
        /// <returns>所有的Sheet集合</returns>
        public static DataSet ReadAsDataSet(string filePath, int rowLimit = -1, int columnLimit = -1)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;

            DataSet ds = new DataSet();
            FileStream fs = null;
            try
            {
                IWorkbook workbook = null;
                using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 判断文件格式
                    if (Path.GetExtension(filePath).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else if (Path.GetExtension(filePath).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        for (var i = 0; i < workbook.NumberOfSheets; ++i)
                        {
                            var sheet = workbook.GetSheetAt(i);
                            if (sheet != null)
                            {
                                ds.Tables.Add(GetDataTableBySheet(sheet, rowLimit, columnLimit));
                            }
                            else
                            {
                                ds.Tables.Add(new DataTable());
                            }
                        }
                    }
                    return ds;
                }
            }
            catch (Exception e)
            {
                ;// throw new DataSourceException("Excel读取失败： " + filePath, e);
            }
            finally
            {
                fs?.Close();
            }
            return null;
        }

        /// <summary>
        /// 由DataTable导出Excel
        /// </summary>
        /// <param name="sourceTable">要导出数据的DataTable</param>
        /// <param name="sheetName">工作薄名称</param>
        /// <param name="filePath">导出路径</param>
        /// <param name="colNames">需要导出的列名,可选</param>
        /// <param name="colAliasNames">导出的列名重命名，可选</param>
        /// <param name="colDataFormats">列格式化集合，可选</param>
        /// <param name="sheetSize">指定每个工作薄显示的记录数，可选（不指定或指定小于0，则表示只生成一个工作薄）</param>
        /// <returns></returns>
        public static void SaveDataTableToExcel(string filePath, DataTable sourceTable, string sheetName)
        {
            if (sourceTable.Rows.Count <= 0) return;

            if (string.IsNullOrEmpty(filePath)) return;

            bool isCompatible = GetIsCompatible(filePath);

            if (isCompatible)
            {
                SaveExcel_NPOI(filePath, sourceTable, sheetName);
            }
            else
            {
                SaveExcel_EPPlus(filePath, sourceTable, sheetName);
            }
        }

        /// <summary>
        /// 从工作表中生成DataTable
        /// 规则：
        /// 1.默认第一行为列名（不是字符型也当做字符处理）
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="headerRowIndex">表头</param>
        /// <param name="rowLimit">行数限制（负数表示无限制）</param>
        /// <param name="columnLimit">列数限制（负数表示无限制）</param>
        /// <returns></returns>
        private static DataTable GetDataTableBySheet(ISheet sheet, int rowLimit = -1, int columnLimit = -1, int headerRowIndex = 0)
        {
            try
            {
                var table = new DataTable();
                if (sheet == null) return null;
                if (sheet.LastRowNum <= 0) return null;
                if (sheet.FirstRowNum == sheet.LastRowNum) return null; //若只有表头无数据，返回空表
                //若首行为空，往下取第一个不为空的行做表头
                var headerRow = sheet.GetRow(headerRowIndex);
                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    if (headerRow != null && !IsEmptyRow(headerRow, headerRow.FirstCellNum, headerRow.LastCellNum))
                    {
                        break;
                    }
                    headerRowIndex++;
                    headerRow = sheet.GetRow(headerRowIndex);
                }
                if (headerRow == null || IsEmptyRow(headerRow, headerRow.FirstCellNum, headerRow.LastCellNum))
                {
                    return null;
                } //该sheet全是空行，返回空表

                //取表头下第一个有数据的行作为数据首行
                int firstDataRowIndex = headerRowIndex + 1;
                var firstDataRow = sheet.GetRow(firstDataRowIndex);
                for (int i = firstDataRowIndex; i <= sheet.LastRowNum; i++)
                {
                    if (firstDataRow != null && !IsEmptyRow(firstDataRow, headerRow.FirstCellNum, headerRow.LastCellNum)){break;}
                    firstDataRowIndex++;
                    firstDataRow = sheet.GetRow(firstDataRowIndex);
                }
                if (firstDataRow == null || IsEmptyRow(firstDataRow, headerRow.FirstCellNum, headerRow.LastCellNum)){return null;} //除表头行外全是空行，返回空表

                // 限制读取的列数
                var cellCount = columnLimit < 0
                    ? headerRow.LastCellNum
                    : Math.Min(headerRow.LastCellNum, headerRow.FirstCellNum + columnLimit);
                List<Type> colsType = new List<Type>();
                List<string> columnNames = new List<string>();
                for (var i = headerRow.FirstCellNum; i < cellCount; ++i)
                {
                    //遇到空列，填充为维度
                    if (headerRow.GetCell(i) == null)
                    {
                        headerRow.CreateCell(i, CellType.String).SetCellValue("column" + i.ToString());
                    }
                    if (string.IsNullOrEmpty(headerRow.GetCell(i).ToString()))
                    {
                        headerRow.CreateCell(i, CellType.String).SetCellValue("column" + i.ToString());
                    }
                    
                    //统一表头为字符型数据；经测试发现时间类型的表格也属于CellType.String
                    if (headerRow.GetCell(i)?.CellType != CellType.String)
                    {
                        headerRow.GetCell(i).SetCellType(CellType.String);
                    }

                    //列名相同时，加数字后缀
                    string columnTempName = headerRow.GetCell(i).StringCellValue;
                    string columnName = columnTempName;
                    columnNames.Add(columnName);
                    int j = 2;
                    while (columnNames.FindAll(x => x.Equals(columnName)).Count > 1)
                    {
                        columnName = (columnTempName + "(" + j + ")").ToString();
                        columnNames.Insert(i, columnName);
                        j++;
                    }

                    //将第一个单元格的数据类型作为整列的数据类型
                    var dataCell = firstDataRow.GetCell(i);
                    if (dataCell == null){dataCell = firstDataRow.CreateCell(i, CellType.String);}
                    colsType.Add(ParseCell(dataCell).GetType());
                    Type colType = colsType[i - headerRow.FirstCellNum];

                    var column = new DataColumn(columnName, colType);
                    table.Columns.Add(column);//设置DataTable
                }
                //Debug.LogWarning("列数cellCount" + cellCount + ",columns" + table.Columns.Count);

                // 读取行数据
                var rowCount = rowLimit < 0
                    ? sheet.LastRowNum
                    : Math.Min(sheet.LastRowNum - headerRowIndex, rowLimit) + headerRowIndex;
                //Debug.LogWarning("行数rowCount" + rowCount);

                for (var i = firstDataRowIndex; i <= rowCount; ++i)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)//如果遇到真空行
                    {
                        //continue; //跳过
                        row = sheet.CreateRow(i);//填充
                    }

                    var dataRow = table.NewRow();
                    for (int j = headerRow.FirstCellNum; j < cellCount; ++j)
                    {
                        var cell = row.GetCell(j);
                        if (cell == null)
                        {
                            cell = row.CreateCell(j);
                            //Debug.LogWarning("创建单元格"+ i + "行" + (j+1) + "列");
                        }

                        if (((ParseCell(cell).GetType() == table.Columns[j - headerRow.FirstCellNum].DataType) ||
                             (table.Columns[j - headerRow.FirstCellNum].DataType == typeof(string)))
                            && (!cell.ToString().Equals(""))) //若单元格数据不为空,且类型正确，正常读取，字符列默认全部读取
                        {
                            dataRow[j - headerRow.FirstCellNum] = cell.ToString();
                            //Debug.LogWarning(i + "行" + (j+1) + "列单元格被读成" + dataRow[j-headerRow.FirstCellNum]);
                        }
                        else //否则赋空值
                        {
                            dataRow[j - headerRow.FirstCellNum] = DBNull.Value;
                            //Debug.LogWarning("空单元格"+i + "行" + (j+1) + "列赋值为" + dataRow[j-headerRow.FirstCellNum]);
                        }
                    }

                    table.Rows.Add(dataRow); //填充DataTable
                }

                //SaveDataTableToExcel(@"f:\res.xls", table, "Sheet Test");
                //SaveDataTableToExcel(@"f:\res.xlsx", table, "Sheet Test");
                return table;
            }
            catch (Exception e)
            {
                ;// throw new DataSourceException("Excel文件读取失败",e);
            }
            return null;
        }

        /// <summary>
        /// 判断是否为空行（包括空行和行数据为空两种情况）
        /// </summary>
        /// <param name="row">需要判断的行</param>
        /// <param name="startCellNum">从此开始读数据</param>
        /// <param name="endCellNum">在此结束读数据</param>
        /// <returns></returns>
        private static bool IsEmptyRow(IRow row, int startCellNum, int endCellNum)
        {
            if (row == null)
            {
                //Debug.LogWarning("空行！");
                return true;
            }

            startCellNum = (startCellNum == null) ? row.FirstCellNum : startCellNum;
            if (startCellNum > endCellNum)
            {
                //Debug.LogError("数据错误！开始列大于结束列！");
            }

            for (int j = startCellNum; j <= endCellNum; j++)
            {
                var cell = row.GetCell(j);
                if (!string.IsNullOrEmpty(cell?.ToString()))//有一个单元格不为空
                {
                    return false;
                }
            }

            return true;
        }

        public static bool CheckExcelDataSource(string filePath, string sheetName, string columnName)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }
            FileStream fs = null;
            try
            {
                IWorkbook workbook = null;
                using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 判断文件格式
                    if (Path.GetExtension(filePath).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else if (Path.GetExtension(filePath).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        if (!string.IsNullOrEmpty(sheetName))
                        {
                            var sheet = workbook.GetSheet(sheetName);
                            if (sheet != null)
                            {
                                DataTable dataTable = ReadAsDataTable(filePath, sheetName);
                                if (dataTable.Columns.Contains(columnName))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                    return false;
                }
            }
            catch (Exception e)
            {
                ;// throw new DataSourceException("Excel文件读取失败", e);
            }
            finally
            {
                fs?.Close();
            }
            return false;
        }

        /// <summary>
        /// 依据值类型读取单元格数据
        /// </summary>
        /// <param name="cell">原始数据</param>
        private static object ParseCell(ICell cell)
        {
            if (cell == null) return "";
            object data = null;

            // 判断单元数据类型
            switch (cell.CellType)
            {
                case CellType.Blank:        // 空值
                    data = "";
                    break;
                case CellType.Numeric:      // 数值型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))
                    {
                        data = cell.DateCellValue;
                    }
                    else
                    {
                        data = cell.NumericCellValue;
                    }
                    break;
                case CellType.String:       // 字符串型
                    data = cell.StringCellValue;
                    break;
                case CellType.Boolean:      // 布尔型
                    data = cell.BooleanCellValue;
                    break;
                case CellType.Error:        // 错误
                    data = cell.ErrorCellValue;
                    break;
                case CellType.Formula:      // 公式型
                case CellType.Unknown:
                    data = cell.ToString();
                    break;
            }
            return data;
        }

        /// <summary>
        /// 统一列类型到数字或DateTime或字符型
        /// </summary>
        /// <returns></returns>
        private static Type UnifyColumnType(ISheet sheet, int rowStart, int colIdx)
        {
            int countNumeric = 0;
            int countDatetime = 0;
            //最多根据前100行判断
            int rowCount = Mathf.Min(rowStart + 100, sheet.LastRowNum);
            for (int i = rowStart; i < rowStart + rowCount; i++)
            {
                var row_i = sheet.GetRow(i);
                //跳过空行
                if (row_i == null)
                {
                    //Debug.LogWarning("第" + i + "为空行");
                    break;
                }

                var cell = row_i.GetCell(colIdx);
                if (cell == null)// 说明有空单元格
                {
                    cell = row_i.CreateCell(colIdx, CellType.Blank);
                }
                //Debug.LogWarning(i+"行"+(colIdx+1)+"列内容是："+cell.ToString());
                //Debug.LogWarning("单元格类型为："+cell.CellType);
                switch (cell.CellType)
                {
                    case CellType.String://只要该列有一个单元格是字符型，就将整列认为字符
                        return typeof(string);
                    case CellType.Numeric:
                        if (HSSFDateUtil.IsCellDateFormatted(cell))
                        {
                            countDatetime++;
                        }
                        else
                        {
                            countNumeric++;
                        }
                        break;
                    case CellType.Blank:
                        break;
                    default://其他类型也按字符处理
                        //todo:或直接报错，读取不了
                        return typeof(string);
                        break;
                }
            }

            if (countDatetime != 0)
            {
                if (countNumeric == 0){return typeof(DateTime);}//该列只有时间没有数字，设为时间型
                else{return typeof(string);}//该列既有时间又有数字，设为字符型
            }
            else
            {
                if (countNumeric != 0){return typeof(double);} //该列只有数字没有时间，设为double型
                else{return typeof(string);}//该列既无数字也无时间，说明是空列，设为字符型
            }

        }

        /// <summary>
        /// 判断是否为兼容模式
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool GetIsCompatible(string filePath)
        {
            string ext = Path.GetExtension(filePath);
            return new[] { ".xls", ".xlt" }.Count(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)) > 0;
        }

        /// <summary>
        /// 创建工作薄
        /// </summary>
        /// <param name="isCompatible"></param>
        /// <returns></returns>
        private static IWorkbook CreateWorkbook(bool isCompatible, string filePath)
        {
            IWorkbook workBook = null;
            if (isCompatible)
            {
                workBook = new HSSFWorkbook();
            }
            else
            {
                workBook = new XSSFWorkbook();
            }
            return workBook;
        }

        /// <summary>
        /// 依据值类型为单元格设置值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="colType"></param>
        public static void SetCellValue(ICell cell, string value, Type colType)
        {
            switch (colType.ToString())
            {
                case "System.String": //字符串类型
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(value);
                    break;
                case "System.DateTime": //日期类型
                    DateTime dateV;
                    if (DateTime.TryParse(value, out dateV))
                    {
                        cell.SetCellValue(dateV);
                    }
                    break;
                case "System.Boolean": //布尔型
                    bool boolV = false;
                    if (bool.TryParse(value, out boolV))
                    {
                        cell.SetCellType(CellType.Boolean);
                        cell.SetCellValue(boolV);
                    }
                    break;
                case "System.Int16": //整型
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    int intV = 0;
                    if (int.TryParse(value, out intV))
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(intV);
                    }
                    break;
                case "System.Decimal": //浮点型
                case "System.Double":
                    double doubV = 0;
                    if (double.TryParse(value, out doubV))
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(doubV);
                    }
                    break;
                case "System.DBNull": //空值处理
                    cell.SetCellType(CellType.Blank);
                    cell.SetCellValue("");
                    break;
                default:
                    cell.SetCellType(CellType.Unknown);
                    cell.SetCellValue(value);
                    break;
            }
        }

        /// <summary>
        /// 通过NPOI库保存*.xls
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sourceTable"></param>
        /// <param name="sheetName"></param>
        /// <param name="colNames"></param>
        /// <param name="colAliasNames"></param>
        /// <param name="colDataFormats"></param>
        /// <param name="sheetSize"></param>
        private static void SaveExcel_NPOI(string filePath, DataTable sourceTable, string sheetName)
        {
            var workbook = CreateWorkbook(true, filePath);
            var sheet = workbook.CreateSheet(sheetName);

            // 设置单元格式
            var style = workbook.CreateCellStyle();
            var font = workbook.CreateFont();
            font.FontName = "Arial";
            font.FontHeightInPoints = 11;
            style.SetFont(font);

            // 写入列名
            var row = sheet.CreateRow(0);
            ICell cell;
            for (var i = 0; i < sourceTable.Columns.Count; ++i)
            {
                cell = row.CreateCell(i);
                cell.CellStyle = style;
                cell.SetCellValue(sourceTable.Columns[i].ColumnName);
            }

            // 写入行数据
            for (int i = 0; i < sourceTable.Rows.Count; i++)
            {
                row = sheet.CreateRow(i + 1);
                for (int j = 0; j < sourceTable.Columns.Count; j++)
                {
                    var value = sourceTable.Rows[i][j];
                    cell = row.CreateCell(j);
                    cell.CellStyle = style;
                    SetCellValue(cell, value.ToString(), sourceTable.Columns[j].DataType);
                }
            }

            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workbook.Write(fs);
            fs.Dispose();
        }

        /// <summary>
        /// 通过EPPlus库保存*.xlsx
        /// TODO : NPOI库在Unity .net 3.5环境下无法保存*.xlsx，临时使用EPPlus库进行替代
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sourceTable"></param>
        /// <param name="sheetName"></param>
        private static void SaveExcel_EPPlus(string filePath, DataTable sourceTable, string sheetName)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                worksheet.Cells.LoadFromDataTable(sourceTable, true);

                package.SaveAs(new FileInfo(filePath));
            }
        }

        private static DataTable Select(DataTable srcTbl, List<string> cols, int maxRows = 200)
        {
            // 因DataTable.DefaultView.ToTable不能查询重复的列，因此只能手动生成
            try
            {
                int colCount = cols.Count;
                Dictionary<string, int> KeyDuplication = new Dictionary<string, int>(colCount);
                Func<string, string> GetColName = (key) =>
                {
                    if (KeyDuplication.ContainsKey(key))
                    {
                        KeyDuplication[key]++;
                        return key + KeyDuplication[key].ToString();
                    }
                    else
                    {
                        KeyDuplication.Add(key, 0);
                        return key;
                    }
                };

                var dstTbl = new DataTable("select" + srcTbl.TableName);
                Dictionary<string, int> srcIndices = new Dictionary<string, int>(colCount);
                for (int i = 0; i < srcTbl.Columns.Count; i++)// 把源表中的所有列名与index的映射记录在字典中
                {
                    srcIndices.Add(srcTbl.Columns[i].ColumnName, i);
                }
                List<int> dstIndices = new List<int>(colCount);// 用来记录目标表中每一列在源表中的列索引
                foreach (var name in cols)
                {
                    var idx = srcIndices[name];
                    dstIndices.Add(idx);
                    var column = new DataColumn(GetColName(name), srcTbl.Columns[idx].DataType);
                    dstTbl.Columns.Add(column);
                }

                int rowCnt = Mathf.Min(maxRows, srcTbl.Rows.Count);
                for (int i = 0; i < rowCnt; i++)
                {
                    var row = dstTbl.NewRow();
                    for (int j = 0; j < colCount; j++)
                    {
                        var colIdx = dstIndices[j];
                        row[j] = srcTbl.Rows[i][colIdx];
                    }
                    dstTbl.Rows.Add(row);
                }

                return dstTbl;
            }
            catch
            {
                ;// throw new DataSourceException("datatable select失败");
            }
            return null;
        }
    }
}