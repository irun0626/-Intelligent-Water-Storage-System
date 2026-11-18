// 引入基础系统命名空间
using System;
// 集合类 List、Dictionary 等
using System.Collections.Generic;
// LINQ查询（这个文件里面实际没用到）
using System.Linq;
// 字符串编码
using System.Text;
// NPOI核心接口：IWorkbook、ISheet、IRow、ICell 都在这里
using NPOI.SS.UserModel;
// NPOI 处理 .xlsx（Excel2007及以后）
using NPOI.XSSF.UserModel;
// NPOI 处理 .xls（Excel2003）
using NPOI.HSSF.UserModel;
// 文件流操作 FileStream
using System.IO;
// DataTable，内存表，ADO.NET
using System.Data;
// 反射，用来读取实体类T的属性（List<T>导出excel那里用到）
using System.Reflection;

// 工具类命名空间
namespace Zhaoxi.Utils
{
    /// <summary>
    /// Excel操作帮助类，基于NPOI
    /// 功能：DataTable导出Excel、Excel读成DataTable、List实体导出Excel
    /// NPOI：不用依赖本机安装Office，纯代码读写Excel
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// 将DataTable数据导入到excel中（DataTable → Excel文件）
        /// </summary>
        /// <param name="data">要导出的内存DataTable数据</param>
        /// <param name="fileName">输出Excel完整路径+文件名</param>
        /// <param name="sheetName">Excel里面工作表Sheet名称</param>
        /// <param name="isColumnWritten">是否把DataTable的列名作为Excel表头写入第一行</param>
        /// <returns>成功写入的行数（包含表头行）；失败返回-1</returns>
        public static int DataTableToExcel(DataTable data, string fileName, string sheetName, bool isColumnWritten)
        {
            int i = 0;   // 行循环变量
            int j = 0;   // 列循环变量
            int count = 0; // 记录已经写入多少行，用来定位行号
            ISheet sheet = null;     // NPOI工作表对象（一页sheet）
            IWorkbook workbook = null; // NPOI工作簿对象（整个Excel文件）

            // 判断文件名是否包含 .xlsx 后缀 → 新版Excel
            if (fileName.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook();
            // 判断是否是 .xls 后缀 → 旧版Excel
            else if (fileName.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook();

            try
            {
                // 创建文件流：打开或新建文件，读写模式
                using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    if (workbook != null)
                    {
                        // 在工作簿里新建一个Sheet
                        sheet = workbook.CreateSheet(sheetName);
                    }
                    else
                    {
                        // 无法识别excel后缀，返回-1失败
                        return -1;
                    }

                    // 如果需要写入表头（列名）
                    if (isColumnWritten == true)
                    {
                        // 创建第0行（Excel第一行，表头）
                        IRow row = sheet.CreateRow(0);
                        // 循环所有DataTable列
                        for (j = 0; j < data.Columns.Count; ++j)
                        {
                            // 在该行创建单元格，写入列名
                            row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                        }
                        count = 1; // 表头占了第0行，下一行数据从行号1开始
                    }
                    else
                    {
                        count = 0; // 不写表头，数据从第0行开始
                    }

                    // 遍历DataTable每一行数据
                    for (i = 0; i < data.Rows.Count; ++i)
                    {
                        // 创建一行，行号为count
                        IRow row = sheet.CreateRow(count);
                        // 遍历该行所有列
                        for (j = 0; j < data.Columns.Count; ++j)
                        {
                            // 单元格写入内容，统一转字符串
                            row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                        }
                        ++count; // 行计数+1
                    }
                    // 把整个workbook写入文件流，保存到磁盘
                    workbook.Write(fs);
                    return count; // 返回一共写入多少行
                }
            }
            catch (Exception ex)
            {
                // 捕获异常，打印错误信息
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable（Excel文件 → DataTable）
        /// </summary>
        /// <param name="fileName">Excel完整路径文件名</param>
        /// <param name="sheetName">要读取的sheet名称</param>
        /// <param name="isFirstRowColumn">Excel第一行是否作为DataTable的列名</param>
        /// <returns>读取成功返回DataTable；异常返回null</returns>
        public static DataTable ExcelToDataTable(string fileName, string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable(); // 新建内存表，用来存放读到的数据
            int startRow = 0; // 从哪一行开始读取数据
            try
            {
                IWorkbook workbook = null;
                // 只读方式打开文件流
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    // 根据后缀创建对应workbook
                    if (fileName.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    else if (fileName.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (sheetName != null)
                    {
                        // 根据sheet名字获取工作表
                        sheet = workbook.GetSheet(sheetName);
                        // 如果找不到这个名字的sheet，默认取第一个sheet
                        if (sheet == null)
                        {
                            sheet = workbook.GetSheetAt(0);
                        }
                    }
                    else
                    {
                        // 没传sheet名，直接取第一个sheet
                        sheet = workbook.GetSheetAt(0);
                    }

                    if (sheet != null)
                    {
                        IRow firstRow = sheet.GetRow(0); // 获取第一行
                        int cellCount = firstRow.LastCellNum; // 该行最后一个单元格序号 = 总列数

                        // 如果第一行是表头
                        if (isFirstRowColumn)
                        {
                            // 循环第一行所有单元格，作为DataTable的列
                            for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                            {
                                ICell cell = firstRow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellValue = cell.StringCellValue;
                                    if (cellValue != null)
                                    {
                                        // 新建DataTable列，加到表里面
                                        DataColumn column = new DataColumn(cellValue);
                                        data.Columns.Add(column);
                                    }
                                }
                            }
                            // 表头读完，数据从下一行开始
                            startRow = sheet.FirstRowNum + 1;
                        }
                        else
                        {
                            // 第一行就是数据，不从表头跳过
                            startRow = sheet.FirstRowNum;
                        }

                        int rowCount = sheet.LastRowNum; // sheet最后一行行号
                        // 循环读取数据行
                        for (int i = startRow; i <= rowCount; ++i)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null) continue; // 空行跳过

                            DataRow dataRow = data.NewRow(); // 在DataTable新建一行
                            // 遍历该行单元格
                            for (int j = row.FirstCellNum; j < cellCount; ++j)
                            {
                                if (row.GetCell(j) != null)
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            data.Rows.Add(dataRow); // 把组装好的行加入DataTable
                        }
                    }
                    return data;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 重载版本：从文件流Stream读取Excel到DataTable
        /// 用处：上传文件场景（前端上传得到流，不用先存本地文件）
        /// </summary>
        /// <param name="fs">excel数据流</param>
        /// <param name="fileName">文件名（只用来判断后缀xlsx/xls）</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="isFirstRowColumn">首行是否表头</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(Stream fs, string fileName, string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                IWorkbook workbook = null;
                // 根据后缀创建workbook，流直接传入
                if (fileName.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum;

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null)
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 将List<T>实体列表导出到Excel（泛型导出）
        /// </summary>
        /// <typeparam name="T">要导出的实体类类型，比如日志类、采集记录类</typeparam>
        /// <param name="list">实体集合List<T></param>
        /// <param name="fileName">输出文件路径</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="colNames">字典：key=实体属性名，value=Excel显示表头文字</param>
        /// <returns>写入行数，失败-1</returns>
        public static int ListToExcel<T>(List<T> list, string fileName, string sheetName, Dictionary<string, string> colNames)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            IWorkbook workbook = null;
            Type type = typeof(T); // 获取泛型T的类型，反射用

            if (fileName.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook();

            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    if (workbook != null)
                    {
                        sheet = workbook.CreateSheet(sheetName);
                    }
                    else
                    {
                        return -1;
                    }
                    // 拿到字典里面的属性key集合（顺序就是导出列顺序）
                    List<string> keys = new List<string>(colNames.Keys);

                    if (colNames.Count > 0) // 有配置表头，写入表头行
                    {
                        IRow row = sheet.CreateRow(0);
                        for (j = 0; j < keys.Count; ++j)
                        {
                            // value是给用户看的中文表头
                            row.CreateCell(j).SetCellValue(colNames[keys[j]]);
                        }
                        count = 1; // 数据从第1行开始
                    }
                    else
                    {
                        count = 0;
                    }

                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum;
                    // 反射拿到T类所有属性
                    PropertyInfo[] props = type.GetProperties();

                    // 遍历每一个实体对象
                    for (i = 0; i < list.Count; ++i)
                    {
                        IRow row = sheet.CreateRow(count);
                        // 按keys顺序取属性值
                        for (j = firstRow.FirstCellNum; j < cellCount; ++j)
                        {
                            // 通过反射，根据属性名字拿到属性对象
                            var p = type.GetProperty(keys[j]);
                            // 读取这个实体上该属性的值
                            object val = p.GetValue(list[i]);
                            if (val == null)
                                val = ""; // null转空字符串
                            row.CreateCell(j).SetCellValue(val.ToString());
                        }
                        ++count;
                    }
                    workbook.Write(fs);
                    return count;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 创建IWorkbook工作簿对象
        /// </summary>
        /// <param name="exName">后缀：".xlsx" 或 ".xls"</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkBook(string exName)
        {
            IWorkbook workbook = null;
            if (exName == ".xlsx")
                workbook = new XSSFWorkbook();
            else if (exName == ".xls")
                workbook = new HSSFWorkbook();
            return workbook;
        }

        /// <summary>
        /// 创建ISheet工作表
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetName">sheet名字</param>
        /// <returns></returns>
        public static ISheet CreateSheet(IWorkbook workbook, string sheetName)
        {
            ISheet sheet = null;
            if (workbook != null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            return sheet;
        }
    }
}
