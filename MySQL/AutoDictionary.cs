using Dapper;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace MySQL
{
    public class AutoDictionary
    {
        private static string _path = AppDomain.CurrentDomain.BaseDirectory;
        /// <summary>
        /// 程序启动
        /// </summary>
        public static void Start()
        {
            try
            {
                var constr = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
                var database = Regex.Match(constr, @"database=([^;]+)").Groups[1].Value;
                var db = new MySqlConnection(constr);
                var sql = "SELECT table_name AS '表名',table_comment AS '表说明' FROM information_schema.TABLES WHERE table_schema = @database;";
                var tables = db.Query<Table>(sql, new { database = database }).ToList(); ;
                var dict = new Dictionary<Table, List<Dict>>();
                foreach (var table in tables)
                {
                    sql = "SELECT ordinal_position AS '字段序号',column_name AS '字段名',CASE WHEN column_key <> '' THEN '√' ELSE '' END AS '标识',CASE WHEN column_key = 'PRI' THEN '√' ELSE '' END AS '主键',data_type AS '数据类型',character_maximum_length AS '占用字节数',numeric_precision AS '长度',numeric_scale AS '小数位数',CASE WHEN is_nullable = 'YES' THEN '√' ELSE '' END AS '允许空',column_default AS '默认值',column_comment AS '字段说明' FROM Information_schema.COLUMNS WHERE	table_schema = @database AND table_Name = @tableName;";
                    var columns = db.Query<Dict>(sql, new { database = database, tableName = table.表名 }).ToList();
                    dict.Add(table, columns);
                }
                #region 设置路径           
                for (var i = 0; i < 2; i++)
                {
                    var sb = new StringBuilder();
                    var list = _path.Split('\\').ToList();
                    list.Remove("");
                    list.RemoveAt(list.Count - 1);
                    foreach (var str in list)
                    {
                        sb.Append(str).Append("\\");
                    }
                    _path = sb.ToString();
                }
                _path += db.Database + ".xlsx";
                #endregion
                GeneratedForm(dict);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        private static void GeneratedForm(Dictionary<Table, List<Dict>> dict)
        {
            Application app = null;
            Workbook workBook = null;
            Worksheet worksheet = null;
            Worksheet sheet = null;
            Range range = null;
            try
            {
                app = new Application()
                {
                    Visible = true,
                    DisplayAlerts = false
                };
                app.Workbooks.Add(true);
                if (File.Exists(_path))
                {
                    File.Delete(_path);
                }
                workBook = app.Workbooks.Add(Missing.Value);
                app.Sheets.Add(Missing.Value, Missing.Value, dict.Count);
                worksheet = (Worksheet)workBook.Sheets[1];//数据库字典汇总表
                worksheet.Name = "数据库字典汇总表";
                worksheet.Cells[1, 1] = "数据库字典汇总表";
                worksheet.Cells[2, 1] = "编号";
                worksheet.Cells[2, 2] = "表英文名称";
                worksheet.Cells[2, 3] = "表中文名称";
                worksheet.Cells[2, 4] = "数据说明";
                worksheet.Cells[2, 5] = "表结构描述(页号)";
                var type = typeof(Dict);
                var properties = type.GetProperties();
                for (var i = 0; i < dict.Count; i++)
                {
                    var table = dict.ElementAt(i).Key;
                    var fields = dict.ElementAt(i).Value;
                    sheet = (Worksheet)workBook.Sheets[i + 2];//数据表
                    sheet.Name = $"{(101d + i) / 100:F}";
                    sheet.Cells[1, 1] = "数据库表结构设计明细";
                    sheet.Cells[2, 1] = $"表名：{table.表名}";
                    sheet.Cells[3, 1] = table.表说明;
                    for (var j = 0; j < properties.Count(); j++)
                    {
                        sheet.Cells[4, j + 1] = properties[j].Name;
                        for (var k = 0; k < fields.Count; k++)
                        {
                            sheet.Cells[k + 5, j + 1] = type.GetProperty(properties[j].Name).GetValue(fields[k], null);
                        }
                    }
                    worksheet.Cells[i + 3, 1] = i + 1;
                    worksheet.Cells[i + 3, 2] = table.表名;
                    worksheet.Cells[i + 3, 3] = table.表说明;
                    worksheet.Cells[i + 3, 4] = string.Empty;
                    worksheet.Cells[i + 3, 5] = $"表{sheet.Name}";
                    #region  数据表样式
                    range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[fields.Count + 4, properties.Count()]];//选取单元格
                    range.VerticalAlignment = XlVAlign.xlVAlignCenter;//垂直居中设置 
                    range.EntireColumn.AutoFit();//自动调整列宽
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;//所有框线 
                    range.Borders.Weight = XlBorderWeight.xlMedium;//边框常规粗细
                    range.Font.Name = "宋体";//设置字体 
                    range.Font.Size = 14;//字体大小  
                    range.NumberFormatLocal = "@";
                    range = sheet.Range[sheet.Cells[4, 1], sheet.Cells[fields.Count + 4, properties.Count()]];//选取单元格
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;//水平居中设置                   
                    range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, properties.Count()]];//选取单元格
                    range.Merge(Missing.Value);
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;//水平居中设置 
                    range.Font.Bold = true;//字体加粗
                    range.Font.Size = 24;//字体大小                           
                    range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, properties.Count()]];//选取单元格
                    range.Merge(Missing.Value);
                    range = sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, properties.Count()]];//选取单元格
                    range.Merge(Missing.Value);
                    #endregion                  
                }
                #region  汇总表样式             
                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[dict.Count + 2, 5]];//选取单元格
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;//水平居中设置 
                range.VerticalAlignment = XlVAlign.xlVAlignCenter;//垂直居中设置 
                range.ColumnWidth = 30;//设置列宽
                range.Borders.LineStyle = XlLineStyle.xlContinuous;//所有框线
                range.Borders.Weight = XlBorderWeight.xlMedium;//边框常规粗细  
                range.Font.Name = "宋体";//设置字体 
                range.Font.Size = 14;//字体大小 
                range.NumberFormatLocal = "@";
                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];//选取单元格
                range.Merge(Missing.Value);
                range.Font.Bold = true;//字体加粗
                range.Font.Size = 24;//字体大小 
                #endregion           
                sheet?.SaveAs(_path);
                worksheet.SaveAs(_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                workBook?.Close();
                app?.Quit();
                range = null;
                sheet = null;
                worksheet = null;
                workBook = null;
                app = null;
                GC.Collect();
            }
        }
    }
}
