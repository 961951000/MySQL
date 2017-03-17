using Dapper;
using MySql.Data.MySqlClient;
using MySQL.Util;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MySQL.Models;

namespace MySQL
{

    public class AutoCreateModels
    {
        /// <summary>
        /// 程序启动
        /// </summary>
        public static int Start()
        {
            try
            {
                var constr = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
                var database = Regex.Match(constr, @"database=([^;]+)").Groups[1].Value;
                var db = new MySqlConnection(constr);
                var sql = "SELECT table_name AS '表名',table_comment AS '表说明' FROM information_schema.TABLES WHERE table_schema = @database";
                var tables = db.Query<Table>(sql, new { database = database }).ToList(); ;
                var dict = new Dictionary<Table, List<Dict>>();
                foreach (var table in tables)
                {
                    sql = "SELECT ordinal_position AS '字段序号',column_name AS '字段名',CASE WHEN column_key <> '' THEN '√' ELSE '' END AS '标识',CASE WHEN column_key = 'PRI' THEN '√' ELSE '' END AS '主键',data_type AS '数据类型',character_maximum_length AS '占用字节数',numeric_precision AS '长度',numeric_scale AS '小数位数',CASE WHEN is_nullable = 'YES' THEN '√' ELSE '' END AS '允许空',column_default AS '默认值',column_comment AS '字段说明' FROM Information_schema.COLUMNS WHERE table_schema = @database AND table_Name = @tableName";
                    var columns = db.Query<Dict>(sql, new { database = database, tableName = table.表名 }).ToList();
                    dict.Add(table, columns);
                }
                var count = CreateModel(dict);
                Console.WriteLine("程序执行成功，共创建{0}个模型", count);
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return 0;
            }
        }

        private static int CreateModel(Dictionary<Table, List<Dict>> dict)
        {
            var space = ConfigurationManager.AppSettings["modelnamespace"];
            var modelsPath = ConfigurationManager.AppSettings["path"];
            if (string.IsNullOrEmpty(space))
            {
                space = "Default.Models";
            }
            if (string.IsNullOrEmpty(modelsPath) || !BaseTool.IsValidPath(modelsPath))
            {
                modelsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Models");
            }
            if (!Directory.Exists(modelsPath))
            {
                Directory.CreateDirectory(modelsPath);
            }
            var count = 0;
            foreach (var table in dict)
            {
                var sb = new StringBuilder();
                var sb1 = new StringBuilder();
                var className = string.Empty;
                if (table.Key.表名.LastIndexOf('_') != -1)
                {
                    foreach (var str in table.Key.表名.Split('_'))
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            className += str.Substring(0, 1).ToUpper() + str.Substring(1).ToLower();
                        }
                    }
                }
                else
                {
                    className = table.Key.表名.Substring(0, 1).ToUpper() + table.Key.表名.Substring(1).ToLower();
                }
                var firstLetter = className.Substring(0, 1);
                if (firstLetter != "_" && !RegexTool.IsLetter(firstLetter))
                {
                    className = $"_{className}";
                }
                sb.Append("using System;\r\nusing System.ComponentModel.DataAnnotations;\r\nusing System.ComponentModel.DataAnnotations.Schema;\r\n\r\nnamespace ");
                sb.Append(space);
                sb.Append("\r\n{\r\n");
                if (!string.IsNullOrEmpty(table.Key.表说明))
                {
                    sb.Append("\t/// <summary>\r\n");
                    sb.Append("\t/// ").Append(table.Key.表说明).Append("\r\n");
                    sb.Append("\t/// </summary>\r\n");
                }
                sb.Append("\t[Table(\"").Append(table.Key.表名).Append("\")]\r\n");  //数据标记
                sb.Append("\tpublic class ");
                sb.Append(className);
                sb.Append("\r\n\t{\r\n");
                sb.Append("\t\t#region Model\r\n");
                var order = 0;
                foreach (var column in table.Value)
                {
                    var propertieName = string.Empty;
                    if (column.字段名.LastIndexOf('_') != -1)
                    {
                        foreach (var str in column.字段名.Split('_'))
                        {
                            if (!string.IsNullOrEmpty(str))
                            {
                                propertieName += str.Substring(0, 1).ToUpper() + str.Substring(1).ToLower();
                            }
                        }
                    }
                    else
                    {
                        propertieName = column.字段名.Substring(0, 1).ToUpper() + column.字段名.Substring(1).ToLower();
                    }
                    if (propertieName == className)
                    {
                        propertieName = $"_{propertieName}";
                    }
                    else
                    {
                        firstLetter = propertieName.Substring(0, 1);
                        if (firstLetter != "_" && !RegexTool.IsLetter(firstLetter))
                        {
                            propertieName = $"_{propertieName}";
                            if (propertieName == className)
                            {
                                propertieName = $"_{propertieName}";
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(column.字段说明))
                    {
                        sb.Append("\t\t/// <summary>\r\n");
                        sb.Append("\t\t/// ").Append(column.字段说明).Append("\r\n");
                        sb.Append("\t\t/// </summary>\r\n");
                    }
                    if (!string.IsNullOrEmpty(column.主键))
                    {
                        sb.Append("\t\t[Key, Column(\"").Append(column.字段名).Append("\", Order = ").Append(order).Append(")]\r\n");
                        order++;
                    }
                    else
                    {
                        sb.Append("\t\t[Column(\"").Append(column.字段名).Append("\")]\r\n");
                    }
                    if (string.IsNullOrEmpty(column.数据类型))
                    {
                        sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                    }
                    else
                    {
                        switch (column.数据类型.ToLower())
                        {
                            case "tinyint":
                                {
                                    sb.Append("\t\tpublic bool? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "smallint":
                                {
                                    sb.Append("\t\tpublic short? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "int":
                                {
                                    sb.Append("\t\tpublic int? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "bigint":
                                {
                                    sb.Append("\t\tpublic long? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "decimal":
                                {
                                    sb.Append("\t\tpublic decimal? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "timestamp":
                                {
                                    sb.Append("\t\tpublic DateTime? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "datetime":
                                {
                                    sb.Append("\t\tpublic DateTime? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "bit":
                                {
                                    sb.Append("\t\tpublic bool " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "money":
                                {
                                    sb.Append("\t\tpublic decimal? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "image":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "nvarchar":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "varchar":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "text":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            default:
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                        }
                    }
                    sb.Append("\t\t\tset;\r\n");
                    sb.Append("\t\t\tget;\r\n");
                    sb.Append("\t\t}\r\n");
                    sb.Append("\r\n");
                    sb1.Append(propertieName);
                    sb1.Append("=\" + ");
                    sb1.Append(propertieName);
                    sb1.Append(" + \",");
                }
                sb1.Remove(sb1.Length - 5, 5);
                sb.Append("\t\tpublic override string ToString()\r\n");
                sb.Append("\t\t{\r\n");
                sb.Append("\t\t\treturn \"");
                sb.Append(sb1);
                sb.Append(";");
                sb.Append("\r\n");
                sb.Append("\t\t}\r\n");
                sb.Append("\t\t#endregion Model\r\n");
                sb.Append("\t}\r\n").Append("}");
                var filePath = Path.Combine(modelsPath, $"{className}.cs");
                if (WriteFile(filePath, sb.ToString()))
                {
                    count++;
                }
            }
            return count;
        }
        /// <summary>
        /// 文件写入
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="text">文本内容</param>
        public static bool WriteFile(string filePath, string text)
        {
            var flag = false;
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                if (!File.Exists(filePath))
                {
                    // 创建写入文件
                    fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);
                    sw.WriteLine(text);

                }
                else
                {
                    // 删除文件在创建
                    File.Delete(filePath);
                    fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);
                    sw.WriteLine(text);
                }
                flag = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                sw?.Flush();
                sw?.Close();
                fs?.Close();
            }
            return flag;
        }
    }
}
