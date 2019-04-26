using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToJson
{
	class Program
	{
		static void Main(string[] args)
		{
			var paths = Directory.GetFiles(Directory.GetCurrentDirectory());
			for (var i = 0; i < paths.Length; i++)
			{
				var p = paths[i];
				if (p.EndsWith(".xlsx"))
				{
					Console.WriteLine("------------------------------------------------------------");
					Console.WriteLine("[{0}] {1}", i, p);
					Console.WriteLine("------------------------------------------------------------");
					ExportExcel(p);
					Console.WriteLine();
				}
			}

			Console.WriteLine("Finished {0} files", paths.Length);
			Console.WriteLine("press any key to quit");
			Console.ReadLine();
		}

		static void ExportExcel(string path)
		{
			StringBuilder sb = new StringBuilder("[\n");

			object missing = System.Reflection.Missing.Value;
			Excel.Application excel = new Excel.Application(); //lauch excel application
			excel.DisplayAlerts = false;
			excel.Visible = false;
			excel.UserControl = true;

			// 以只读的形式打开EXCEL文件
			Excel.Workbook wb = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
				missing, missing, missing, true, missing, missing, missing, missing, missing);
			//取得第一个工作薄
			Excel.Worksheet ws = (Excel.Worksheet) wb.Worksheets[1];

			//取得总记录行数   (包括标题列)
			int rowsCount = ws.UsedRange.Cells.Rows.Count; //得到行数
			int columnsCount = ws.UsedRange.Cells.Columns.Count; //得到列数

			for (int i = 3; i <= rowsCount; i++)
			{
				sb.Append("    {");
				for (int j = 1; j <= columnsCount; j++)
				{
					string fieldName = ((Excel.Range) ws.Cells[j][1]).Text;
					string fieldType = ((Excel.Range) ws.Cells[j][2]).Text;
					string fieldValue = ((Excel.Range) ws.Cells[j][i]).Text;
					Console.Write(fieldValue + "\t");
					sb.AppendFormat("\"{0}\":{1}", fieldName, ConvertValue(fieldType, fieldValue));
					if (j != columnsCount)
					{
						sb.Append(", ");
					}
				}

				Console.WriteLine();
				if (i != rowsCount)
				{
					sb.Append("},\n");
				}
				else
				{
					sb.Append("}\n");
				}
			}

			wb.Close();
			excel.Quit();
			sb.Append("]\n");

			path = path.Replace(".xlsx", ".json");
			File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
		}

		static string ConvertValue(string type, string value)
		{
			if (type == "string")
			{
				value = "\"" + value + "\"";
			}
			else if (type == "bool")
			{
				if (value == "FALSE")
				{
					return "false";
				}
				else if (value == "TRUE")
				{
					return "true";
				}
				else
				{
					return "false";
				}
			}
			else
			{
				if (string.IsNullOrEmpty(value))
				{
					return "0";
				}
			}

			return value;
		}
	}
}
