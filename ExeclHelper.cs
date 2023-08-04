using ConsoleApp1.Attributes.ReportAttribute;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class ExeclHelper
    {
        private static ICellStyle GetTableTitleStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;

            IFont font = workbook.CreateFont();
            font.FontName = "微软雅黑";
            font.IsBold = true;
            font.FontHeightInPoints = 16;
            cellStyle.SetFont(font);
            return cellStyle;
        }
        private static ICellStyle GetRowTitleStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;

            cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightTurquoise.Index;
            cellStyle.FillPattern = FillPattern.SolidForeground;
            IFont colFont = workbook.CreateFont();
            colFont.FontName = "微软雅黑";
            colFont.IsBold = true;
            colFont.FontHeightInPoints = 12;
            cellStyle.SetFont(colFont);
            return cellStyle;
        }
        private static ICellStyle GetRowStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;

            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.FontName = "微软雅黑";
            cellStyle.SetFont(font);
            return cellStyle;
        }
        private static bool IsFundamental(Type type)
        {
            return type.IsPrimitive || type.IsEnum || type.Equals(typeof(string)) || type.Equals(typeof(DateTime));
        }
        private static int GetPropertyCount(object obj)
        {
            var count = 0;
            Type type = obj.GetType();
            if (type.IsGenericType)
            {
                var listVal = obj as IEnumerable<object>;
                if (listVal == null) return count;
                for (int j = 0; j < listVal.Count(); j++)
                {
                    var item = listVal.ToList()[j];
                    var val = GetPropertyCount(item);
                    count = val > count ? val : count;
                }
                return count-1;
            }
            var properties = type.GetProperties();
            count = properties.Count();
            foreach (var Property in properties)
            {
                if (Property.PropertyType.IsGenericType)
                {
                    var listVal = Property.GetValue(obj, null) as IEnumerable<object>;
                    if (listVal == null || listVal.Count() < 0) continue;
                    var val = GetPropertyCount(listVal.ToList()[0]);
                    count = val > count ? val : count;
                }
                else if (!IsFundamental(Property.PropertyType))///基础类型
                {
                    var item = Property.GetValue(obj, null) as object;
                    var val = GetPropertyCount(item);
                    count = val > count ? val : count;
                }
            }
            return count;
        }
        private static bool ContainsGenericByProperty(object item)
        {
            var listcount = item.GetType().GetProperties().Where(x =>
            {
                if (x.PropertyType.IsGenericType)
                {
                    var list = x.GetValue(item, null) as IEnumerable<object>;
                    if (list != null && list.Count() > 0 && list.ToList()[0].GetType() != typeof(string))
                        return true;
                }
                return false;
            }).Count();
            return listcount > 0;
        }
        private static bool SetCellWidth(ICell cell,ICell cellTitle,int colnum,ref Dictionary<int, double> colwids)
        {
            var width = cellTitle == null ? Encoding.UTF8.GetBytes(cell.StringCellValue).Length : Encoding.UTF8.GetBytes(cell.StringCellValue).Length > Encoding.UTF8.GetBytes(cellTitle.StringCellValue).Length ?
                           Encoding.UTF8.GetBytes(cell.StringCellValue).Length : Encoding.UTF8.GetBytes(cellTitle.StringCellValue).Length;
            if (!colwids.ContainsKey(colnum))
            {
                colwids.Add(colnum, width);
            }
            else
            {
                if (colwids[colnum] < width)
                    colwids[colnum] = width;
                else
                    return false;
            }
            return true;
        }
        private static void FillSheet(IWorkbook xsWorkbook,ISheet sheet,object obj,int colCount,bool isData,Dictionary<int,double> colwids,ref int rownum)
        {
            var type = obj.GetType();
            IRow row;
            IRow rowTitle=null;
            int colnum = 0;
            if (type.IsGenericType)
            {
                var listVal = obj as IEnumerable<object>;
                if (listVal == null) return;
                var first = true;
                for (int j = 0; j < listVal.Count(); j++)
                {
                    var item = listVal.ToList()[j];
                    if (!ContainsGenericByProperty(item))
                    {
                        FillSheet(xsWorkbook, sheet, item, colCount, first,colwids, ref rownum);
                        first = false;
                    }
                    else
                    {
                        FillSheet(xsWorkbook, sheet, item, colCount, false,colwids, ref rownum);
                    }
                }
                sheet.CreateRow(rownum++);
                return;
            }
            var properties = type.GetProperties();

            if (isData)
            {
                rowTitle = sheet.CreateRow(rownum++);
            }
            row = sheet.CreateRow(rownum++);
            
            for (int i = 0; i < properties.Count(); i++)
            {
                var property = properties[i];
                
                if (property.PropertyType.IsGenericType)
                {
                    var listVal = property.GetValue(obj, null) as IEnumerable<object>;
                    if (listVal == null) continue;
                    var first = true;
                    for (int j = 0; j < listVal.Count(); j++)
                    {
                        var item = listVal.ToList()[j];
                        if (!ContainsGenericByProperty(item))
                        {
                            FillSheet(xsWorkbook, sheet, item, colCount, first,colwids, ref rownum);
                            first = false;
                        }
                        else
                        {
                            FillSheet(xsWorkbook, sheet, item, colCount, false,colwids, ref rownum);
                        }
                    }
                    sheet.CreateRow(rownum++);
                }
                else if (!IsFundamental(property.PropertyType))///基础类型
                {
                    var item = property.GetValue(obj, null) as object;
                }
                else
                {
                    var tableNameAttributes = property.GetCustomAttributes(typeof(TableNameAttribute), true);
                    if (tableNameAttributes.Count() > 0)
                    {
                        CellRangeAddress titleRangeAddress = new CellRangeAddress(0, 0, 0, colCount);
                        sheet.AddMergedRegion(titleRangeAddress);
                        row = sheet.CreateRow(0);
                        row.CreateCell(0).SetCellValue(property.GetValue(obj).ToString());
                        row.GetCell(0).CellStyle = GetTableTitleStyle(xsWorkbook);
                        row.Height =(short)(row.GetCell(0).CellStyle.GetFont(xsWorkbook).FontHeight * 3);
                        rownum++;
                    }

                    var displayNameAttributes = property.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                    if (displayNameAttributes.Count() > 0)
                    {
                        row.CreateCell(colnum).SetCellValue(property.GetValue(obj).ToString());
                        row.GetCell(colnum).CellStyle=GetRowStyle(xsWorkbook);
                        if (rowTitle != null)
                        {
                            rowTitle.CreateCell(colnum).SetCellValue((displayNameAttributes[0] as DisplayNameAttribute).DisplayName.ToString());
                            rowTitle.GetCell(colnum).CellStyle = GetRowTitleStyle(xsWorkbook);
                            rowTitle.Height = (short)(row.GetCell(colnum).CellStyle.GetFont(xsWorkbook).FontHeight * 2);
                        }
                        if(SetCellWidth(row.GetCell(colnum), rowTitle?.GetCell(colnum),colnum, ref colwids))
                        {
                            sheet.SetColumnWidth(colnum, Convert.ToInt32(colwids[colnum]*300));
                        }
                        colnum++;
                    }
                }
            }
        }

        public static string ExportExcel(object obj,string sheetName)
        {
            IWorkbook xsWorkbook = new SXSSFWorkbook();
            var colcount = GetPropertyCount(obj);
            var sheet = xsWorkbook.CreateSheet(sheetName);
            int rownum = 0;
            Dictionary<int, double> colwids = new Dictionary<int, double>();
            FillSheet(xsWorkbook,sheet,obj,colcount,false, colwids, ref rownum);

            string path = $"/upload/excelFile/{DateTime.Now.ToString("yyyyMMdd")}";
            string savePath = AppDomain.CurrentDomain.BaseDirectory;
            string newName = $"";
            if (!Directory.Exists(savePath))
                Directory.CreateDirectory(savePath);
            string fileName = $"{DateTime.Now.ToString("yyyyMMddHHmmssffff")}.xlsx";
            var strFullName = savePath + "/" + fileName;  //存储位置+文件名
            using (FileStream file = new FileStream(strFullName, FileMode.Create))
            {
                xsWorkbook.Write(file);
                file.Close();
                return path + "/" + fileName;
            }
        }
    }
}
