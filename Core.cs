using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FastReport.Export.Dxf.Groups;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace ConsoleApp1
{
    public class Core
    {
        private int _colnum = -1;
        private int _rownum = -1;

        private Dictionary<int,int> _maxCols=new Dictionary<int, int>();
        private Dictionary<int,int> _maxRows=new Dictionary<int, int>();
        private Dictionary<int, Dictionary<string,string>> _colWhere = new Dictionary<int, Dictionary<string, string>>();
        private Dictionary<int,List<int>> _cloneCellNum = new Dictionary<int, List<int>>();
         
        private ICell CreateNewCol(IRow row, ref int colnum)
        {
            colnum++;
            var col = row.CreateCell(colnum);
            //Console.WriteLine($"创建一个单元格：{col.RowIndex}列{col.ColumnIndex}");
            if (_colnum < colnum) _colnum = colnum;
            return col;
        }
        private IRow CreateNewRow(ISheet sheet,int rownum)
        {
            rownum++;
            var row = sheet.CreateRow(rownum);
            if (_rownum < rownum) _rownum = rownum;
            return row;
        }

        private void MergedRegion(CellProperty cell,ISheet sheet, int startColIndex,int startRowIndex)
        {
            if (cell.ColSpan > 1 && cell.RowSpan > 1)
            {
                CellRangeAddress titleRangeAddress = new CellRangeAddress(startRowIndex, cell.RowSpan+startRowIndex-1, startColIndex, cell.ColSpan+startColIndex-1);
                sheet.AddMergedRegion(titleRangeAddress);
            }
            else if (cell.RowSpan > 1)
            {
                CellRangeAddress titleRangeAddress = new CellRangeAddress(startRowIndex, cell.RowSpan + startRowIndex-1, startColIndex, startColIndex);
                sheet.AddMergedRegion(titleRangeAddress);
            }
            else if(cell.ColSpan > 1)
            {
                CellRangeAddress titleRangeAddress = new CellRangeAddress(startRowIndex, startRowIndex, startColIndex, cell.ColSpan + startColIndex - 1);
                sheet.AddMergedRegion(titleRangeAddress);
            }
            if(cell.RowSpan == 0)
            {
                if (!_maxRows.ContainsKey(startRowIndex))
                {
                    _maxRows.Add(startRowIndex, 0);
                }
            }
            if(cell.ColSpan == 0)
            {
                if (!_maxCols.ContainsKey(startColIndex))
                {
                    _maxCols.Add(startColIndex, 0);
                }
            }
        }

        private void SetCellValue(ICell cell,string content)
        {
            if(cell.ColumnIndex==1&& cell.RowIndex == 2)
            {

            }
            Console.WriteLine($"写入数据：行{cell.RowIndex}列{cell.ColumnIndex}值{content}");
            cell.SetCellValue(content);
            
        }
        private void CellHandle<T>(CellProperty cell, ISheet sheet, IRow sheetRow, List<T> data, ref int colnum, Dictionary<string, string> rowwhere = null) where T : class
        {
            if (!string.IsNullOrEmpty(cell.Name))
            {
                SetCellValue(CreateNewCol(sheetRow, ref colnum), cell.Name);
                MergedRegion(cell, sheet, colnum, _rownum);
            }
            else if (cell.Group != null && !string.IsNullOrEmpty(cell.Group.Name))
            {
                ParameterExpression m_Parameter = Expression.Parameter(typeof(T), "x");
                MemberExpression member = Expression.PropertyOrField(m_Parameter, cell.Group.Name);
                var exprelamada = Expression.Lambda<Func<T, string>>(member, m_Parameter).Compile();

                var groupNames = data.GroupBy(exprelamada).Select(x => x.Key);
                var tempColnum = colnum;
                for (int l = 0; l < groupNames.Count(); l++)
                {
                    var groupName = groupNames.ToList()[l];
                    if (cell.Group.Mode)///列头
                    {
                        SetCellValue(CreateNewCol(sheetRow, ref colnum), groupName);
                        if (!_colWhere.ContainsKey(colnum))
                        {
                            _colWhere.Add(colnum, new Dictionary<string, string> { { cell.Group.Name, groupName } });

                        }
                        //else
                        //{
                        //    colWhere[colnum].Add("");
                        //}
                        if (!_cloneCellNum.ContainsKey(tempColnum))
                        {
                            _cloneCellNum.Add(tempColnum, new List<int> { colnum - 1 });
                        }
                        else
                        {
                            _cloneCellNum[tempColnum].Add(colnum - 1);
                        }

                    }
                    else///数据
                    {
                        var rowColnum = colnum;
                        if (l != 0)
                        {
                            sheetRow = CreateNewRow(sheet, _rownum);
                        }
                        rowwhere.Add(cell.Group.Name, groupName);
                        SetCellValue(CreateNewCol(sheetRow, ref rowColnum), groupName);

                        if (cell.Cells != null && cell.Cells.Count > 0)
                        {
                            for (int i = 0; i < cell.Cells.Count; i++)
                            {
                                if (cell.Cells[i].NewRow)
                                {
                                    sheetRow = CreateNewRow(sheet, _rownum);
                                }
                                if (_cloneCellNum.ContainsKey(rowColnum))
                                {
                                    foreach (var cloneCol in _cloneCellNum[rowColnum])
                                    {
                                        var temp = cloneCol;
                                        CellHandle(cell.Cells[i], sheet, sheetRow, data, ref temp, rowwhere);
                                        rowColnum = temp;
                                    }
                                }
                                else
                                {
                                    CellHandle(cell.Cells[i], sheet, sheetRow, data, ref rowColnum, rowwhere);
                                }
                            }
                        }
                        rowwhere.Remove(cell.Group.Name);
                    }
                }
            }
            else if (cell.CellVal != null && !string.IsNullOrEmpty(cell.CellVal.PropertyName))
            {
                var col = CreateNewCol(sheetRow, ref colnum);
                var sqlcolwhere = new Dictionary<string, string>();
                if (_colWhere.ContainsKey(colnum))
                    sqlcolwhere = _colWhere[colnum];

                ParameterExpression m_Parameter = Expression.Parameter(typeof(T), "x");
                MemberExpression member = null;
                Expression expRes = null;
                List<Expression> exp = new List<Expression>();
                ///列条件
                foreach (var item in sqlcolwhere)
                {
                    if (expRes == null)
                    {
                        member = Expression.PropertyOrField(m_Parameter, item.Key);
                        expRes = Expression.Equal(member, Expression.Constant(item.Value, member.Type));
                    }
                    else
                    {
                        member = Expression.PropertyOrField(m_Parameter, item.Key);
                        expRes = Expression.And(expRes, Expression.Equal(member, Expression.Constant(item.Value, member.Type)));
                    }

                }
                ///行条件
                foreach (var item in rowwhere)
                {
                    if (expRes == null)
                    {
                        member = Expression.PropertyOrField(m_Parameter, item.Key);
                        expRes = Expression.Equal(member, Expression.Constant(item.Value, member.Type));
                    }
                    else
                    {
                        member = Expression.PropertyOrField(m_Parameter, item.Key);
                        expRes = Expression.And(expRes, Expression.Equal(member, Expression.Constant(item.Value, member.Type)));
                    }
                }
                ///构建select
                if (member != null)
                {
                    var exprelamada = Expression.Lambda<Func<T, bool>>(expRes, m_Parameter).Compile();
                    var list = data.Where(exprelamada);
                    if (cell.CellVal.Group == null || string.IsNullOrEmpty(cell.CellVal.Group.Name))
                    {
                        MethodInfo method = typeof(Enumerable).GetMethods().Where(a => a.Name == cell.CellVal.FunName && a.GetParameters().Length == 2)
                                                    .FirstOrDefault().MakeGenericMethod(typeof(T));
                        var parameter = Expression.Parameter(typeof(T), "x");
                        member = Expression.PropertyOrField(parameter, cell.CellVal.PropertyName);
                        var lambda = Expression.Lambda<Func<T, int>>(member, parameter).Compile();
                        var val = method.Invoke(null, new object[] { list, lambda });
                        SetCellValue(col, val.ToString());
                    }
                    else
                    {
                        ParameterExpression valParameter = Expression.Parameter(typeof(T), "x");
                        MemberExpression valMember = Expression.PropertyOrField(m_Parameter, cell.CellVal.Group.Name);
                        var valExprelamada = Expression.Lambda<Func<T, string>>(member, m_Parameter).Compile();

                        Type groupType = typeof(IGrouping<string, T>);  // 注意 GroupBy 函数返回的类型
                        ParameterExpression pge = Expression.Parameter(groupType, "pg");
                        MemberExpression meKeyGender = Expression.MakeMemberAccess(pge, groupType.GetProperty("Key"));  // 获取其中的属性，与上面动态拼接 Select 相同           

                        Type groupByResultType = typeof(GroupValue);
                        MemberAssignment maGender = Expression.Bind(groupByResultType.GetProperty(nameof(GroupValue.GroupName)), meKeyGender);  // 使用 Bind 方法将目标类型的属性与源类型的属性值绑定，与上面动态拼接 Select 相同
                        MethodInfo funMethod = typeof(Enumerable).GetMethods().Where(a => a.Name == cell.CellVal.FunName && a.GetParameters().Length == 2)
                                                                    .FirstOrDefault().MakeGenericMethod(typeof(T));    // 获取 Count 方法
                        var parameter = Expression.Parameter(typeof(T), "y");
                        var funmember = Expression.PropertyOrField(parameter, cell.CellVal.PropertyName);
                        var vallambda= Expression.Lambda<Func<T, int>>(funmember, parameter);
                        var funlambda = Expression.Lambda<Func<T, Func<T,int>>>(vallambda, pge);
                        
                        //var vallambda = Expression.Lambda<Func<T, int>>(funmember, parameter);
                        //var valLambda = Expression.Call(funMethod, funlambda);
                        MemberAssignment maCount = Expression.Bind(groupByResultType.GetProperty(nameof(GroupValue.Value)), funlambda);  //使用 Bind 方法将目标类型的属性与源类型调用方法的返回值绑定

                        NewExpression ne = Expression.New(groupByResultType);
                        MemberInitExpression mie = Expression.MemberInit(ne, maGender, maCount);

                        Expression<Func<IGrouping<string, T>, GroupValue>> personSelectExpression =
                            Expression.Lambda<Func<IGrouping<string, T>, GroupValue>>(mie, pge);
                        var personList1 = list.GroupBy(valExprelamada).Select(personSelectExpression.Compile()).ToList();
                        StringBuilder sb = new StringBuilder();
                        foreach (var person in personList1)
                        {

                            //var template = Mustachio.Parser.Parse(cell.CellVal.);

                        }
                    }
                }
            }
        }
        private void FillSheet<T>(ReportTemplate template, List<T> data,ISheet sheet) where T:class
        {
            ///报表头部
            for (int i = 0; i < template.Head.Count; i++)
            {
                var sheetRow = CreateNewRow(sheet,_rownum);
                var row = template.Head[i];
                var colnum = -1;
                for (int j = 0; j < row.cells.Count; j++)
                {
                    var cell = row.cells[j];
                    CellHandle(cell, sheet, sheetRow, data, ref colnum);
                }
            }

            ///报表数据
            for (int i = 0; i < template.Body.Count; i++)
            {
                var sheetRow = CreateNewRow(sheet, _rownum);
                var row = template.Body[i];
                var colnum = -1;
                var rowwhere = new Dictionary<string, string>();
                for (int j = 0; j < row.cells.Count; j++)
                {
                    var cell = row.cells[j]; 
                    CellHandle(cell, sheet, sheetRow, data, ref colnum, rowwhere);
                    if (j + 1 <= row.cells.Count)
                        j = cell.ColSpan + j - 1; 
                }
            }


        }

        public string ExportExcel<T>(ReportTemplate template, List<T> data, string sheetName)where T : class
        {
            IWorkbook xsWorkbook = new SXSSFWorkbook();
            var sheet = xsWorkbook.CreateSheet(sheetName);

            //try
            //{
            FillSheet<T>(template, data, sheet);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.StackTrace);
            //}

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
    public class GroupValue
    {
        public string GroupName { get; set; }
        public int Value { get; set; }
    }
}
