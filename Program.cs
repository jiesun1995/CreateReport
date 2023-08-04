using FastReport.Export.Image;
using FastReport;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Formula.Functions;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using System.Linq.Expressions;

namespace ConsoleApp1
{
    internal class Program
    {
        private static DataSet dataSet;
        private static readonly string outFolder;
        private static readonly string inFolder;

        static Program()
        {
            inFolder = Utils.FindDirectory("in");
            outFolder = Path.Combine(Directory.GetParent(inFolder).FullName, "out");
        }

        private static void CreateDataSet()
        {
            // create simple dataset with one table
            dataSet = new DataSet();

            DataTable table = new DataTable("Table");            
            dataSet.Tables.Add(table);

            table.Columns.Add("BuildingName", typeof(string));
            table.Columns.Add("DepaName", typeof(string));
            table.Columns.Add("Level", typeof(string));
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("Freebed", typeof(int));

            table.Rows.Add("31#","交通运输与物流学院","预科生","预科生", 23);
            table.Rows.Add("31#", "交通运输（专）","19级","专", 23);
            table.Rows.Add("31#", "交通运输（专）", "19级", "本", 23);
            table.Rows.Add("31#", "轨道技术学院","20级","本", 23);
            table.Rows.Add("31#", "机电学院（专）","21级","专", 23);
            table.Rows.Add("31#", "机电学院（专）", "21级", "本", 23);
            table.Rows.Add("31#", "体育学院","22级", "本", 23);
            table.Rows.Add("31#", "体育预科生","19级", "本", 23);
            table.Rows.Add("31#", "虚拟产业学院", "19级","本", 23);
            table.Rows.Add("31#", "电气培训班","20级","本", 23);

            table.Rows.Add("34#", "交通运输与物流学院", "预科生", "预科生", 23);
            table.Rows.Add("34#", "交通运输（专）", "19级", "专", 23);
            table.Rows.Add("34#", "交通运输（专）", "19级", "本", 23);
            table.Rows.Add("34#", "轨道技术学院", "20级", "本", 23);
            table.Rows.Add("34#", "机电学院（专）", "21级", "专", 23);
            table.Rows.Add("34#", "机电学院（专）", "21级", "本", 23);
            table.Rows.Add("34#", "体育学院", "22级", "本", 23);
            table.Rows.Add("34#", "体育预科生", "19级", "本", 23);
            table.Rows.Add("34#", "虚拟产业学院", "19级", "本", 23);
            table.Rows.Add("34#", "电气培训班", "20级", "本", 23);
        }

        private static List<Student_Building> CreateData()
        {
            var list=new List<Student_Building>
            {
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="交通运输与物流学院",
                    Level="预科生",
                    Type="专",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="交通运输（专）",
                    Level="19级",
                    Type="专",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="交通运输（本）",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="轨道技术学院",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                 new Student_Building
                {
                    BuildingName="31#",
                    DepaName="轨道技术学院",
                    Level="20级",
                    Type="专",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="机电学院（专）",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="体育学院",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="体育预科生",
                    Level="预科生",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="虚拟产业学院",
                    Level="预科生",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="31#",
                    DepaName="电气培训班",
                    Level="培训班",
                    Type="本",
                    Freebed=23,
                },

                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="交通运输与物流学院",
                    Level="预科生",
                    Type="专",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="交通运输（专）",
                    Level="19级",
                    Type="专",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="交通运输（本）",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building            
                {
                    BuildingName="32#",
                    DepaName="轨道技术学院",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="机电学院（专）",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="体育学院",
                    Level="20级",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="体育预科生",
                    Level="预科生",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="虚拟产业学院",
                    Level="预科生",
                    Type="本",
                    Freebed=23,
                },
                new Student_Building
                {
                    BuildingName="32#",
                    DepaName="电气培训班",
                    Level="培训班",
                    Type="本",
                    Freebed=23,
                },
            };
            return list;
        }

        static void Main(string[] args)
        {
            //Console.WriteLine("Welcome! \nThis demo shows how to:\n -create a simple data set\n -add it to the report" +
            //                  "\n -export raw report (.frx) to prepared one (.fpx).\nPress any key to proceed...");
            ////Console.ReadKey();
            //CreateDataSet();

            //// create report instance
            //Report report = new Report();

            //// load the existing report
            //report.Load(Path.Combine(inFolder, "test.frx"));

            //// register the dataset
            //report.RegisterData(dataSet);

            //// prepare the report
            //report.Prepare();

            //// save prepared report
            //if (!Directory.Exists(outFolder))
            //    Directory.CreateDirectory(outFolder);
            //report.SavePrepared(Path.Combine(outFolder, "Prepared Report.fpx"));

            //// export to image
            //FastReport.Export.OoXML.Excel2007Export excel2007Export=new FastReport.Export.OoXML.Excel2007Export();
            //excel2007Export.ShowProgress = false;
            //report.Export(excel2007Export, Path.Combine(outFolder, "report.xlsx"));
            ////ImageExport image = new ImageExport();
            ////image.ImageFormat = ImageExportFormat.Jpeg;
            ////report.Export(image, Path.Combine(outFolder, "report.jpg"));
            ////report.Export

            //// free resources used by report
            //report.Dispose();

            //Console.WriteLine("\nPrepared report and report exported as image have been saved into the 'out' folder.");
            //Console.WriteLine("Press any key to exit...");
            //Console.ReadKey();


            Core core = new Core();
            var report = new Student_Campus_Report();
            var data = CreateData();
            //data).Select(x => new { val = x. });
            Expression<Func<IGrouping<string, Student_Building>>> exp = (x) => new GroupValue { GroupName = x.Key, Value = x.Sum(y => y.Freebed) };

            MethodInfo method = typeof(Enumerable).GetMethods().Where(a => a.Name == "Sum" && a.GetParameters().Length == 2)
                                                  .FirstOrDefault().MakeGenericMethod(typeof(Student_Building));
            var valparameter = Expression.Parameter(typeof(Student_Building), "y");
            var parameter = Expression.Parameter(typeof(Student_Building), "x");
            var member = Expression.PropertyOrField(parameter, "Freebed");
            var lambda = Expression.Lambda<Func<Student_Building, int>>(member, parameter);
            Expression.Lambda<Func<Student_Building, int>>(lambda, valparameter);
            //Expression.Call(lambda, method);
            //var val = method.Invoke(null, new object[] { data, lambda.Compile() });
            core.ExportExcel(report.Template, data, "北区");
            Console.ReadLine();

        }
    }
    public class Student_Building
    {
        public string BuildingName { get; set; }
        public string DepaName { get; set; }
        public string Level { get; set; }
        public string Type { get; set; }
        public int Freebed { get; set; }
    }
}
