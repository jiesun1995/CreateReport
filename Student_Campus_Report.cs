using ConsoleApp1.Attributes.ReportAttribute;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class Student_Campus_Report
    {
        private ReportTemplate reportTemplate=new ReportTemplate();

        public ReportTemplate Template=>reportTemplate;

        public Student_Campus_Report()
        {
            reportTemplate = new ReportTemplate
            {
                Head = new List<RowProperty>() { new RowProperty { cells = new List<CellProperty> { new CellProperty { Name = "北区各栋各学院各年级人数及空床位统计表（2023.6.2）", ColSpan = 0 } } } },
                Body = new List<RowProperty>()
                {
                    new RowProperty
                    {
                        cells= new List<CellProperty>
                        {
                            new CellProperty{ Name="楼栋" },
                            new CellProperty{ Name="学院" },
                            new CellProperty{ Group = new GroupProperty{ Name="Level",Mode=true } },
                            new CellProperty{ Name="合计" },
                            new CellProperty{ Name="空床位数" },
                        },
                    },
                    new RowProperty
                    {
                        cells =new List<CellProperty>
                        {
                            new CellProperty
                            {
                                Group=new GroupProperty{ Name="BuildingName" ,Mode=false},
                                Cells=new List<CellProperty>
                                {
                                    new CellProperty
                                    {
                                        Group = new GroupProperty { Name = "DepaName", Mode = false },
                                        Cells=new List<CellProperty>
                                        {
                                            new CellProperty{ CellVal=new CellVal{ FunName="Sum",PropertyName="Freebed",Group=new GroupProperty{ Name="Type" } } },
                                            new CellProperty{ CellVal=new CellVal{ FunName="Sum",PropertyName="Freebed" } },
                                        }
                                    },
                                    new CellProperty{ Name="合计",NewRow =true },
                                    ///学院总计
                                    new CellProperty{ CellVal=new CellVal{ FunName="Sum",PropertyName="Freebed" }  },
                                    new CellProperty{ CellVal=new CellVal{ FunName="Sum",PropertyName="Freebed" }  },
                                    ///楼栋统计
                                    new CellProperty{ CellVal=new CellVal{ FunName="Sum",PropertyName="Freebed" }  },
                                }
                            },
                            //楼栋总计
                            //new CellProperty{ CellVal=new CellVal{ FunName="Sum",PropertyName="Freebed" } }
                        }
                    },
                    new RowProperty
                    {
                        cells=new List<CellProperty>
                        {
                            new CellProperty
                            {
                                Name="合计" ,
                                ColSpan=2 ,
                                Cells=new List<CellProperty>
                                {
                                    new CellProperty{  Group = new GroupProperty { Name = "Type", Mode = false } }
                                }
                            }
                        }
                    }

                }
            };
        }

        //public List<List<CellProperty>> Table { get; set; } = new List<List<CellProperty>>
        //{
        //    ///表头
        //    new List<CellProperty>{ new CellProperty{ Name= "北区各栋各学院各年级人数及空床位统计表（2023.6.2）",ColSpan=0 } },
        //    ///表列头
        //    new List<CellProperty>
        //    {
        //        new CellProperty{ Name="楼栋" },
        //        new CellProperty{ Name="学院" },
        //        new CellProperty{ GroupName="Level" },
        //        new CellProperty{ Name="合计" },
        //        new CellProperty{ Name="空床位数" },
        //    },
        //    ///表数据
        //    new List<CellProperty>
        //    {
        //        new CellProperty
        //        {
        //            Group=new GroupProperty{ Name="BuildingName" ,Mode=false},
        //            Cells=new List<CellProperty>
        //            {
        //                new CellProperty{ GroupName="DepaName" },
        //                new CellProperty{ PropertyName="Freebed" },
        //                new CellProperty{ PropertyName="Freebed",FunName="SUM" }
        //            }
        //        },
        //        new CellProperty{ PropertyName="Freebed",FunName="SUM" }
        //    },
        //    ///表合计
        //    new List<CellProperty>
        //    {
        //        new CellProperty
        //        { 
        //            Name="合计" ,
        //            ColSpan=2 ,
        //            Cells=new List<CellProperty>
        //            {
        //                new CellProperty{ GroupName="Type" }
        //            }
        //        }
        //    }
        //};
    }


    public class ReportTemplate
    {
        public List<RowProperty> Head { get; set; } = new List<RowProperty>();
        public List<RowProperty> Body { get; set; } = new List<RowProperty>();

    }
    public class RowProperty
    {
        public List<CellProperty> cells { get; set; } = new List<CellProperty>();
    }
    public class CellProperty
    {
        public string Name { get; set; }
        public GroupProperty Group { get; set; }
        public CellVal CellVal { get; set; }
        public List<CellProperty> Cells { get; set; }
        /// <summary>
        /// 为0时，已sheet最大列数
        /// </summary>
        public int ColSpan { get; set; } = 1;
        public int RowSpan { get; set; } = 1;
        public bool NewRow { get;set; }
    }

    public class GroupProperty
    {
        public string Name { get; set; }
        /// <summary>
        /// 行模式或者列模式
        /// False:行模式
        /// True:列模式
        /// </summary>
        public bool Mode { get;set; }
    }
    public class CellVal
    {
        public string PropertyName { get; set; }
        public string FunName { get; set; }
        public GroupProperty Group { set; get; }
    }
}
