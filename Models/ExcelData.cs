using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    [EpplusTable()]
    public class ExcelData
    {
        public ExcelData()
        {
            this.Id = Id;
            this.Name = Name;
            this.IsTrue = IsTrue;
            this.Email = Email;
            this.Date = Date;
        }
        [EpplusTableColumn(Order = 1)]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 2)]
        public string Name { get; set; }

        [EpplusTableColumn(Order = 3)]
        public bool IsTrue { get; set; }

        [EpplusTableColumn(Order = 4)]
        public string Email { get; set; }

        [EpplusTableColumn(Order = 5)]
        public DateTime Date { get; set; }
    }
}
