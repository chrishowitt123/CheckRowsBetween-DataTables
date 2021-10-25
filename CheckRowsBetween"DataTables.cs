using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupsCareProviderPerDepartment
{
    class Program
    {
        static void Main(string[] args)
        {

            var dt = GetDataTableFromExcel(@"M:\MSG Open Episodes\Done\Current Episode_Final Appointment_Outcome Discharged\Copy of Current Episode_Final Appointment_Outcome Discharged_TEST1 - Copy.xlsx");

            dt.Columns.Remove("URN");
            dt.Columns.Remove("EpisodeNumber");
            dt.Columns.Remove("Deceased");
            dt.Columns.Remove("PatientSurname");
            dt.Columns.Remove("PatientFirstName");
            dt.Columns.Remove("Gender");
            dt.Columns.Remove("PaitentDOB");
            dt.Columns.Remove("AddressFirstLine");
            dt.Columns.Remove("AddressSecondLine");
            dt.Columns.Remove("PostCode");
            dt.Columns.Remove("EpisodeVisitStatus");
            dt.Columns.Remove("LastAppointmentDate");
            dt.Columns.Remove("LastAppointmentTime");
            dt.Columns.Remove("LastAppointmentOutcome");


            dt = dt.DefaultView.ToTable( /*distinct*/ true);

            DataView dv = dt.DefaultView;
            dv.Sort = "Group, CareProviderGroup, EpisodeSpecialty";
            var dtToFindDT = dv.ToTable();

            var toSearchDT = GetDataTableFromExcel(@"M:\MSG Open Episodes\Done\Test\Current Episodes_Same Speciality and Care Provider.xlsx");

            toSearchDT.Columns.Add("Group");


            string data = string.Empty;
            StringBuilder sb = new StringBuilder();
            foreach (DataRow row1 in dtToFindDT.Rows)
            {
                string toCompaire1 = row1["EpisodeCareProvider"].ToString()
                    + row1["EpisodeSpecialty"].ToString()
                    + row1["LastAppointmentLocationDescription"].ToString()
                    + row1["CareProviderGroup"].ToString();

                foreach (DataRow row2 in toSearchDT.Rows)
                {
                    string toCompaire2 = row2["EpisodeCareProvider"].ToString()
                    + row2["EpisodeSpecialty"].ToString()
                    + row2["LastAppointmentLocationDescription"].ToString()
                    + row2["CareProviderGroup"];

                    if(toCompaire2.Equals(toCompaire1))
                    {

                        row2["Group"] = row1["Group"];
                    }
                }
            }

            data = sb.ToString();
            Console.WriteLine(data);


            var xlsxFile = $@"M:\MSG Open Episodes\Done\Test\Test.xlsx";

            if (File.Exists(xlsxFile))
            {
                File.Delete(xlsxFile);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo(xlsxFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add($"UniqueRows");
                ws.Cells["A1"].LoadFromDataTable(toSearchDT, true);
                ws.Cells.AutoFitColumns();
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ws.View.FreezePanes(2, 1);
                package.Save();
                dt.Clear();
            }



        }
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable dt = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = dt.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return dt;
            }
        }
    }

}
