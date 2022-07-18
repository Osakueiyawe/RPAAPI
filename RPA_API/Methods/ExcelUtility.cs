using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelApp = Microsoft.Office.Interop.Excel;
namespace RPA_API.Methods
{
    public class ExcelUtility : IExcelUtility
    {
        private IConfiguration Configuration { get; set; }
        public ExcelUtility(IConfiguration configuration)
        {
            Configuration = configuration;
        }
        public async Task<string> Createnewworkbook()
        {
            string path = "";
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                string filelocation = Configuration.GetSection("excelfilelocation").Value;
                if (!Directory.Exists(filelocation))
                {
                    Directory.CreateDirectory(filelocation);
                }
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                ExcelApp.Workbook newworKbooK = excelApp.Workbooks.Add(Type.Missing);                
                ExcelApp.Worksheet newexcelSheet = (ExcelApp.Worksheet)newworKbooK.Sheets[1];
                newexcelSheet.Cells[1,1] = "NETWORK";
                newworKbooK.SaveAs(filelocation + "Excel" + ".xlsx");
                excelApp.Quit();                
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return path;
        }
        public async Task<bool> Atmtechnical1(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookatmtechnical1 = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetatmtechnical = (ExcelApp.Worksheet)excelBookatmtechnical1.Sheets[1];                
                excelBookatmtechnical1.Save();
                excelBookatmtechnical1.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> network(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBooknetwork = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetnetwork = (ExcelApp.Worksheet)excelBooknetwork.Sheets[1];
                excelBooknetwork.Save();
                excelBooknetwork.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> sysadmin(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBooksysadmin = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetsysadmin = (ExcelApp.Worksheet)excelBooksysadmin.Sheets[1];
                excelBooksysadmin.Save();
                excelBooksysadmin.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> esupport(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookesupport = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetesupport = (ExcelApp.Worksheet)excelBookesupport.Sheets[1];
                excelBookesupport.Save();
                excelBookesupport.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> basissupport1(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookbasissupport1 = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetbasissupport1 = (ExcelApp.Worksheet)excelBookbasissupport1.Sheets[1];
                excelBookbasissupport1.Save();
                excelBookbasissupport1.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> basissupport2(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookbasissupport1 = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetbasissupport1 = (ExcelApp.Worksheet)excelBookbasissupport1.Sheets[1];
                excelBookbasissupport1.Save();
                excelBookbasissupport1.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> atmtechnical2(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookatmtechnical2 = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetatmtechnical2 = (ExcelApp.Worksheet)excelBookatmtechnical2.Sheets[1];
                excelBookatmtechnical2.Save();
                excelBookatmtechnical2.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> datacentre(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookdatacentre = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetdatacentre = (ExcelApp.Worksheet)excelBookdatacentre.Sheets[1];
                excelBookdatacentre.Save();
                excelBookdatacentre.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }

        public async Task<bool> consolidatedreport(string path)
        {
            bool result = false;
            try
            {
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    LogError.Errhandler("Excel is not installed!!");
                }
                ExcelApp.Workbook excelBookconsolidatedreport = excelApp.Workbooks.Open(path);
                ExcelApp.Worksheet excelSheetconsolidatedreport = (ExcelApp.Worksheet)excelBookconsolidatedreport.Sheets[1];
                excelBookconsolidatedreport.Save();
                excelBookconsolidatedreport.Close();
                result = true;
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return result;
        }
    }
}
