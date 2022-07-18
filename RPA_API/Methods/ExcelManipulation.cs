using Microsoft.Extensions.Configuration;
using RPA_API.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RPA_API.Methods
{
    public class ExcelManipulation : IExcelManipulation
    {
        private IConfiguration Configuration { get; }
        private readonly IExcelUtility _excelutility;
        public ExcelManipulation(IConfiguration configuration, IExcelUtility excelUtility)
        {
            Configuration = configuration;
            _excelutility = excelUtility;
        }
        public async Task<ExcelResponse> ConnectToExcel(ExcelRequest excelRequest)
        {
            var excelresponse = new ExcelResponse();
            try
            {
                string path = await _excelutility.Createnewworkbook();                
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
            return excelresponse;
        }
    }
}
