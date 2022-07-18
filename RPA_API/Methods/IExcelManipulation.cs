using RPA_API.Models;
using System.Threading.Tasks;

namespace RPA_API.Methods
{
    public interface IExcelManipulation
    {
        Task<ExcelResponse> ConnectToExcel(ExcelRequest excelRequest);
    }
}