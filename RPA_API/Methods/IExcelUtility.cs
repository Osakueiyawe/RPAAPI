using System.Threading.Tasks;

namespace RPA_API.Methods
{
    public interface IExcelUtility
    {
        Task<string> Createnewworkbook();
        Task<bool> Atmtechnical1(string path);
        Task<bool> atmtechnical2(string path);
        Task<bool> basissupport1(string path);
        Task<bool> basissupport2(string path);
        Task<bool> consolidatedreport(string path);
        Task<bool> datacentre(string path);
        Task<bool> esupport(string path);
        Task<bool> network(string path);
        Task<bool> sysadmin(string path);
    }
}