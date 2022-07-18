using RPA_API.Models;
using System.Threading.Tasks;

namespace RPA_API.Methods
{
    public interface IConnectToOutlook
    {
        Task<OutlookResponse> Outlookdetails(OutlookRequest userdetails);
    }
}