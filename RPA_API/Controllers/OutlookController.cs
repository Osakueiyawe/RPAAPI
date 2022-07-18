using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using RPA_API.Methods;
using RPA_API.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RPA_API.Controllers
{
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class OutlookController : ControllerBase
    {
        private readonly IConnectToOutlook _connectToOutlook;
        private readonly IExcelManipulation _excelmanipulation;
        private readonly ILogger<OutlookController> _logger;
        public OutlookController(IConnectToOutlook connectToOutlook, IExcelManipulation excelmanipulation, ILogger<OutlookController> logger)
        {
            _connectToOutlook = connectToOutlook;
            _excelmanipulation = excelmanipulation;
            _logger = logger;
        }
        [HttpPost]
        public async Task<IActionResult> GetEmailDetails([FromBody] OutlookRequest logindetails)
        {
            if (ModelState.IsValid)
            {
                _logger.LogInformation("About to work");
                _logger.LogError("error message");
               
                var outlookresult = await _connectToOutlook.Outlookdetails(logindetails);
                return Ok(outlookresult);
            }
            else
            {
                var outlookresponse = new OutlookResponse();
                outlookresponse.responsemessage = "Wrong Request";
                return BadRequest(outlookresponse);
            }
        }

        [HttpPost]
        public async Task<IActionResult> ManipulateExcel([FromBody] ExcelRequest excelrequest)
        {
            if (ModelState.IsValid)
            {
                var excelresponse = await _excelmanipulation.ConnectToExcel(excelrequest);
                return Ok(excelresponse);
            }
            else
            {
                return BadRequest();
            }
        }

    }
}
