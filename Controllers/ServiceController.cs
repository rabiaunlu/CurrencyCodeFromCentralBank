using ExchangeRate.ClientMethods;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ExchangeRate.Controllers
{
    public class ServiceController : ApiController
    {
        [HttpGet]
        public void Call(string year, string month)
        {
            TCMBClient client = new TCMBClient();
            client.TCMBMetod("2020", "01", "01");
        }
    }
}