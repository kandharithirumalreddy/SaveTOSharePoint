using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace SaveToSharepointWeb.Controllers
{
    public class TestServiceController : ApiController
    {
        // GET api/<controller>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "Testvalue1", "Testvalue2" };
        }

        // POST api/<controller>
        [HttpPost]
        public string Post([FromBody]object value)
        {
            return $"Testing {value} azure";
        }
    }
}