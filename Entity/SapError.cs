using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.Entity
{
    class SapError : Exception
    {
        public SapError(string response) : base(get_message(response))
        {
        }


        public static string get_message(string response)
        {
            var @object = JObject.Parse(response);
            var token = @object["error"]["message"]["value"];
            return token.ToString();
        }
    }
}
