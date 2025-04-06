using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace automator_actioner
{
    internal class Output
    {
        public class _OutputValue
        {
            public string type { get; set; } = "";
            public string content { get; set; } = "";
        }

        private class _Output
        {
            public string errCode { get; set; } = "";
            public string errMsg { get; set; } = "";

            public _OutputValue outputValue { get; set; }
        }

        private _Output output;

        public Output(string errCode, string errMsg, string type, string content)
        {
            _OutputValue value = new _OutputValue();
            value.type = type;
            value.content = content;

            this.output = new _Output();
            this.output.errCode = errCode;
            this.output.errMsg = errMsg;
            this.output.outputValue = value;
        }

        public string toString()
        {
            return JsonSerializer.Serialize(this.output);
        }
    }
}
