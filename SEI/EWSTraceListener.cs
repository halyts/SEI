using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace SEI
{
    class EWSTraceListener : ITraceListener
    {
        public void Trace(string traceType, string traceMessage)
        {
            //Logger.Log(traceType + ": " + traceMessage, Logger.LogSeverity.DebugEWS);
        }

    }
}
