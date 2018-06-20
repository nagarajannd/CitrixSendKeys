using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace CitrixSendKeys
{
    public class CTXSendKeys : SimulateKeyPress
    {
        public void SendKeys(string typeintext, int delay)
        {
            InsertKeystroke(typeintext, delay);
        }
    }
}
