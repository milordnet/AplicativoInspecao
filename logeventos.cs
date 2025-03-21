using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public static class logeventos
    {
        public static void RegistraEventoLog(EventLog EventLog, string mensagem, EventLogEntryType tipo)
        {
            EventLog.WriteEntry(mensagem, tipo);
        }
    }
}
