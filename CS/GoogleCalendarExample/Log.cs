using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleCalendarExample {
    public static class Log {
        static Action<string> logAction;
        public static void Register(Action<string> logDelegate) {
            logAction = logDelegate;
        }

        public static void WriteLine(string message) {
            if (logAction == null)
                return;
            logAction(message);
        }
    }
}
