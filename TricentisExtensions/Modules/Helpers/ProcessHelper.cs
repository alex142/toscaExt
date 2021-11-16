using AutoIt;
using System.Threading;
using TricentisExtensions.Modules.WindowForms;

namespace TricentisExtensions.Modules.Helpers
{
    public static class ProcessHelper
    {

        public static bool ProcessExists(string process)
        {
            var res = AutoItX.ProcessExists(process);
            return true ? res != 0 : false;
        }

        public static int EndProcess(string process)
        {
            return AutoItX.ProcessClose(process);
        }

        public static void SapLogon(string client, string user, string password, int sec = 15)
        {
            (new Thread(() =>
            {
                var win = new Window("Logon to SAP BusinessObjects BI Platform", "", sec);
                win.Close();

                var logonWin = new LoginWindow(sec);
                logonWin.LogIn(client, user, password);

                var prompts = new PromptWindow(sec);
                prompts.Close();

            })).Start();
        }
    }
}
