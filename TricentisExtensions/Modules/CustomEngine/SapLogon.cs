using System;
using System.IO;
using System.Linq;
using System.Threading;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.WindowForms;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("SapLogon")]
    class SapLogon : SpecialExecutionTask
    {
        public SapLogon(Validator val) : base(val)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue seconds = testAction.GetParameterAsInputValue("Seconds", true, new[] { ActionMode.Input });
            IInputValue client = testAction.GetParameterAsInputValue("Client", false, new[] { ActionMode.Input });
            IInputValue user = testAction.GetParameterAsInputValue("User Name", false, new[] { ActionMode.Input });
            IInputValue password = testAction.GetParameterAsInputValue("Password", false, new[] { ActionMode.Input });

            
            try
            {
                new Thread(() =>
                {
                    var sec = 15;

                    if (!string.IsNullOrEmpty(seconds.Value))
                    { sec = Convert.ToInt32(seconds.Value); }


                    var win = new Window("Logon to SAP BusinessObjects BI Platform", "", sec);
                    win.Close();

                    var logonWin = new LoginWindow(sec);
                    logonWin.LogIn(client.Value, user.Value, password.Value);

                    var prompts = new PromptWindow(sec);
                    prompts.Close();
                }).Start();

                return new PassedActionResult($"Logon Successful");
            }
            catch (Exception e)
            {
                return new UnknownFailedActionResult("Logon failed", e.StackTrace, "");
            }
            
        }
    }
}
