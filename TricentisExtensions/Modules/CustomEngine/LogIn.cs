using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("LogIn")]
    public class LogIn : SpecialExecutionTask
    {
        public LogIn(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue userName = testAction.GetParameterAsInputValue("UserName", false, new[] { ActionMode.Input });
            IInputValue password = testAction.GetParameterAsInputValue("Password", false, new[] { ActionMode.Input });
            IInputValue system = testAction.GetParameterAsInputValue("System", false, new[] { ActionMode.Input });
            IInputValue dataSource = testAction.GetParameterAsInputValue("Data Source", true, new[] { ActionMode.Input });

            AOOperations ao;

            if (!String.IsNullOrEmpty(dataSource.Value))
                ao = new AOOperations(dataSource.Value);
            else
                ao = new AOOperations();

            var logon = ao.LogIn(system.Value, userName.Value, password.Value);

            if (logon == 1) { return new PassedActionResult($"Successfully logged in"); }
            else { return new UnknownFailedActionResult("Logon attempt failed failed", $"Additional info: {ao.GetLastError()}", ""); }
        }
    }
}
