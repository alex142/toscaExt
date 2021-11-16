using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;


namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("LogOff")]
    public class LogOff : SpecialExecutionTask
    {
        public LogOff(Validator val) : base(val)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            var aoApp = new AOOperations();
            var result = aoApp.LogOff();
            
            if (result == 1) { return new PassedActionResult($"Successfully logged off"); }
            else { return new UnknownFailedActionResult("Logoff attempt failed failed", $"Additional info: {aoApp.GetLastError()}", ""); }
        }
    }
}
