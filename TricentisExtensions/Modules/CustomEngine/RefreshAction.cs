using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("RefreshAction")]

    public class RefreshAction : SpecialExecutionTask
    {
        public RefreshAction(Validator validator) : base(validator)
        {
        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue action = testAction.GetParameterAsInputValue("Action", false, new[] { ActionMode.Input });
            
            var aoApp = new AOOperations();
            int result;
            if (action.Value.ToUpper().Equals("PAUSE"))
            {
                result = aoApp.PauseRefresh();
            }
            else if (action.Value.ToUpper().Equals("ACTIVATE"))
            {
                result = aoApp.ActivateRefresh();
            }
            else
            {
                result = 0;
            }

            if (result == 1) { return new PassedActionResult($"Refresh action : {action.Value} completed"); }
            else { return new UnknownFailedActionResult("Action failed", $"Additional info: {aoApp.GetLastError()}", ""); }
        }
    }
}
