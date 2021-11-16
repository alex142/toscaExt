using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.Helpers;


namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("EndProcess")]

    class EndProcessInstance : SpecialExecutionTask
    {
        public EndProcessInstance(Validator val) : base(val)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue process = testAction.GetParameterAsInputValue("Process Name", false, new[] { ActionMode.Input });

            int result = 0;

            try
            {
                while (ProcessHelper.ProcessExists(process.Value))
                {
                    result = ProcessHelper.EndProcess(process.Value);
                }

                if (result == 1) { return new PassedActionResult("Process ended."); }
                else { return new PassedActionResult($"No active {process.Value} instances found"); }

            }
            catch (Exception e)
            {
                return new UnknownFailedActionResult(e.Message);
            }
           
        }
    }
}
