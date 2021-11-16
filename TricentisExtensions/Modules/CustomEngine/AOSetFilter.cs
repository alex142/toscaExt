using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("AOSetFilter")]
    public class AOSetFilter : SpecialExecutionTask
    {
        public AOSetFilter(Validator validator) : base(validator)
        {
        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue dimName = testAction.GetParameterAsInputValue("DimensionName", false, new[] { ActionMode.Input });
            IInputValue dimleValue = testAction.GetParameterAsInputValue("DimensionValue", false, new[] { ActionMode.Input });
            IInputValue dataSource = testAction.GetParameterAsInputValue("Data Source", true, new[] { ActionMode.Input });

            AOOperations ao;

            if (!String.IsNullOrEmpty(dataSource.Value))
                ao = new AOOperations(dataSource.Value);
            else
                ao = new AOOperations();


            var result = ao.SetFilter(dimName.Value.ToString(), dimleValue.Value);

            if (result == 1) { return new PassedActionResult($"Filter value : {dimName.Value} is set to Dimension : {dimleValue.Value}"); }
            else { return new UnknownFailedActionResult("Failed to set value.", $"Additional info: {ao.GetLastError()}", ""); }
        }
    }
}
