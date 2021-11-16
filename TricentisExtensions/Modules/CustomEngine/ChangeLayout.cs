using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("ChangeLayout")]
    public class ChangeLayout : SpecialExecutionTask
    {
        public ChangeLayout(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue dimension = testAction.GetParameterAsInputValue("DimensionName", false, new[] { ActionMode.Input });
            IInputValue targetArea = testAction.GetParameterAsInputValue("TargetArea", false, new[] { ActionMode.Input });
            IInputValue dataSource = testAction.GetParameterAsInputValue("Data Source", true, new[] { ActionMode.Input });

            AOOperations ao;

            if (!String.IsNullOrEmpty(dataSource.Value))
                ao = new AOOperations(dataSource.Value);
            else
                ao = new AOOperations();

            int result;

            if (targetArea.Value.ToUpper().Equals("ROWS"))
            {
                result = ao.AddToRows(dimension.Value);
            }
            else if (targetArea.Value.ToUpper().Equals("COLUMNS"))
            {
                result = ao.AddToColumns(dimension.Value);
            }
            else if (targetArea.Value.ToUpper().Equals("REMOVE"))
            {
                result = ao.AddToFilters(dimension.Value);
            }
            else
            {
                return new UnknownFailedActionResult("Action failed", $"Incorrect value for Target Area parameter : {targetArea.Value}", targetArea.Value);
            }

            if (result == 1) { return new PassedActionResult($"Action completed, {dimension.Value} dimension is added to {targetArea.Value}"); }
            else { return new UnknownFailedActionResult("Action failed", $"Additional info: {ao.GetLastError()}", ""); }
        }
    }
}
