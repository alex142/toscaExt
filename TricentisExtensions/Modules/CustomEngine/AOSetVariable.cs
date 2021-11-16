using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("AOSetVariable")]
    public class AOSetVariable : SpecialExecutionTask
    {
        public AOSetVariable(Validator validator) : base(validator)
        {
        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue variableName = testAction.GetParameterAsInputValue("VariableName", false, new[] { ActionMode.Input });
            IInputValue variableValue = testAction.GetParameterAsInputValue("VariableValue", false, new[] { ActionMode.Input });
            IInputValue dataSource = testAction.GetParameterAsInputValue("Data Source", true, new[] { ActionMode.Input });

            AOOperations ao;

            if (!String.IsNullOrEmpty(dataSource.Value))
                ao = new AOOperations(dataSource.Value);
            else
                ao = new AOOperations();

            var result = ao.SetVariable(variableName.Value.ToString(), variableValue.Value);

            if (result == 1) { return new PassedActionResult($"Value : {variableValue.Value} is set to Variable : {variableName.Value}"); }
            else { return new UnknownFailedActionResult("Failed to set value.", $"Additional info: {ao.GetLastError()}", ""); }
        }

    }
}
