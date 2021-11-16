using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using Tricentis.Automation.Execution.Results;
using TricentisExtensions.Modules.Helpers;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName(V)]
    class EnableAddin : SpecialExecutionTask
    {
        private const string V = "EnableAddin";

        public EnableAddin(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue excelName = testAction.GetParameterAsInputValue("Excel Name", false, new[] { ActionMode.Input });
            IInputValue client = testAction.GetParameterAsInputValue("Client", false, new[] { ActionMode.Input });
            IInputValue userName = testAction.GetParameterAsInputValue("User Name", false, new[] { ActionMode.Input });
            IInputValue password = testAction.GetParameterAsInputValue("Password", false, new[] { ActionMode.Input });
            IInputValue progId = testAction.GetParameterAsInputValue("Addin ProgID", true, new[] { ActionMode.Input });

            var file = new FileHelper();

            try
            {
                file.TryGetInstance(excelName.Value);

                if (progId != null)
                {
                    file.EnableComAddin(client.Value, userName.Value, password.Value, progId.Value);
                }

                else
                {
                    file.EnableComAddin(client.Value, userName.Value, password.Value);
                }

                return new PassedActionResult($"Addin activated!");
            }
            catch (Exception e)
            {
                return new UnknownFailedActionResult("Failed to activate Addin", e.Message, "");
            }

        }
    }
}
