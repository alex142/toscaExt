using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.WindowForms;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("WaitForPrompt")]
    class WaitForAOPrompt : SpecialExecutionTask
    {
        public WaitForAOPrompt(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue _timeout = testAction.GetParameterAsInputValue("Timeout", false, new[] { ActionMode.Input });

            bool success = Int32.TryParse(_timeout.Value, out int timeout);

            if (success)
            {
                try
                {
                    var win = new PromptWindow(timeout);
                }
                catch (Exception e)
                {
                    return new UnknownFailedActionResult(e.Message);
                }

                return new PassedActionResult("Wait operation completed!");
            }
            else
            {
                return new UnknownFailedActionResult($"Incorrect value input for parameter Timeout : {_timeout.Value}");
            }
        }
    }
}
