using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.WindowForms;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("WaitWindowDisappear")]
    class WaitWindowDisappear : SpecialExecutionTask
    {
        public WaitWindowDisappear(Validator val) : base(val)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue _timeout = testAction.GetParameterAsInputValue("Timeout", false, new[] { ActionMode.Input });

            var timeout = Convert.ToInt32(_timeout.Value);

            try
            {
                var win = new CancelDialog();
                win.WaitWindowDisappear(timeout);
            }
            catch (NullReferenceException)
            {
                return new PassedActionResult("Window did not appear.");
            }
            catch (Exception e)
            {
                return new UnknownFailedActionResult(e.Message);
            }

            return new PassedActionResult("Wait completed!");
        }
    }
}
