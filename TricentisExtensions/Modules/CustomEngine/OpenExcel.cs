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
    [SpecialExecutionTaskName("OpenExcel")]
    class OpenExcel : SpecialExecutionTaskEnhanced
    {
        public OpenExcel(Validator validator) : base(validator)
        {

        }

        public override void ExecuteTask(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue filePath = testAction.GetParameterAsInputValue("File Path", false, new[] { ActionMode.Input });

            if (ProcessHelper.ProcessExists("EXCEL.EXE"))
            {
                ProcessHelper.EndProcess("EXCEL.EXE");
            }

            var xlApp = new FileHelper();
            try
            {
                xlApp.OpenExcel(filePath.Value);
                testAction.SetResult(SpecialExecutionTaskResultState.Ok, "Excel file is opened");
            }
            catch (Exception e)
            {
                testAction.SetResult(SpecialExecutionTaskResultState.Failed, $"Failed to open file:\n{e.Message}");
            }
        }
    }
}
