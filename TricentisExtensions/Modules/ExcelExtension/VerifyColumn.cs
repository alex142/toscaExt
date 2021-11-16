using System;
using System.Linq;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.Helpers;

namespace TricentisExtensions.Modules.ExcelExtension
{
    [SpecialExecutionTaskName("VerifyColumn")]
    class VerifyColumn : SpecialExecutionTask
    {
        public VerifyColumn(Validator val) : base(val)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue targetFilePath = testAction.GetParameterAsInputValue("Source Path", false, new[] { ActionMode.Input });
            IInputValue sourceFilePath = testAction.GetParameterAsInputValue("Target Path", false, new[] { ActionMode.Input });
            IInputValue targetSheetName = testAction.GetParameterAsInputValue("Source Sheet Name", false, new[] { ActionMode.Input });
            IInputValue sourceSheetName = testAction.GetParameterAsInputValue("Target Sheet Name", false, new[] { ActionMode.Input });
            IInputValue columnName = testAction.GetParameterAsInputValue("Column Name", false, new[] { ActionMode.Input });

            var targetTable = ExcelHelper.ConvertExcelToDataTable(targetFilePath.Value, targetSheetName.Value);
            var sourceTable = ExcelHelper.ConvertExcelToDataTable(sourceFilePath.Value, sourceSheetName.Value);

            var result = ExcelHelper.AreColumnsEqual(targetTable, sourceTable, columnName.Value);

            if (result == true) { return new PassedActionResult("Columns are matched"); }
            else { return new UnknownFailedActionResult("Columns aren't matched."); }
        }
    }
}
