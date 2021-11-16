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
    [SpecialExecutionTaskName("CompareExcel")]

    class CompareExcel : SpecialExecutionTask
    {
        public CompareExcel(Validator validator) : base(validator)
        {

        }

        /*Task: 
            Compares two excel sheets. Path and sheet name provided by input parameters in tosca
            */
        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue targetFilePath = testAction.GetParameterAsInputValue("SourcePath", false, new[] { ActionMode.Input });
            IInputValue sourceFilePath = testAction.GetParameterAsInputValue("TargetPath", false, new[] { ActionMode.Input });
            IInputValue targetSheetName = testAction.GetParameterAsInputValue("SourceSheetName", false, new[] { ActionMode.Input });
            IInputValue sourceSheetName = testAction.GetParameterAsInputValue("TargetSheetName", false, new[] { ActionMode.Input });

            var targetTable = ExcelHelper.ConvertExcelToDataTable(targetFilePath.Value.ToString(), targetSheetName.Value.ToString());
            var sourceTable = ExcelHelper.ConvertExcelToDataTable(sourceFilePath.Value.ToString(), sourceSheetName.Value.ToString());

            var result = ExcelHelper.GetTablesDifference(targetTable, sourceTable);

            if (result.Count() == 0){ return new PassedActionResult("Datasets are matched");}
            else { return new UnknownFailedActionResult("Datasets aren't matched.", $"Count of missmatched rows: {result.Count()}","");}
        }
        
    }
}
