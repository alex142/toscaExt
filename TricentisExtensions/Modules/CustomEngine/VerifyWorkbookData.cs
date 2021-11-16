using System;
using System.IO;
using System.Linq;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.Helpers;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("VerifyWorkbookData")]

    public class VerifyWorkbookData : SpecialExecutionTask
    {
        private readonly string tempFileOne = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Tosca Temp Files", "ref.xlsx");
        private readonly string tempFileTwo = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Tosca Temp Files", "src.xlsx");

        public VerifyWorkbookData(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue sourceFilePath = testAction.GetParameterAsInputValue("Source File Path", true, new[] { ActionMode.Input });
            IInputValue sourceSheetName = testAction.GetParameterAsInputValue("Source Sheet Name", true, new[] { ActionMode.Input });
            IInputValue targetFilePath = testAction.GetParameterAsInputValue("Target File Path", false, new[] { ActionMode.Input });
            IInputValue targetSheetName = testAction.GetParameterAsInputValue("Target Sheet Name", false, new[] { ActionMode.Input });
            IInputValue reportName = testAction.GetParameterAsInputValue("Report Name", false, new[] { ActionMode.Input });
            IInputValue rowsToSkip = testAction.GetParameterAsInputValue("Rows to Skip", false, new[] { ActionMode.Input });
            IInputValue columnsToCheck = testAction.GetParameterAsInputValue("Number of Columns", false, new[] { ActionMode.Input });

            string srcPath;
            string srcSheet;

            var refFile = targetFilePath.Value;
            var refSheet = targetSheetName.Value;

            var rowsCnt = Convert.ToInt32(rowsToSkip.Value);
            var colCnt = Convert.ToInt32(columnsToCheck.Value);

            //If Input Parameter is set
            if (sourceFilePath != null)
            {
                srcPath = sourceFilePath.Value;
                srcSheet = sourceSheetName.Value;
            }

            //Get current XL running instance
            else
            {
                var xlapp = new FileHelper();

                xlapp.TryGetInstance(reportName.Value);

                srcPath = xlapp.XlBook.FullName;

                srcSheet = xlapp.XlSheet.Name;

            }

            var sourceData = ExcelHelper.ConvertExcelToDataTable(srcPath, srcSheet);
            var refData = ExcelHelper.ConvertExcelToDataTable(refFile, refSheet);

            //Remove technical rows and columns
            ExcelHelper.DeleteDataTableRows(sourceData, 0, rowsCnt);
            ExcelHelper.DeleteDataTableColumns(sourceData, colCnt);

            sourceData = ExcelHelper.RemoveEmptyRows(sourceData);
            refData = ExcelHelper.RemoveEmptyRows(refData);

            //Fill in buffer files
            ExcelHelper.ExportToExcel(refData, tempFileOne);
            ExcelHelper.ExportToExcel(sourceData, tempFileTwo);

            //Get data from buffer files
            var xldt1 = ExcelHelper.ConvertExcelToDataTable(tempFileOne, "Sheet1");
            var xldt2 = ExcelHelper.ConvertExcelToDataTable(tempFileTwo, "Sheet1");

            //Compare data
            var srcToTarget = ExcelHelper.GetTablesDifference(xldt1, xldt2).Count();
            var targetTosrc = ExcelHelper.GetTablesDifference(xldt2, xldt1).Count();


            if (srcToTarget == 0 && srcToTarget == targetTosrc) { return new PassedActionResult($"Report data matched to refference data"); }
            else { return new UnknownFailedActionResult("Report data didn't match to refference"); }
        }
    }
}
