using System;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using TricentisExtensions.Modules.Helpers;

namespace TricentisExtensions.Modules.CustomEngine
{
    [SpecialExecutionTaskName("WriteToFile")]
    class WriteResultToFile : SpecialExecutionTask
    {
        public WriteResultToFile(Validator val) : base(val)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue sourceFilePath = testAction.GetParameterAsInputValue("TargetPath", false, new[] { ActionMode.Input });
            IInputValue sheetName = testAction.GetParameterAsInputValue("SheetName", false, new[] { ActionMode.Input });
            IInputValue columnName = testAction.GetParameterAsInputValue("ColumnName", false, new[] { ActionMode.Input });
            IInputValue dataSource = testAction.GetParameterAsInputValue("Data Source", true, new[] { ActionMode.Input });

            AOOperations ao;

            if (!String.IsNullOrEmpty(dataSource.Value))
                ao = new AOOperations(dataSource.Value);
            else
                ao = new AOOperations();

            var colName = columnName.Value;
            var filePath = sourceFilePath.Value;
            var sheet = sheetName.Value;

            ColumnType column;

            try
            {
                column = (ColumnType)Enum.Parse(typeof(ColumnType), colName);
            }
            catch (ArgumentException)
            {
                return new UnknownFailedActionResult("Invalid column name", $"Parse Failed with value: {colName}", colName);
                //throw ex;
            }

            using (var file = new FileHelper(filePath, sheet))
            {
                switch (column)
                {
                    case ColumnType.undefined:
                        break;
                    case ColumnType.Layout:
                        var layout = ao.GetRows();
                        layout.AddRange(ao.GetMeasures());
                        file.FillColumn("Default Layout", layout);
                        break;
                    case ColumnType.Dimensions:
                        var dimensions = ao.GetDimensions();
                        file.FillColumn("Dimensions", dimensions);
                        break;
                    case ColumnType.VarList:
                        var allPrompts = ao.GetVariables();
                        file.FillColumn("All Variables", allPrompts.Keys);
                        break;
                    case ColumnType.DefaultPrompts:
                        var prompts = ao.GetVariables("PROMPTS_FILLED");
                        file.FillColumn("Variable Name", prompts.Keys);
                        file.FillColumn("Variable Value", prompts.Values);
                        break;
                    case ColumnType.ActiveFilters:
                        var filters = ao.GetActiveFilters();
                        file.FillColumn("Filter Name", filters.Keys);
                        file.FillColumn("Filter Value", filters.Values);
                        break;
                    default:
                        break;
                }

            }

            return new PassedActionResult("Writer executed without exceptions, please check file");
        }

        private enum ColumnType
        {
            undefined,
            Layout,
            Dimensions,
            VarList,
            DefaultPrompts,
            ActiveFilters
        }
    }
}
