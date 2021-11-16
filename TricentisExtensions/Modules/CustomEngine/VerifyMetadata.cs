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
    [SpecialExecutionTaskName("VerifyMetadata")]

    public class VerifyMetadata : SpecialExecutionTask
    {
        public VerifyMetadata(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue targetFilePath = testAction.GetParameterAsInputValue("SourcePath", false, new[] { ActionMode.Input });
            IInputValue sourceFilePath = testAction.GetParameterAsInputValue("TargetPath", false, new[] { ActionMode.Input });
            IInputValue targetSheetName = testAction.GetParameterAsInputValue("SourceSheetName", false, new[] { ActionMode.Input });
            IInputValue sourceSheetName = testAction.GetParameterAsInputValue("TargetSheetName", false, new[] { ActionMode.Input });
            IInputValue dataSource = testAction.GetParameterAsInputValue("Data Source", true, new[] { ActionMode.Input });

            AOOperations ao;

            if (!String.IsNullOrEmpty(dataSource.Value))
                ao = new AOOperations(dataSource.Value);
            else
                ao = new AOOperations();

            using (var res = new FileHelper(targetFilePath.Value, targetSheetName.Value))
            {
                var layout = ao.GetRows();
                layout.AddRange(ao.GetMeasures());
                var dimensions = ao.GetDimensions();
                var prompts = ao.GetVariables();

                res.FillColumn("Dimensions", dimensions);
                res.FillColumn("Default Layout", layout);
                res.FillColumn("Variable Name", prompts.Keys);
                res.FillColumn("Variable Value", prompts.Values);

                var isMatched = FileHelper.VerifyMetadata(targetFilePath.Value, sourceFilePath.Value, targetSheetName.Value, sourceSheetName.Value);

                if (isMatched) { return new PassedActionResult($"Metadata is correct"); }
                else { return new UnknownFailedActionResult("Metadata isn't correct."); }
            }
            
        }
    }
}
