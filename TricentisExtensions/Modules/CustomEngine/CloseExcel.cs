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
    [SpecialExecutionTaskName("CloseExcel")]
    class CloseExcel : SpecialExecutionTask
    {
        private bool _save = true;

        public CloseExcel(Validator validator) : base(validator)
        {

        }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            IInputValue excelName = testAction.GetParameterAsInputValue("Excel Name", true, new[] { ActionMode.Input });
            IInputValue save = testAction.GetParameterAsInputValue("Save", true, new[] { ActionMode.Input });

            var file = new FileHelper();

            try
            {
                if (save != null)
                {
                    bool.TryParse(save.Value, out _save);
                }

                if (excelName != null)
                {                    

                    file.TryGetInstance(excelName.Value);

                    if (file.GetXlApp() != null)
                    {                       
                        file.XlBook.Close(_save);
                        file.XlBook = null;

                        file.Dispose();

                        return new PassedActionResult($"File {excelName.Value} successfully closed!");
                    }
                    return new UnknownFailedActionResult("No running excel processes found");
                }
                else
                {
                    var xlApp = file.GetActiveInstance();
                    int cnt = xlApp.Workbooks.Count;
                    foreach (Microsoft.Office.Interop.Excel.Workbook book in xlApp.Workbooks)
                    {                        
                        book.Close(_save);                        
                    }

                    file.Dispose();

                    return new PassedActionResult($"Closed {cnt} workbook(s)!");
                }

            }
            catch (Exception e)
            {
                return new UnknownFailedActionResult("Failed to close excel", e.Message, "");
            }
            finally
            {
                file.Dispose();
            }
        }
    }
}
