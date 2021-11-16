namespace TricentisExtensions.Modules.WindowForms
{
    public class PromptWindow : Window
    {
        public PromptWindow() : base("Prompts")
        {
        }

        public PromptWindow(int seconds) : base("Prompts", "", seconds)
        {
        }

        public void ClosePrompts()
        {
            Close();
        }
    }
}
