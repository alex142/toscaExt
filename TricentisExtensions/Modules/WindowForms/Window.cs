using AutoIt;

namespace TricentisExtensions.Modules.WindowForms
{
    public class Window
    {
        public string Title { get; set; }
        public string Text { get; set; }

        public Window(string title, string text = "", int sec = 10)
        {
            Title = title;
            Text = text;
            Wait(sec);
        }

        protected void Wait(int seconds)
        {
            var exists = AutoItX.WinWait(Title, Text, seconds);

            if (exists == 0)
            {
                throw new System.NullReferenceException($"Window {Title} is not found. Waited for {seconds} seconds");
            }
        }

        protected int WinExists(int seconds)
        {
            return AutoItX.WinWait(Title, Text, seconds);
        }

        public void Close()
        {
            AutoItX.WinClose(Title, Text);
        }

        public void Kill()
        {
            AutoItX.WinKill(Title, Text);
        }

        public void Activate()
        {
            AutoItX.WinActivate(Title, Text);
        }
        
    }
}
