using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoIt;

namespace TricentisExtensions.Modules.WindowForms
{
    public class Control
    {
        private string Name { get; set; }
        private int ID { get; set; }
        private readonly string title;
        private readonly string winText;

        public Control(Window window, string name)
        {
            Name = name;
            title = window.Title;
            winText = window.Text;
        }

        public Control(Window window, int id)
        {
            ID = id;
            title = window.Title;
            winText = window.Text;
        }

        public void Click()
        {
            AutoItX.ControlClick(title, winText, $"[NAME:{Name}]");
        }

        public void SetText(string text)
        {
            AutoItX.ControlSetText(title, winText, $"[NAME:{Name}]", text);
        }
    }
}
