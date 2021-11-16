namespace TricentisExtensions.Modules.WindowForms
{
    class LoginWindow : Window
    {
        private Control ClientBox => new Control(this, "mClientTextBox");
        private Control UserNameBox => new Control(this, "mUserTextBox");
        private Control PassBox => new Control(this, "mPasswordTextBox");
        private Control OkBtn => new Control(this, "mOkButton");

        public LoginWindow() : base("Logon to System S4HANA")
        {
        }

        public LoginWindow(int seconds) : base("Logon to System S4HANA", "", seconds)
        {
        }

        public void LogIn(string client, string uname, string pass)
        {
            ClientBox.SetText(client);
            UserNameBox.SetText(uname);
            PassBox.SetText(@pass);
            Activate();
            OkBtn.Click();
        }
    }
}
