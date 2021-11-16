namespace TricentisExtensions.Test.Utils
{
    public static class WriteToFile
    {
        public static void Reset()
        {
            System.IO.File.WriteAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", "");
        }

        public static void Write(string text)
        {
            System.IO.File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", text);
        }
    }
}
