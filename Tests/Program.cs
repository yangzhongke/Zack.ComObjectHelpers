using System;
using System.IO;
using Zack.ComObjectHelpers;

namespace Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            if(PowerPointHelper.IsPowerPointInstalled()==false)
            {
                Console.WriteLine("PowerPoint is not installed!");
                return;
            }
            string workDir = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(workDir, "6.pptx");
            using (COMReferenceTracker t = new COMReferenceTracker())
            {
                dynamic ppApp = t.CreatePowerPointApplication();                
                var ppFile = t.T(t.T(ppApp.Presentations).Open(filePath, false, false, false));
                int counter = 0;
                foreach(var slide in t.T(ppFile.Slides))
                {
                    slide.Export(@$"{workDir}\{counter}.png", "PNG");
                    counter++;
                }
                ppFile.Close();
                ppApp.Quit();
            }                
        }
    }
}
