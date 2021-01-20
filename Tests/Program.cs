using System;
using System.IO;
using System.Runtime.InteropServices;
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
                
                var slides = t.T(ppFile.Slides);
                /*
                int i = 0;
                foreach (var slide in slides)
                {
                    t.T(slide);
                    slide.Export(@$"{workDir}\{i}.png", "PNG");
                    i++;
                }*/
                //用foreach在Debug模式powerpoint不会退出
                //https://github.com/dotnet/runtime/issues/47249
                for (int i= 0;i < slides.Count;i++)
                {
                    var slide = t.T(slides[i+1]);
                    slide.Export(@$"{workDir}\{i}.png", "PNG");
                }
                ppFile.Close();
                ppApp.Quit();
            }
        }
    }
}
