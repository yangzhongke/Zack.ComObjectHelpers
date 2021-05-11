using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Zack.ComObjectHelpers
{
    public class COMReferenceTracker : IDisposable
    {
        private List<dynamic> objects = new List<dynamic>();

        public dynamic T(dynamic obj)
        {
            if(Marshal.IsComObject(obj)==false)
            {
                throw new ArgumentException("obj is not a ComObject.");
            }
            lock (objects)
            {
                objects.Add(obj);
                return obj;
            }
        }

        public void Dispose()
        {
            foreach (var obj in objects)
            {
                try
                {
                    Marshal.FinalReleaseComObject(obj);
                }
                catch(InvalidComObjectException ex)
                {
                    Debug.WriteLine(ex);
                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
