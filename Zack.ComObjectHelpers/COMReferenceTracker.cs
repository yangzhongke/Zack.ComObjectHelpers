using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Zack.ComObjectHelpers
{
    public class COMReferenceTracker : IDisposable
    {
        private List<dynamic> objects = new List<dynamic>();

        public dynamic T(dynamic obj)
        {
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
                Marshal.FinalReleaseComObject(obj);
            }
            GC.Collect();
        }
    }
}
