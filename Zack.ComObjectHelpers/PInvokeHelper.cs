using System;
using System.Runtime.InteropServices;

namespace Zack.ComObjectHelpers
{
    public static class PInvokeHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(int hWnd, out int processId);


    }
}
