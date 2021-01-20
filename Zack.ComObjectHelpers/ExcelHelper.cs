namespace Zack.ComObjectHelpers
{
    public static class ExcelHelper
    {
        public static bool IsExcelInstalled()
        {
            return ComHelper.IsProgIDInstalled("Excel.Application");
        }

        public static object CreateExcelApplication()
        {
            return ComHelper.CreateInstanceFromProgID("Excel.Application");
        }

        public static object CreateExcelApplication(this COMReferenceTracker t)
        {
            var app = CreateExcelApplication();
            return t.T(app);
        }
    }
}
