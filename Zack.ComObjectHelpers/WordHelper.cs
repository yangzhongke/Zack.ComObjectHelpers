namespace Zack.ComObjectHelpers
{
    public static class WordHelper
    {
        public static bool IsWordInstalled()
        {
            return ComHelper.IsProgIDInstalled("Word.Application");
        }

        public static object CreateWordApplication()
        {
            return ComHelper.CreateInstanceFromProgID("Word.Application");
        }

        public static object CreateWordApplication(this COMReferenceTracker t)
        {
            var app = CreateWordApplication();
            return t.T(app);
        }
    }
}
