namespace Zack.ComObjectHelpers
{
    public static class PowerPointHelper
    {
        public static bool IsPowerPointInstalled()
        {
            return ComHelper.IsProgIDInstalled("PowerPoint.Application");
        }

        public static object CreatePowerPointApplication()
        {
            return ComHelper.CreateInstanceFromProgID("PowerPoint.Application");
        }

        public static object CreatePowerPointApplication(this COMReferenceTracker t)
        {
            var app = CreatePowerPointApplication();
            return t.T(app);
        }
    }
}
