using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Zack.ComObjectHelpers
{
    public static class ComHelper
    {

        public static bool IsProgIDInstalled(string progID)
        {
            return Registry.ClassesRoot.GetSubKeyNames().Contains(progID);
        }

        public static object CreateInstanceFromProgID(string progID, bool throwOnError=true)
        {
            Type type = Type.GetTypeFromProgID(progID, throwOnError);
            return Activator.CreateInstance(type);
        }

        public static string GetComObjectDescription(object comObject)
        {
            if(comObject==null)
            {
                return "null";
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("TypeName:").AppendLine(TypeDescriptor.GetClassName(comObject));
            sb.AppendLine("Properties:");
            //https://github.com/dotnet/runtime/issues/47248
            //TypeDescriptor.GetProperties depends on System.Windows.Forms
            foreach (PropertyDescriptor p in TypeDescriptor.GetProperties(comObject))
            {
                sb.AppendLine($"Name:{p.Name},DisplayName:{p.DisplayName},PropertyType:{p.PropertyType}");
            }
            return sb.ToString();
        }

        public static string TypeName(object comObject)
        {
            return TypeDescriptor.GetClassName(comObject);
        }
    }
}
