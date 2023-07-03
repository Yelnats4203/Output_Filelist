using NPOI.HPSF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel.DataAnnotations;
using System.Text;
using System.Threading.Tasks;

namespace prjFilesList.Models
{
    public static class CDisplayName
    {
        public static string? GetDisplayName(this object instance, string propertyName)
        {
            //get property
            PropertyInfo? propertyInfo = instance.GetType().GetProperty(propertyName);
            if (propertyInfo == null)
                return null;

            //get property's display attribute
            DisplayAttribute? displayAttribute = propertyInfo.GetCustomAttribute<DisplayAttribute>();
            if (displayAttribute == null)
                return null;

            return displayAttribute.Name;
        }

    }
}
