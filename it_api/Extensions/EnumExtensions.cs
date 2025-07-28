using System.ComponentModel;
using System.Reflection;

namespace it_api.Extensions
{
    public static class EnumExtensions
    {
        public static string GetDescription(this Enum value)
        {
            if (value != null)
            {
                var field = value.GetType().GetField(value.ToString());
                var attr = field?.GetCustomAttribute<DescriptionAttribute>();
                return attr?.Description ?? value.ToString();
            }
            return null;
        }
    }
}
