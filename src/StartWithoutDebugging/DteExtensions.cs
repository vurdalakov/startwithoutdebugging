namespace StartWithoutDebugging
{
    using System;
    using EnvDTE;

    internal static class DteExtensions
    {
        public static T GetValue<T>(this Properties properties, String propertyName)
        {
            try
            {
                return (T)properties?.Item(propertyName)?.Value;
            }
            catch
            {
                return default(T);
            }
        }
    }
}
