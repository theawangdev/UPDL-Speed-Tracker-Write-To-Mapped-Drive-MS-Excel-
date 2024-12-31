using System;
using System.Reflection;
using System.Configuration;

namespace UPDL_Speed_Tracker
{
    //GetAppSetting
    class GetAppSetting
    {
        public static string Get(string key)
        {
            try
            {
                var asmPath = Assembly.GetExecutingAssembly().Location;
                var config = ConfigurationManager.OpenExeConfiguration(asmPath);
                var setting = config.AppSettings.Settings[key];

                if (setting == null)
                {
                    return null;
                }
                else
                {
                    return setting.Value;
                }
            }
            catch (Exception e)
            {
                throw new InvalidOperationException("Error reading configuration setting", e);
            }
        }
    }
}
