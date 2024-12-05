using Microsoft.AspNetCore.Authorization;
using Microsoft.Win32;
using System.Diagnostics;
using System.Security.Principal;

namespace ExcelSink.Rtns
{
    public class RegistryRtns
    {
        private static bool IsUserAdministrator()
        {
            bool isAdmin;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException ex)
            {
                isAdmin = false;
            }
            catch (Exception ex)
            {
                isAdmin = false;
            }
            return isAdmin;
        }

        public static void Register(string fileType,
                   string shellKeyName, string menuText, string menuCommand)
        {
            if (!IsUserAdministrator())
            {
                MessageBox.Show("For this feature run as administrator");
                return;
            }
            // create path to registry location
            string regPath = string.Format(@"{0}\shell\{1}",
                                           fileType, shellKeyName);

            // add context menu to the registry
            using (RegistryKey key =
                   Registry.ClassesRoot.CreateSubKey(regPath))
            {
                key.SetValue(null, menuText);
            }

            // add command that is invoked to the registry
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(
                string.Format(@"{0}\command", regPath)))
            {
                key.SetValue(null, menuCommand);
            }
        }



        public static bool AddContextMenuItem(string Extension, string MenuName, string MenuDescription, string MenuCommand)
        {
            bool ret = false;
            RegistryKey rkey = Registry.ClassesRoot.OpenSubKey(Extension);

            if (rkey != null)
            {
                string extstring = rkey.GetValue("").ToString();
                rkey.Close();
                if (extstring != null)
                {
                    if (extstring.Length > 0)
                    {
                        rkey = Registry.ClassesRoot.OpenSubKey(
                          extstring, true);
                        if (rkey != null)
                        {
                            string strkey = "shell\\" + MenuName + "\\command";
                            RegistryKey subky = rkey.CreateSubKey(strkey);
                            if (subky != null)
                            {
                                subky.SetValue("", MenuCommand);
                                subky.Close();
                                subky = rkey.OpenSubKey("shell\\" +
                                  MenuName, true);
                                if (subky != null)
                                {
                                    subky.SetValue("", MenuDescription);
                                    subky.Close();
                                }
                                ret = true;
                            }
                            rkey.Close();
                        }
                    }
                }
            }
            return ret;
        }



        public static void Unregister(string fileType, string shellKeyName)
        {
            Debug.Assert(!string.IsNullOrEmpty(fileType) &&
                !string.IsNullOrEmpty(shellKeyName));

            // path to the registry location
            string regPath = string.Format(@"{0}\shell\{1}",
                                           fileType, shellKeyName);

            // remove context menu from the registry
            Registry.ClassesRoot.DeleteSubKeyTree(regPath);
        }
    }
}
