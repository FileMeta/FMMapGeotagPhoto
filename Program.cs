using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FMMapGeotagPhoto
{
    static class Program
    {
        const string c_AppName = "FMMapGeotagPhoto";

        const string c_CommandLabel = "Map Photo";

        const string c_HowToUse =
@"FMMapGeotagPhoto is intended for use as a right-click command from the
Windows File Explorer. First, you must register the application. Do this by
copying the file 'FMMapGeotagPhoto' to an apprpriate directory and then
registering it by executing the following command:
   FMMapGeotagPhoto -register

Once the application has been registered, you can right-click on any .jpg or
.jpeg file and select 'Map Photo'. If the photo has been geotagged (usually
the case for phone cameras) then you will get a map of the location where
the photo was taken. If the photo has not been geotagged then you get a
message, 'Photo is not geotagged. Cannot map its location.'

To unregister the application, when moving it or removing from the computer,
execute the following command:
   FMMapGeotagPhoto -unregister";

        const string c_Registered =
@"Application has registered the 'Map Photo' command for .jpg and .jpeg files.";

        const string c_Unregistered =
@"Application has unregistered the 'Map Photo' command for .jpg and .jpeg files.";

        const string c_InternalRegCommand = "::~register~::";
        const string c_InternalUnregCommand = "::~unregister~::";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                String[] args = Environment.GetCommandLineArgs();
                if (args.Length != 2)
                {
                    MessageBox.Show(ForegroundWindow.Instance, c_HowToUse, c_AppName);
                    return;
                }

                switch (args[1].ToLowerInvariant())
                {
                    case "-register":
                    case "/register":
                        ElevateCommand(c_InternalRegCommand);
                        break;

                    case "-unregister":
                    case "/unregister":
                        ElevateCommand(c_InternalUnregCommand);
                        break;

                    case c_InternalRegCommand:
                        RegisterApplication();
                        break;

                    case c_InternalUnregCommand:
                        UnregisterApplication();
                        break;

                    default:
                        MapGeotagPhoto(args[1]);
                        break;
                }

            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), c_AppName);
            }
        }

        static void MapGeotagPhoto(string path)
        {
            double latitude;
            double longitude;
            if (!Props.GetLatitudeLongitude(path, out latitude, out longitude))
            {
                MessageBox.Show(
                    ForegroundWindow.Instance,
                    "Photo is not geotagged. Cannot map its location.",
                    c_AppName
                );
                return;
            }

            string url = string.Format("http://www.bing.com/maps?&where1={0:f8}%20{1:f8}", latitude, longitude);
            System.Diagnostics.Process.Start(url);
        }

        static void ElevateCommand(string command)
        {
            ProcessStartInfo proc = new ProcessStartInfo();
            proc.UseShellExecute = true;
            proc.WorkingDirectory = Environment.CurrentDirectory;
            proc.FileName = Application.ExecutablePath;
            proc.Arguments = command;
            proc.Verb = "runas";

            try
            {
                Process.Start(proc);
            }
            catch
            {
                // The user refused to allow privileges elevation.
                // Do nothing
            }
        }

        static void RegisterApplication()
        {
            string command = string.Format("\"{0}\" \"%1\"", Application.ExecutablePath);
            RegisterFileAssoc(".jpg", c_CommandLabel, command);
            RegisterFileAssoc(".jpeg", c_CommandLabel, command);
            MessageBox.Show(ForegroundWindow.Instance, c_Registered, c_AppName);
        }

        static void RegisterFileAssoc(string extension, string command, string commandLine)
        {
            var keyFileAssoc = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("SystemFileAssociations", true);
            var key = keyFileAssoc.CreateSubKey(extension);
            key = key.CreateSubKey("Shell");
            key = key.CreateSubKey(command);
            key = key.CreateSubKey("Command");
            key.SetValue(string.Empty, commandLine);
        }

        static void UnregisterApplication()
        {
            UnregisterFileAssoc(".jpg", c_CommandLabel);
            UnregisterFileAssoc(".jpeg", c_CommandLabel);
            MessageBox.Show(ForegroundWindow.Instance, c_Unregistered, c_AppName);
        }

        static void UnregisterFileAssoc(string extension, string command)
        {
            var keyFileAssoc = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("SystemFileAssociations", true);
            var key = keyFileAssoc.OpenSubKey(extension, true);
            if (key == null) return;
            key = key.OpenSubKey("Shell", true);
            if (key == null) return;
            key.DeleteSubKeyTree(command, false);
        }
    }

    class Props
    {

        static WinShell.PROPERTYKEY s_PkLatitude;
        static WinShell.PROPERTYKEY s_PkLongitude;
        static WinShell.PROPERTYKEY s_PkLatitudeRef;
        static WinShell.PROPERTYKEY s_PkLongitudeRef;

        static Props()
        {
            using (WinShell.PropertySystem propsys = new WinShell.PropertySystem())
            {
                s_PkLatitude = propsys.GetPropertyKeyByName("System.GPS.Latitude");
                s_PkLongitude = propsys.GetPropertyKeyByName("System.GPS.Longitude");
                s_PkLatitudeRef = propsys.GetPropertyKeyByName("System.GPS.LatitudeRef");
                s_PkLongitudeRef = propsys.GetPropertyKeyByName("System.GPS.LongitudeRef");
            }
        }

        public static bool GetLatitudeLongitude(string filename, out double latitude, out double longitude)
        {
            latitude = longitude = 0.0;

            using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(filename))
            {
                double[] angle = (double[])store.GetValue(s_PkLatitude);
                string direction = (string)store.GetValue(s_PkLatitudeRef);
                if (angle == null || direction == null) return false;
                //Debug.WriteLine("Latitude: {0} {1},{2},{3}", direction, angle[0], angle[1], angle[2]);
                latitude = DegMinSecToDouble(direction, angle);

                angle = (double[])store.GetValue(s_PkLongitude);
                direction = (string)store.GetValue(s_PkLongitudeRef);
                if (angle == null || direction == null) return false;
                //Debug.WriteLine("Longitude: {0} {1},{2},{3}", direction, angle[0], angle[1], angle[2]);
                longitude = DegMinSecToDouble(direction, angle);
            }

            return true;
        }

        static double DegMinSecToDouble(string direction, double[] dms)
        {
            double result = dms[0];
            if (dms.Length > 1) result += dms[1] / 60.0;
            if (dms.Length > 2) result += dms[2] / 3600.0;

            if (string.Equals(direction, "W", StringComparison.OrdinalIgnoreCase) || string.Equals(direction, "S", StringComparison.OrdinalIgnoreCase))
            {
                result = -result;
            }

            return result;
        }
        
        public static void DumpProperties(string filename)
        {
            using (WinShell.PropertySystem propsys = new WinShell.PropertySystem())
            {
                using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(filename))
                {
                    int count = store.Count;
                    for (int i = 0; i < count; ++i)
                    {
                        WinShell.PROPERTYKEY key = store.GetAt(i);

                        string name;
                        try
                        {
                            using (WinShell.PropertyDescription desc = propsys.GetPropertyDescription(key))
                            {
                                name = string.Concat(desc.CanonicalName, " ", desc.DisplayName);
                            }
                        }
                        catch
                        {
                            name = string.Format("({0}:{1})", key.fmtid, key.pid);
                        }

                        object value = store.GetValue(key);
                        string strValue;

                        if (value is string[])
                        {
                            strValue = string.Join(";", (string[])value);
                        }
                        else if (value is double[])
                        {
                            strValue = string.Join(",", (double[])value);
                        }
                        else
                        {
                            strValue = value.ToString();
                        }

                        Debug.WriteLine("{0}: {1}", name, strValue);
                    }
                }
            }
        }
    }

    public class ForegroundWindow : IWin32Window
    {
        private static ForegroundWindow _window = new ForegroundWindow();

        private ForegroundWindow()
        {
        }
        
        public static IWin32Window Instance
        {
            get
            {
                return _window;
            }
        }
        
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        IntPtr IWin32Window.Handle
        {
            get
            {
                return GetForegroundWindow();
            }
        }
    }
}
