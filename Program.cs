using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;

namespace FMMapGeotagPhoto
{
    static class Program
    {
        static string s_AppName = "FMMapGeotagPhoto";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                String[] args = Environment.GetCommandLineArgs();
                if (args.Length < 2)
                {
                    MessageBox.Show(ForegroundWindow.Instance, "No file name specified.", s_AppName);
                    return;
                }

                /*
                DialogResult dr = MessageBox.Show(ForegroundWindow.Instance,
                    string.Format("Filename: {0}", args[1]),
                    s_AppName,
                    MessageBoxButtons.OKCancel
                );
                */

                double latitude;
                double longitude;
                if (!Props.GetLatitudeLongitude(args[1], out latitude, out longitude))
                {
                    MessageBox.Show(
                        ForegroundWindow.Instance,
                        "Photo is not geotagged. Cannot map its location.",
                        s_AppName
                    );
                    return;
                }

                /*
                DialogResult dr = MessageBox.Show(
                    ForegroundWindow.Instance,
                    string.Format("Lat, Long: {0}, {1}", latitude, longitude),
                    s_AppName,
                    MessageBoxButtons.OKCancel
                );
                */

                string url = string.Format("http://www.bing.com/maps?&where1={0:f8}%20{1:f8}", latitude, longitude);
                System.Diagnostics.Process.Start(url);

            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), s_AppName);
            }
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
