using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using static System.String;

namespace CADShark.Common.SolidWorks.Core
{
    internal class Extension
    {
        /// <summary>
        /// Returns the SOLIDWORKS installation directory for the specified year (if it exists).
        /// </summary>
        /// <param name="year">Year</param>
        /// <returns><see cref="DirectoryInfo"/> object, null if failed.</returns>
        public static DirectoryInfo GetSolidworksInstallationDirectory(int year)
        {
            using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                using (var key = hklm.OpenSubKey(@"SOFTWARE\SolidWorks\SOLIDWORKS " + year + @"\Setup"))
                {
                    if (key == null)
                        return null;
                    else
                    {
                        return new DirectoryInfo(key.GetValue("SolidWorks Folder") as string);
                    }

                }
            }
        }

        /// <summary>
        /// Creates a new instance of the SOLIDWORKS application using <see cref="Process.Start()"/>.
        /// </summary>
        /// <param name="timeoutSec"></param>
        /// <param name="suppressDialog">True to suppress SOLIDWORKS dialogs.</param>
        /// <returns>Pointer to the new instance of SOLIDWORKS.</returns>
        /// <exception cref="TimeoutException">Thrown if method times out.</exception>
        public static SldWorks CreateSldWorks(string commandlineParameters = "", YearE _year = YearE.Latest, int timeoutSec = 30)
        {
            var years = ReleaseYears();
            if (years.Length == 0)
                throw new Exception($"SOLIDWORKS is not installed on this computer [{System.Environment.MachineName}].");
            Array.Sort(years);

            var year = years.Last();

            if (_year != (int)YearE.Latest)
            {
                if (years.Contains((int)_year) == false)
                    throw new Exception($"Could not find installation directory for SOLIDWORKS. [{_year}].");

                year = (int)_year;
            }

            var installationDirectory = GetSolidworksInstallationDirectory(year);
            if (installationDirectory == null)
                throw new Exception($"Could not find installation directory for SOLIDWORKS. Year = [{year}].");

            var appPath = installationDirectory.FullName;
            var timeout = TimeSpan.FromSeconds(timeoutSec);
            var startTime = DateTime.Now;
            var args = IsNullOrWhiteSpace(commandlineParameters) ? "/r" : commandlineParameters;
            var prc = Process.Start(appPath + "sldworks.exe", args);
            SldWorks app = null;
            while (app == null)
            {
                if (DateTime.Now - startTime > timeout)
                {
                    if (prc.Id != 0)
                        if (prc.HasExited == false)
                            prc.Kill();

                    throw new TimeoutException($"Could not create a new SOLIDWORKS process within {timeoutSec} seconds.");
                }
                app = GetSwAppFromProcess(prc.Id);
            }
            return app;
        }

        /// <summary>
        /// Returns an array of all installed SOLIDWORKS release years.
        /// </summary>
        /// <returns>Array of integers.</returns>
        public static int[] ReleaseYears()
        {
            var solidworksKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\SolidWorks");
            var names = solidworksKey?.GetSubKeyNames();
            var years = new List<int>();
            
            var regex = new Regex(@"^solidworks ([\d]{4})$", RegexOptions.IgnoreCase);
            foreach (var name in names)
            {
                if (regex.IsMatch(name))
                {
                    int year = int.MinValue;
                    var match = regex.Match(name);
                    var capture = match.Groups[1].Value;
                    var ret = int.TryParse(capture, out year);
                    if (ret)
                        years.Add(year);
                }
            }
            return years.ToArray();
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static SldWorks GetSwAppFromProcess(int processId)
        {
            var monikerName = "SolidWorks_PID_" + processId.ToString();

            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);
                rot.EnumRunning(out monikers);

                var moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException)
                        {
                        }
                    }

                    if (string.Equals(monikerName,
                        name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        object app;
                        rot.GetObject(curMoniker, out app);
                        return (SldWorks)app;
                    }
                }
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (context != null)
                {
                    Marshal.ReleaseComObject(context);
                }
            }

            return null;
        }

        /// <summary>
        /// Converts a release year to the SOLIDWORKS revision number.
        /// </summary>
        /// <param name="releaseYear">Release year.</param>
        /// <returns>SOLIDWORKS revision number</returns>
        /// <remarks>This method produces correct results for revision number from the year 2003 and newer.</remarks>
        public static int ConvertYearToSwRevisionNumber(int releaseYear)
        {
            return releaseYear - 1992;
        }

        /// <summary>
        /// Converts the SOLIDWORKS revision number to a release year.
        /// </summary>
        /// <param name="revNumber">Revision number.</param>
        /// <returns>Release year</returns>
        /// <remarks><ul>
        /// <li>This method produces correct results for revision number from the year 2003 and newer.</li>
        /// <li>Method return -1 if it fails.</li>
        /// </ul></remarks>
        public static int ConvertSwRevisionNumberToYear(string revNumber)
        {
            try
            {
                var spN = int.Parse(revNumber.Split('.').First());
                return 1992 + spN;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        /// <summary>
        /// Converts the SOLIDWORKS revision number to a release year.
        /// </summary>
        /// <param name="revNumber">Revision number.</param>
        /// <returns>Release year</returns>
        /// <remarks><ul>
        /// <li>This method produces correct results for revision number from the year 2003 and newer.</li>
        /// <li>Method return -1 if it fails.</li>
        /// </ul></remarks>
        public static int ConvertSWRevisionNumberToYear(int revNumber)
        {
            return 1992 + revNumber;
        }
    }
}
