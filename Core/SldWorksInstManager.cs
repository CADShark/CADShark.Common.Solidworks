using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

namespace CADShark.Common.SolidWorks.Core
{
    public class SldWorksInstManager : ISldWorksInstManager
    {
        public const string StartSwNoJournalDialogAndSuppressAllDialogs = "/r /b";

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable port);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        /// <summary>
        /// Gets the SOLIDWORKS instance from process identifier.
        /// </summary>
        /// <returns>
        /// Return SOLIDWORKS instance.
        /// </returns>
        public SldWorks GetSolidworksInstanceFromProcessId()
        {
            var pid = Process.GetProcessesByName("SLDWORKS").First().Id;
            var numFetched = IntPtr.Zero;
            var monikers = new IMoniker[1];

            GetRunningObjectTable(0, out var runningObjectTable);
            runningObjectTable.EnumRunning(out var monikerEnumerator);

            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                CreateBindCtx(0, out var ctx);

                monikers[0].GetDisplayName(ctx, null, out var runningObjectName);

                if (!runningObjectName.ToLower().Contains("solidworks")) continue;

                runningObjectTable.GetObject(monikers[0], out var runningObjectVal);

                // we should be safe to cast to our "real" solidworks object

                if (runningObjectVal is SldWorks swObj && swObj.GetProcessID() == pid)
                {
                    return swObj;
                }
            }

            return null;
        }

        /// <summary>
        ///  Creates a new instance of SOLIDWORKS
        /// </summary>
        /// <param name="commandlineParameters"></param>
        /// <param name="year"></param>
        /// <param name="timeout">30 seconds for time out.</param>
        /// <returns></returns>
        public SldWorks GetNewInstance(string commandlineParameters = StartSwNoJournalDialogAndSuppressAllDialogs,
            YearE year = YearE.Latest, int timeout = 30)
        {
            var swApp = Extension.CreateSldWorks(commandlineParameters, year, timeout);
            return swApp;
        }

        /// <summary>
        /// Attempts to restart SOLIDWORKS.
        /// </summary>
        /// <param name="swApp"></param>
        /// <param name="commandLineParameters"></param>
        /// <param name="year"></param>
        /// <param name="timeout"></param>
        /// <param name="attempts"></param>
        public void RestartInstance(ref SldWorks swApp, string commandLineParameters = "", YearE year = YearE.Latest,
            int timeout = 30, int attempts = 5)
        {
            if (attempts <= 2)
                attempts = 5;
            if (swApp != null)
            {
                swApp.CloseAllDocuments(true);
                swApp.ExitApp();
                ReleaseInstance(swApp);
                swApp = null;
            }

            for (var i = 1; i <= attempts; i++)
            {
                try
                {
                    swApp = GetNewInstance(commandLineParameters, year, timeout);
                    if (swApp != null)
                        break;
                }
                catch (TimeoutException e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        /// <summary>
        /// Releases sldworks from memory since it is a com object not managed by garbage collector.
        /// </summary>
        /// <param name="swApp"></param>
        public void ReleaseInstance(SldWorks swApp)
        {
            if (swApp != null)
                Marshal.ReleaseComObject(swApp);
        }
    }
}