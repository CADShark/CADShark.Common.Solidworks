using SolidWorks.Interop.sldworks;

namespace CADShark.Common.SolidWorks.Core
{
    public interface ISolidWorksInstanceManager
    {
        SldWorks GetSOLIDWORKSInstanceFromProcessID();

        /// <summary>
        /// Gets a new instance.
        /// </summary>
        /// <param name="commandlineArgs">commandline args on how to start SOLIDWORSKS. For a full list of arguments, please see this <see href="https://www.cadoverflow.com/t/solidworks-command-line-arguments/279">thread</see>.</param>
        /// <param name="timeoutInSeconds">The timeout in seconds.</param>
        /// <param name="_year">Year version of SOLIDWORKS.</param>
        /// <returns>Pointer to the SOLIDWORKS application.</returns>
        SldWorks GetNewInstance(
            string commandLineParameters = SolidWorksInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs,
            YearE _year = YearE.Latest, int timeout = 30);


        void ReleaseInstance(SldWorks swApp);

        /// <summary>
        ///  Restarts the specified SOLIDWORKS application.
        /// </summary>
        /// <param name="swApp">Reference to the SOLIDWORKS application.</param>
        /// <param name="commandLineParameters">Command line params.</param>
        /// <param name="_year">Year. Default is latest.</param>
        /// <param name="timeout">Timeout in seconds.</param>
        /// <param name="attempts">Number of attempts</param>
        void RestartInstance(ref SldWorks swApp,
            string commandLineParameters = SolidWorksInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs,
            YearE _year = YearE.Latest, int timeout = 30, int attempts = 5);
    }
}