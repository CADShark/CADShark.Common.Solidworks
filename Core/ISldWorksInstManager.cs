using SolidWorks.Interop.sldworks;

namespace CADShark.Common.SolidWorks.Core;

public interface ISldWorksInstManager
{
    SldWorks GetSolidworksInstanceFromProcessId();

    /// <summary>
    /// Gets a new instance.
    /// </summary>
    /// <param name="commandLineParameters"></param>
    /// <param name="year">Year version of SOLIDWORKS.</param>
    /// <param name="timeout"></param>
    /// <returns>Pointer to the SOLIDWORKS application.</returns>
    SldWorks GetNewInstance(
        string commandLineParameters = SldWorksInstManager.StartSwNoJournalDialogAndSuppressAllDialogs,
        YearE year = YearE.Latest, int timeout = 30);


    void ReleaseInstance(SldWorks swApp);

    /// <summary>
    ///  Restarts the specified SOLIDWORKS application.
    /// </summary>
    /// <param name="swApp">Reference to the SOLIDWORKS application.</param>
    /// <param name="commandLineParameters">Command line params.</param>
    /// <param name="year">Year. Default is latest.</param>
    /// <param name="timeout">Timeout in seconds.</param>
    /// <param name="attempts">Number of attempts</param>
    void RestartInstance(ref SldWorks swApp,
        string commandLineParameters = SldWorksInstManager.StartSwNoJournalDialogAndSuppressAllDialogs,
        YearE year = YearE.Latest, int timeout = 30, int attempts = 5);
}