using System;
using System.Windows.Forms;
using CADBooster.SolidDna;
using CADShark.Common.Logging;

namespace CADShark.Common.SolidWorks.SplitConfig
{
    public class SplitConfigExtraLeyer
    {
        public void DoSplit()
        {
            var bb = new FolderBrowserDialog
            {
                Site = null,
                Tag = null,
                ShowNewFolderButton = false,
                SelectedPath = null,
                RootFolder = Environment.SpecialFolder.Desktop,
                Description = @"Виберіть папку з моделями для розділення"
            };

            if (DialogResult.OK != bb.ShowDialog()) return;

            var path = bb.SelectedPath;

            if (!AddInIntegration.ConnectToActiveSolidWorksForStandAlone())
            {
                CadLogger.Warning(@"Не знайдено запущеного процесу SOLIDWORKS!");
            }

            var splitFile = new Common.SolidWorks.SplitConfig.SplitConfig();
            splitFile.SplitModelConfigs(path);
        }
    }
}