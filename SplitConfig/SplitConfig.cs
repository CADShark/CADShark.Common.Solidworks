using System.Collections.Generic;
using System.IO;
using CADBooster.SolidDna;
using CADShark.Common.Logging;
using CADShark.Common.SolidWorks.Errors;

namespace CADShark.Common.SolidWorks.SplitConfig
{
	public class SplitConfig
	{
		public SplitConfig()
		{
			if (AddInIntegration.ConnectToActiveSolidWorksForStandAlone()) return;

			const string message = "Failed to connect to SolidWorks";
			CadLogger.Error(message);
			throw new ConnectionException(message);
		}
		public void SplitModelConfigs(string fileName)
		{
			var model = SolidWorksEnvironment.Application.OpenFile(fileName, OpenDocumentOptions.ReadOnly, configuration: null);

			if (model == null || !model.IsPart) return;

			CadLogger.Info($@"Open File {fileName}");

			var pathModel = Path.GetDirectoryName(model.FilePath);
			var configurations = model.ConfigurationNames;
			var paths = new List<string>();

			foreach (var config in configurations)
			{
				var swConf = model.ActivateConfiguration(config);

				CadLogger.Info($@"Activate Configuration {swConf}");

				if (pathModel == null) continue;

				var path = Path.Combine(pathModel, config + ".SLDPRT");

				var status = model.SaveAs(path, SaveAsVersion.CurrentVersion, SaveAsOptions.Silent);
				if (!status.Successful)
				{
					CadLogger.Error(status.ToString());
					continue;
				}
				paths.Add(path);
			}


			SolidWorksEnvironment.Application.CloseFile(pathModel);

			GetPart(paths);
		}

		private static void GetPart(List<string> paths)
		{
			foreach (var path in paths)
			{
				var model = SolidWorksEnvironment.Application.OpenFile(path, OpenDocumentOptions.Silent, configuration: null);
				var configurations = model.ConfigurationNames;
				foreach (var configurationName in configurations)
				{
					var status = model.DeleteConfiguration(configurationName);
					if (!status)
					{
						CadLogger.Error(false.ToString()); continue;
					}

					CadLogger.Info($@"Delete Configuration {configurationName}");
				}
				var statSaveResult = model.SaveAs(path, SaveAsVersion.CurrentVersion, SaveAsOptions.Silent);
				CadLogger.Error(statSaveResult.ToString());

				SolidWorksEnvironment.Application.CloseFile(path);
			}
		}
	}
}