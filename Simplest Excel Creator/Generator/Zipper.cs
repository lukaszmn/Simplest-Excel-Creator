using System.IO;
using System.IO.Compression;
using System.Reflection;
using ITLN.Utils.Disk;

namespace ITLN.SimplestExcelCreator.Generator {

	internal class Zipper {

		public string XlsxRootFolder { get; private set; }

		public void PrepareFolder() {
			XlsxRootFolder = Path.Combine(
				Path.GetTempPath(),
				Assembly.GetExecutingAssembly().GetName().Name,
				"xlsx"
			);

			try {
				Directory.Delete(XlsxRootFolder, true);
			} catch (IOException) { }

			string templatePath = getTemplatePath();
			// TODO remove ref to ITLN Utils
			FileCopier.CopyFolder(templatePath, XlsxRootFolder, true, true, OverwriteAction.Throw, true);
		}


		public void Pack(string output) {
			if (File.Exists(output))
				File.Delete(output);

			ZipFile.CreateFromDirectory(XlsxRootFolder, output);
		}


		private string getTemplatePath() {
			return Path.Combine(
				Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
				"template"
			);
		}

	}
}
