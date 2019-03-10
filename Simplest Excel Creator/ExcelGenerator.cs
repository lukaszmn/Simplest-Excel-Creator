using ITLN.SimplestExcelCreator.Generator;

namespace ITLN.SimplestExcelCreator {

	public static class ExcelGenerator {

		public static void Generate(object[,] data, string path) {
			var zipper = new Zipper();
			zipper.PrepareFolder();

			// TODO: "template\" should be packed to inside the DLL because it will not be copied to bin of referenced projects

			var sheet = new EditedFile(zipper.XlsxRootFolder, @"xl\worksheets\sheet1.xml");
			var strings = new EditedFile(zipper.XlsxRootFolder, @"xl\sharedStrings.xml");

			sheet.Open();
			strings.Open();

			var stringsGenerator = new SharedStringsGenerator(strings);
			new SheetGenerator(data, sheet, stringsGenerator).Generate();
			stringsGenerator.Generate();

			sheet.Save();
			strings.Save();

			zipper.Pack(path);
		}

	}
}
