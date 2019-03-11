using ITLN.SimplestExcelCreator;

namespace Sample_Application {
	class Program {

		static void Main(string[] args) {
			object[,] data = {
				{ null, null },
				{ "val 1", null },
				{ null, 3 },
				{ null, 9.3 },
				{ "15%", "12.3%" }
			};

			ExcelGenerator.Generate(data, @"C:\abcd.xlsx");
		}

	}
}
