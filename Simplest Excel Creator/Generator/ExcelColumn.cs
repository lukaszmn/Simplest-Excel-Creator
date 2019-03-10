using System;

namespace ITLN.SimplestExcelCreator.Generator {

	public static class ExcelColumn {

		public static string GetName(int index) {
			--index;

			if (index < 0)
				throw new Exception("Invalid Column #" + (index + 1).ToString());
			else if (index < 26)
				return ((char)('A' + index)).ToString();
			else
				return GetName(index / 26) + GetName(index % 26 + 1);
		}

	}
}
