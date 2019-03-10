using System;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using ITLN.SimplestExcelCreator.Utils;

namespace ITLN.SimplestExcelCreator.Generator {

	internal class SheetGenerator {

		private readonly object[,] data;

		private readonly EditedFile sheet;

		private readonly SharedStringsGenerator stringsGenerator;

		private readonly int rowsCount;
		private readonly int columnsCount;

		private StringBuilder rows;


		public SheetGenerator(object[,] data, EditedFile sheet, SharedStringsGenerator stringsGenerator) {
			this.data = data;
			this.sheet = sheet;
			this.stringsGenerator = stringsGenerator;

			rowsCount = data.GetLength(0);
			columnsCount = data.GetLength(1);
		}


		public void Generate() {
			setSize();
			setRows();
		}


		private void setSize() {
			string column = ExcelColumn.GetName(columnsCount);

			string size = $"A1:{column}{rowsCount}";
			sheet.Replace("%SIZE%", size);
		}


		private void setRows() {
			rows = new StringBuilder();

			for (int row = 0; row < rowsCount; ++row) {
				if (!rowEmpty(row))
					setRow(row);
			}

			sheet.Replace("%ROWS%", rows.ToString());
		}


		private bool rowEmpty(int row) {
			for (int column = 0; column < columnsCount; ++column) {
				if (data[row, column] != null)
					return false;
			}

			return true;
		}


		private void setRow(int row) {
			rows.Append($"<row r=\"{row + 1}\" spans=\"1:{columnsCount}\" x14ac:dyDescent=\"0.25\">");
			setColumns(row);
			rows.Append("</row>");
		}


		private void setColumns(int row) {
			for (int column = 0; column < columnsCount; ++column) {
				object val = data[row, column];
				if (val != null)
					setColumn(row, column, val);
			}
		}


		private void setColumn(int row, int column, object val) {
			Debug.Assert(val != null);

			string columnName = ExcelColumn.GetName(column + 1);
			string cellRef = columnName + (row + 1);

			// TODO parse date: <c r="A2" s="2"><v>42248</v></c>

			string valS = val.ToString();

			if (valS.EndsWith("%") && valS.Substring(0, valS.Length - 1).IsNumeric(out decimal? perc)) {
				// percentage
				// TODO: please note these values will be displayed with 0 decimal points (though they will be stored in cell)
				string valString = (perc.Value / 100).ToString(CultureInfo.InvariantCulture);
				rows.Append($"<c r=\"{cellRef}\" s=\"1\"><v>{valString}</v></c>");

			} else if (val.IsNumeric()) {
				// these steps are necessary to convert decimal packed in an object to an invariant string
				decimal dec = Convert.ToDecimal(val, CultureInfo.InvariantCulture);
				string valString = dec.ToString(CultureInfo.InvariantCulture);
				rows.Append($"<c r=\"{cellRef}\"><v>{valString}</v></c>");

			} else {
				rows.Append($"<c r=\"{cellRef}\" t=\"s\"><v>{stringsGenerator.NextIndex}</v></c>");
				stringsGenerator.Add(valS);
			}
		}

	}
}