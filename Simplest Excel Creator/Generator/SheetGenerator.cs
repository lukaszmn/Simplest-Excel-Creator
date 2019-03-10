using System;
using System.Collections.Generic;
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
			string columnName = ExcelColumn.GetName(column + 1);
			string cellRef = columnName + (row + 1);

			// TODO parse date: <c r="A2" s="1"><v>42248</v></c>
			// TODO parse percentage: <c r="A5" s="2"><v>0.05</v></c>

			if (val.IsNumeric()) {
				// TODO use invariant conversion
				string valString = val.ToString().Replace(',', '.');
				rows.Append($"<c r=\"{cellRef}\"><v>{valString}</v></c>");
			} else {
				rows.Append($"<c r=\"{cellRef}\" t=\"s\"><v>{stringsGenerator.NextIndex}</v></c>");
				stringsGenerator.Add(val.ToString());
			}
		}

	}
}