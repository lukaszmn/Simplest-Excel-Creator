using System.Collections.Generic;
using System.Text;

namespace ITLN.SimplestExcelCreator.Generator {

	internal class SharedStringsGenerator {

		private readonly EditedFile sharedStrings;

		private readonly List<string> values = new List<string>();

		public int NextIndex => values.Count;

		public void Add(string val) => values.Add(val);


		public SharedStringsGenerator(EditedFile sharedStrings) {
			this.sharedStrings = sharedStrings;
		}


		public void Generate() {
			sharedStrings.Replace("%COUNT%", values.Count.ToString());

			StringBuilder sb = new StringBuilder();
			foreach (string val in values) {
				// TODO encode XML
				sb.Append($"<si><t>{val}</t></si>");
			}

			sharedStrings.Replace("%STRINGS%", sb.ToString());
		}
	
	}
}
