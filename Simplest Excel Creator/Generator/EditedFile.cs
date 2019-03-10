using System;
using System.Collections.Generic;
using System.IO;

namespace ITLN.SimplestExcelCreator.Generator {

	class EditedFile {

		public string Path { get; private set; }


		private string contents;

		public string Contents {
			get {
				if (contents == null)
					throw new InvalidOperationException("Open file first.");
				else
					return contents;
			}
			private set { contents = value; }
		}


		private Dictionary<string, string> replacements;


		private EditedFile() {
			replacements = new Dictionary<string, string>();
		}

		public EditedFile(string path) : this() {
			Path = path;
		}


		public EditedFile(string root, string subpath) : this() {
			Path = System.IO.Path.Combine(root, subpath);
		}


		public void Open() {
			Contents = File.ReadAllText(Path);
		}


		public void Save() {
			foreach (var pair in replacements)
				Replace(pair.Key, pair.Value);

			File.WriteAllText(Path, Contents);
		}


		public void Replace(string find, string replace) {
			Contents = Contents.Replace(find, replace);
		}


		public void AddReplacement(string find, string replace) {
			replacements.TryGetValue(find, out string val);
			val += replace;
			replacements[find] = val;
		}

	}
}
