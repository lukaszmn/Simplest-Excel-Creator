using System;
using System.Globalization;

namespace ITLN.SimplestExcelCreator.Utils {

	internal static class StringExtensions {

		public static bool IsNumeric(this string value, CultureInfo culture = null) {
			try {
				Convert.ToDecimal(value, culture ?? CultureInfo.InvariantCulture);
				return true;
			} catch {
				return false;
			}
		}


		public static bool IsNumeric(this string value, out decimal? val, CultureInfo culture = null) {
			try {
				val = Convert.ToDecimal(value, culture ?? CultureInfo.InvariantCulture);
				return true;
			} catch {
				val = null;
				return false;
			}
		}

	}
}