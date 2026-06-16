
using System;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	public class LanguageConverter : IValueConverter
	{
		public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
		{
			string header = value.ToString();

			if (header.StartsWith("en-") ||
				header.StartsWith("es-") ||
				header.StartsWith("pt-") ||
				header.StartsWith("de-") ||
				header.StartsWith("fr-") ||
				header.StartsWith("ko-") ||
				header.StartsWith("ru-") ||
				header.StartsWith("zh-"))
				return HightlightBrush;

			return null;
		}

		public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
		{
			throw new NotImplementedException();
		}

		static SolidColorBrush HightlightBrush = new SolidColorBrush(Color.FromArgb(50, 255, 100, 0));
	}
}
