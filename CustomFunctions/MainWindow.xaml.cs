//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;

using MindFusion.Spreadsheet.Wpf.Expressions;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.CustomFunctions
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			var activeSheet = workbook.Worksheets.Add();
			workbookView.ActiveWorksheet = activeSheet;

			workbook.ExtendedEvaluator = new CustomFunctions();

			IInformationProvider infoProvider = null;
			using (Stream s = GetType().Assembly.GetManifestResourceStream("MindFusion.Spreadsheet.Wpf.Samples.CS.CustomFunctions.CustomFunctionInformation.xml"))
			{
				using (var reader = new StreamReader(s))
					infoProvider = XmlInformationProvider.Create(XDocument.Load(reader));
			}
			workbook.ExtendedInformationProvider = infoProvider;

			activeSheet.BeginInit();
			activeSheet.Cells[1, 0].Data = "Result";
			activeSheet.Cells[2, 0].Data = "Formula";
			activeSheet.Cells[3, 0].Data = "Description";
			activeSheet.Cells[1, 1].Data = "=EASTER(YEAR(TODAY()))";
			activeSheet.Cells[2, 1].Data = "'=EASTER(YEAR(TODAY()))";
			activeSheet.Cells[3, 1].Data = infoProvider.GetFunctionSummary("EASTER");
			activeSheet.Cells[1, 2].Data = "=CLENGTH(10)";
			activeSheet.Cells[2, 2].Data = "'=CLENGTH(10)";
			activeSheet.Cells[3, 2].Data = infoProvider.GetFunctionSummary("CLENGTH");
			activeSheet.Cells[1, 3].Data = "=CAREA(10)";
			activeSheet.Cells[2, 3].Data = "'=CAREA(10)";
			activeSheet.Cells[3, 3].Data = infoProvider.GetFunctionSummary("CAREA");
			activeSheet.CellRanges[1, 0, 3, 0].Style.TextBrush = Brushes.Gray;
			activeSheet.EndInit();
			workbookView.ResizeColumnsToFit(1, 3);
		}

		public class CustomFunctions : IExpressionEvaluator
		{
			public object EvaluateIdentifier(IExpressionEvaluatorContext context, string identifier)
			{
				return null;
			}

			public object InvokeFunction(IExpressionEvaluatorContext context, string name, IList parameters)
			{
				if (name == "EASTER")
				{
					if (parameters.Count != 1)
						return context.GetError(ErrorType.Value, "Wrong number of arguments.");

					double? year = context.ToDouble(parameters[0]);
					if (year == null)
						return context.GetError(ErrorType.Value, "Wrong argument type.");

					return EasterSundayOf((int)year.Value);
				}
				else if (name == "CLENGTH")
				{
					if (parameters.Count != 1)
						return context.GetError(ErrorType.Value, "Wrong number of arguments.");

					double? radius = context.ToDouble(parameters[0]);
					if (radius == null)
						return context.GetError(ErrorType.Value, "Wrong argument type.");

					return 2 * Math.PI * radius.Value;
				}
				else if (name == "CAREA")
				{
					if (parameters.Count != 1)
						return context.GetError(ErrorType.Value, "Wrong number of arguments.");

					double? radius = context.ToDouble(parameters[0]);
					if (radius == null)
						return context.GetError(ErrorType.Value, "Wrong argument type.");

					// Note: obtain PI by calling the respective function of the core evaluator
					double? pi = context.ToDouble(context.GetCoreEvaluator().InvokeFunction(context, "PI", new object[] { }));
					if (pi == null)
						return context.GetError(ErrorType.Value, "Cannot evaluate PI.");

					return pi.Value * Math.Pow(radius.Value, 2);
				}


				return null;
			}

			private static DateTime EasterSundayOf(int year)
			{
				int Y = year;
				int a = Y % 19;
				int b = Y / 100;
				int c = Y % 100;
				int d = b / 4;
				int e = b % 4;
				int f = (b + 8) / 25;
				int g = (b - f + 1) / 3;
				int h = (19 * a + b - d - g + 15) % 30;
				int i = c / 4;
				int k = c % 4;
				int L = (32 + 2 * e + 2 * i - h - k) % 7;
				int m = (a + 11 * h + 22 * L) / 451;
				int Month = (h + L - 7 * m + 114) / 31;
				int Day = ((h + L - 7 * m + 114) % 31) + 1;
				return new DateTime(year, Month, Day);
			}
		}
	}
}
