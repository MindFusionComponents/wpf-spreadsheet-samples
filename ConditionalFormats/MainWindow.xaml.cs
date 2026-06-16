//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;

using MindFusion.Spreadsheet.Wpf.StandardForms;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.ConditionalFormats
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			workbookView.FontFamily = new FontFamily("Calibri");
			workbookView.FontSize = 15;

			var worksheet = workbook.Worksheets.Add();
			workbookView.ActiveWorksheet = worksheet;

			// Create sample data
			string[] products = new string[]
			{
				"Chai",
				"Chang",
				"Aniseed Syrup",
				"Chef Anton's Cajun Seasoning",
				"Chef Anton's Gumbo Mix",
				"Grandma's Boysenberry Spread",
				"Uncle Bob's Organic Dried Pears",
				"Northwoods Cranberry Sauce",
				"Mishi Kobe Niku",
				"Ikura",
				"Queso Cabrales",
				"Queso Manchego La Pastora",
				"Konbu",
				"Tofu",
				"Genen Shouyu",
				"Pavlova",
				"Alice Mutton",
				"Carnarvon Tigers",
				"Teatime Chocolate Biscuits",
				"Sir Rodney's Marmalade",
			};

			string[] categories = new string[]
			{
				"Beverages",
				"Beverages",
				"Condiments",
				"Condiments",
				"Condiments",
				"Condiments",
				"Produce",
				"Condiments",
				"Meat/Poultry",
				"Seafood",
				"Dairy Products",
				"Dairy Products",
				"Seafood",
				"Produce",
				"Condiments",
				"Confections",
				"Meat/Poultry",
				"Seafood",
				"Confections",
				"Confections",
			};

			double[] prices = new double[]
			{
				11.5,
				38,
				25.45,
				16,
				21.95,
				19.85,
				31.45,
				44.5,
				21.85,
				1.65,
				0.8,
				9.95,
				11.5,
				14,
				13.5,
				8.9,
				21.9,
				6.75,
				4.3,
				17.6,
			};

			if (worksheet == null)
				worksheet = workbook.Worksheets.Add();

			int lastColumn = worksheet.Columns.Count - 1;
			int lastRow = worksheet.Rows.Count - 1;
			var style = worksheet.CellRanges[0, 0, lastColumn, lastRow].Style;
			style.FontName = "Calibri";
			style.FontSize = 11;
			style.PaddingLeft = "3pt";
			style.PaddingTop = "3pt";
			style.PaddingRight = "3pt";
			style.PaddingBottom = "3pt";
			style.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			var rowStyle = worksheet.Rows[0, lastRow].Style;
			rowStyle.Height = "17pt";

			worksheet.BeginInit();

			worksheet.Cells[1, 2].Data = "Product";
			worksheet.Cells[2, 2].Data = "Category";
			worksheet.Cells[3, 2].Data = "Price";

			for (int i = 0; i < products.Length; i++)
				worksheet.Cells[1, i + 3].Data = products[i];
			for (int i = 0; i < categories.Length; i++)
				worksheet.Cells[2, i + 3].Data = categories[i];
			for (int i = 0; i < prices.Length; i++)
				worksheet.Cells[3, i + 3].Data = prices[i];

			var titleStyle = worksheet.CellRanges[1, 2, 3, 2].Style;
			titleStyle.FontBold = true;
			titleStyle.Background = Brushes.LimeGreen;
			var categoriesStyle = worksheet.CellRanges[2, 3, 2, 22].Style;
			var pricesStyle = worksheet.CellRanges[3, 3, 3, 22].Style;
			pricesStyle.FontBold = true;
			pricesStyle.Format = "#,##0.00;(#,##0.00)";

			// Set some conditional formats
			var format1 = pricesStyle.ConditionalFormats.Add();
			format1.Type = ConditionalFormatType.CellValue;
			format1.Operator = ComparisonOperator.GreaterThan;
			format1.Second = "20";
			format1.Style.Background = new SolidColorBrush(Color.FromArgb(255, 255, 120, 85));
			format1.Style.Background.Freeze();

			var format2 = categoriesStyle.ConditionalFormats.Add();
			format2.Type = ConditionalFormatType.CellValue;
			format2.Operator = ComparisonOperator.Contain;
			format2.First = "Condiments";
			format2.Style.TextBrush = Brushes.DarkGoldenrod;
			format2.Style.FontItalic = true;

			worksheet.EndInit();

			workbookView.ResizeColumnsToFit(1, 3);
		}

		private void button_Click(object sender, RoutedEventArgs e)
		{
			if (workbookView.Selection.Range == null)
				return;

			var form = new ConditionalFormatForm(workbook, workbookView.Selection.Range.Style);
			if (form.ShowDialog() == true)
				form.Apply(workbookView.Selection.Range.Style);
		}
	}
}
