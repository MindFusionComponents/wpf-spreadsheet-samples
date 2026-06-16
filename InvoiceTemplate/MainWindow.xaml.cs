//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Media;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.InvoiceTemplate
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			workbookView.ShowGridLines = false;
			workbookView.AllowAutoFill = false;
			// TODO:
			//workbookView.ShowTabs = false;
			workbookView.AllowMoveCells = false;
			workbookView.AllowMoveHeaders = false;
			workbookView.DrawFilledCellBorders = true;

			var activeSheet = workbook.Worksheets.Add();
			workbookView.ActiveWorksheet = activeSheet;
			activeSheet.BeginInit();

			var globalStyle = activeSheet.CellRanges[0, 0, activeSheet.Columns.Count - 1, activeSheet.Rows.Count - 1].Style;
			globalStyle.FontName = "Tahoma";
			globalStyle.FontSize = 8;

			var heading = activeSheet.CellRanges["A1:G1"];
			heading.Merge();
			activeSheet.Cells["A1"].Data = "Invoice";
			var headingStyle = heading.Style;
			headingStyle.FontSize = 14;
			headingStyle.FontBold = true;
			headingStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			headingStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			headingStyle.PaddingRight = "5pt";
			headingStyle.TextBrush = Brushes.DarkSlateBlue;
			headingStyle.BorderBottomBrush = Brushes.DarkSlateBlue;
			headingStyle.BorderBottomSize = "3pt";

			var address = activeSheet.CellRanges["A3:C3"];
			address.Merge();
			activeSheet.Cells["A3"].Data = "(address)";
			var addressStyle = address.Style;
			addressStyle.TextBrush = Brushes.Gray;
			addressStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			addressStyle.FontItalic = true;
			addressStyle.FontSize = 7.5;

			var date = activeSheet.Cells["F3"];
			date.Data = "Date:";
			var dateStyle = date.Style;
			dateStyle.FontBold = true;
			dateStyle.TextBrush = Brushes.DarkSlateBlue;
			dateStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			dateStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			date = activeSheet.Cells["G3"];
			date.Data = DateTime.Now;
			dateStyle = date.Style;
			dateStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Left;
			dateStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			dateStyle.Format = "dd-MMM-yy";

			var shipTo = activeSheet.Cells["A5"];
			shipTo.Data = "Ship to:";
			var shipToStyle = shipTo.Style;
			shipToStyle.FontBold = true;
			shipToStyle.TextBrush = Brushes.DarkSlateBlue;
			shipToStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			shipToStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;

			activeSheet.CellRanges["B5:C5"].Merge();
			activeSheet.Cells["B5"].Data = "(ship name)";
			activeSheet.CellRanges["B6:C6"].Merge();
			activeSheet.Cells["B6"].Data = "(ship address)";
			activeSheet.CellRanges["B7:C7"].Merge();
			activeSheet.Cells["B7"].Data = "(ship region)";
			activeSheet.CellRanges["B8:C8"].Merge();
			activeSheet.Cells["B8"].Data = "(ship country)";
			shipToStyle = activeSheet.CellRanges["B5:C8"].Style;
			shipToStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			shipToStyle.TextBrush = Brushes.Gray;

			var billTo = activeSheet.Cells["D5"];
			billTo.Data = "Bill to:";
			var billToStyle = billTo.Style;
			billToStyle.FontBold = true;
			billToStyle.TextBrush = Brushes.DarkSlateBlue;
			billToStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			billToStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;

			activeSheet.CellRanges["E5:F5"].Merge();
			activeSheet.Cells["E5"].Data = "(bill name)";
			activeSheet.CellRanges["E6:F6"].Merge();
			activeSheet.Cells["E6"].Data = "(bill address)";
			activeSheet.CellRanges["E7:F7"].Merge();
			activeSheet.Cells["E7"].Data = "(bill region)";
			activeSheet.CellRanges["E8:F8"].Merge();
			activeSheet.Cells["E8"].Data = "(bill country)";
			shipToStyle = activeSheet.CellRanges["E5:F8"].Style;
			shipToStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			shipToStyle.TextBrush = Brushes.Gray;

			string[] orderTexts = new string[] { "Order ID", "Customer ID", "Employee ID",
				"Order Date", "Required Date", "Shipped Date", "Ship Via" };
			for (int i = 0; i < orderTexts.Length; i++)
				activeSheet.Cells[i, 9].Data = orderTexts[i];
			var orderHeadingStyle = activeSheet.CellRanges[0, 9, 6, 9].Style;
			orderHeadingStyle.BorderLeftBrush = Brushes.DarkBlue;
			orderHeadingStyle.BorderTopBrush = Brushes.DarkBlue;
			orderHeadingStyle.BorderRightBrush = Brushes.DarkBlue;
			orderHeadingStyle.BorderBottomBrush = Brushes.DarkBlue;
			orderHeadingStyle.Background = Brushes.DarkBlue;
			orderHeadingStyle.TextBrush = Brushes.White;
			orderHeadingStyle.FontBold = true;
			orderHeadingStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			orderHeadingStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Bottom;

			string[] orderPlaceholders = new string[] { "(order id)", "(customer id)", "(employee id)",
				"(order date)", "(required date)", "(shipped date)", "(ship via)" };
			for (int i = 0; i < orderPlaceholders.Length; i++)
				activeSheet.Cells[i, 10].Data = orderPlaceholders[i];
			var orderContentStyle = activeSheet.CellRanges[0, 10, 6, 10].Style;
			orderContentStyle.BorderLeftBrush = Brushes.DarkBlue;
			orderContentStyle.BorderTopBrush = Brushes.DarkBlue;
			orderContentStyle.BorderRightBrush = Brushes.DarkBlue;
			orderContentStyle.BorderBottomBrush = Brushes.DarkBlue;
			orderContentStyle.TextBrush = Brushes.Gray;
			orderContentStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			orderContentStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			activeSheet.CellRanges["D11:F11"].Style.Format = "dd-MMM-yy";

			string[] productTexts = new string[] { "Product ID", "Product Name", "Product Name",
				"Quantity", "Unit Price", "Discount", "Extended Price" };
			for (int i = 0; i < productTexts.Length; i++)
				activeSheet.Cells[i, 12].Data = productTexts[i];
			var productHeadingStyle = activeSheet.CellRanges[0, 12, 6, 12].Style;
			productHeadingStyle.BorderTopBrush = Brushes.DarkBlue;
			productHeadingStyle.BorderBottomBrush = Brushes.DarkBlue;
			productHeadingStyle.Background = Brushes.DarkBlue;
			productHeadingStyle.TextBrush = Brushes.White;
			productHeadingStyle.FontBold = true;
			productHeadingStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Left;
			productHeadingStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Bottom;
			activeSheet.CellRanges[1, 12, 2, 12].Merge();

			AddProduct();

			var total = activeSheet.Cells["F15"];
			total.Data = "Total:";
			var totalStyle = total.Style;
			totalStyle.FontBold = true;
			totalStyle.TextBrush = Brushes.DarkSlateBlue;
			totalStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			totalStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			total = activeSheet.Cells["G15"];
			total.Data = "=SUM(G14)";
			totalStyle = total.Style;
			totalStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Left;
			totalStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			totalStyle.Format = "[$$-409]#,##0.00;([$$-409]#,##0.00)";

			activeSheet.Rows[0].Height = "48pt";
			activeSheet.Rows[1].Height = "5pt";
			activeSheet.Rows[2].Height = "32pt";
			activeSheet.Rows[3].Height = "5pt";
			activeSheet.Rows[8].Height = "20pt";
			activeSheet.Rows[9].Height = "20pt";
			activeSheet.Rows[10].Height = "18pt";
			activeSheet.Rows[12].Height = "20pt";
			activeSheet.Rows[13, 14].Height = "18pt";
			activeSheet.Columns[0, 6].Width = "72pt";

			activeSheet.EndInit();

			workbook.WorksheetCellChanging += OnWorksheetCellChanging;
			workbook.WorksheetCellChanged += OnWorksheetCellChanged;
			workbookView.InplaceEditStarting += OnInplaceEditStarting;
		}

		private void OnWorksheetCellChanging(object sender, CellValidationEventArgs e)
		{
			e.Cancel = !CanEdit(e.Cell);
		}

		private void OnWorksheetCellChanged(object sender, CellEventArgs e)
		{
			e.Cell.Style.TextBrush = Brushes.Black;
		}

		private void OnInplaceEditStarting(object sender, InplaceEditValidationEventArgs e)
		{
			if (e.Item is Cell)
			{
				var cell = e.Item as Cell;
				e.Cancel = !CanEdit(cell);
			}
		}

		private bool CanEdit(Cell cell)
		{
			int c = cell.Column;
			int r = cell.Row;

			if (r == 2 && c == 0 ||
				r == 2 && c == 6 ||
				r == 10 && c <= 6 ||
				r >= 13 && r <= 12 + products && c <= 5 ||
				c == 1 && r >= 4 && r <= 7 ||
				c == 4 && r >= 4 && r <= 7)
			{
				return true;
			}

			return false;
		}

		private void AddProduct()
		{
			var activeSheet = workbookView.ActiveWorksheet;

			activeSheet.BeginInit();
			string[] productPlaceholders = new string[] { "(product id)", "(product name)", "(product name)",
				"(quantity)", "(unit price)", "(discount)", };
			for (int i = 0; i < productPlaceholders.Length; i++)
				activeSheet.Cells[i, 12 + products].Data = productPlaceholders[i];
			activeSheet.Cells[6, 12 + products].SetData("=IFERROR(D14*E14*(1-F14), 0)", 0, products - 1);
			var productContentStyle = activeSheet.CellRanges[0, 12 + products, 6, 12 + products].Style;
			productContentStyle.BorderBottomBrush = Brushes.DarkBlue;
			productContentStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Left;
			productContentStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			productContentStyle = activeSheet.CellRanges[0, 12 + products, 5, 12 + products].Style;
			productContentStyle.TextBrush = Brushes.Gray;
			activeSheet.CellRanges[1, 12 + products, 2, 12 + products].Merge();
			activeSheet.Cells[4, 12 + products].Style.Format = "[$$-409]#,##0.00;([$$-409]#,##0.00)";
			activeSheet.Cells[5, 12 + products].Style.Format = workbook.DefaultNumericFormats.PercentFormats[0].Format;
			activeSheet.Cells[6, 12 + products].Style.Format = "[$$-409]#,##0.00;([$$-409]#,##0.00)";
			activeSheet.EndInit();
		}

		private void addButton_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			activeSheet.Rows.Insert(13 + products);

			products++;
			AddProduct();

			// Need to update the total to include the new product
			activeSheet.Cells[6, 13 + products].Data = string.Format("=SUM(G14:G{0})", 13 + products);
		}

		private void previewButton_Click(object sender, RoutedEventArgs e)
		{
			var printer = new WorkbookPrinter();
			printer.PrintPreview(workbook);
		}


		private int products = 1;
	}
}
