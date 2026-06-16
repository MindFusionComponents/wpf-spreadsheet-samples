//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for FreezingPanel.xaml
	/// </summary>
	public partial class FreezingPanel : UserControl
	{
		public FreezingPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();

			sheet.BeginInit();

			sheet.Columns[0, 6].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(80);
			sheet.Rows[0].Height = MindFusion.Spreadsheet.Wpf.Measure.Point(20);
			sheet.Rows[202].Height = MindFusion.Spreadsheet.Wpf.Measure.Point(20);

			sheet.Cells["A1"].Data = "Order Number";
			sheet.Cells["B1"].Data = "Invoice Number";
			sheet.Cells["C1"].Data = "Unit Price";
			sheet.Cells["D1"].Data = "Quantity";
			sheet.Cells["E1"].Data = "Subtotal";
			sheet.Cells["F1"].Data = "VAT";
			sheet.Cells["G1"].Data = "Total";

			for (int i = 0; i <= 200; i++)
			{
				string row = (i + 2).ToString();
				sheet.Cells["A" + row].Data = "'" + (i + 1).ToString("D8");
				sheet.Cells["B" + row].Data = string.Format("'201402{0}", (i + 1).ToString("D5"));
				sheet.Cells["C" + row].Data = (rnd.Next(2, 29) + rnd.Next(12, 99) % 3 == 0 ? 0.49 : 0.99).ToString("C");
				sheet.Cells["D" + row].Data = rnd.Next(1, 5);
				sheet.Cells["E" + row].Data = string.Format("=C{0}*D{0}", row);
				sheet.Cells["F" + row].Data = string.Format("=E{0}*0.2", row);
				sheet.Cells["G" + row].Data = string.Format("=E{0}*1.2", row);
			}

			sheet.Cells["A203"].Data = "February Orders";
			sheet.Cells["G203"].Data = "=SUM(G2:G202)";

			sheet.CellRanges["A203:F203"].Merge();

			sheet.FrozenRowCount = 1;

			IStyle style = sheet.Rows[0].Style;
			style.FontBold = true;
			style.FontSize = 10;
			style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			style.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;

			style = sheet.Rows[202].Style;
			style.FontBold = true;
			style.FontItalic = true;
			style.FontSize = 10;
			style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			style.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			style.PaddingRight = MindFusion.Spreadsheet.Wpf.Measure.Point(5);

			style = sheet.CellRanges[0, 0, 6, 202].Style;
			style.BorderBottomBrush = style.BorderLeftBrush = style.BorderRightBrush = style.BorderTopBrush = SystemColors.ActiveBorderBrush;

			style = sheet.CellRanges[0, 0, 6, 0].Style;
			style.BorderBottomSize = style.BorderLeftSize = style.BorderRightSize = style.BorderTopSize = MindFusion.Spreadsheet.Wpf.Measure.Point(2);
			style = sheet.CellRanges[0, 202, 6, 202].Style;
			style.BorderBottomSize = style.BorderLeftSize = style.BorderRightSize = style.BorderTopSize = MindFusion.Spreadsheet.Wpf.Measure.Point(2);

			sheet.EndInit();

			workbookView.ScrollY = 190;
		}

		private void freezeUnfreezeButton_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			var activeCell = workbookView.ActiveCell;
			if (activeSheet.FrozenColumnCount > 0 || activeSheet.FrozenRowCount > 0)
				workbookView.Unfreeze();
			else
				workbookView.Freeze();
		}


		private static readonly Random rnd = new Random(DateTime.Now.Millisecond);
	}
}
