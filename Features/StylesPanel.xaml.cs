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
	/// Interaction logic for StylesPanel.xaml
	/// </summary>
	public partial class StylesPanel : UserControl
	{
		public StylesPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();

			var sheetStyle = sheet.CellRanges[0, 0, 0x400, 0x100000].Style;
			sheetStyle.FontName = "Calibri";
			sheetStyle.TextBrush = new SolidColorBrush(Color.FromArgb(0xff, 0x80, 0x80, 0x80));
			sheetStyle.TextBrush.Freeze();
			sheet.CellRanges[4, 4, 0x400, 0x100000].Style.Background = Brushes.WhiteSmoke;

			var rangeStyle = sheet.CellRanges[1, 1, 5, 5].Style;
			rangeStyle.FontSize = 12;
			rangeStyle.TextBrush = new SolidColorBrush(Color.FromArgb(0xff, 0x30, 0x30, 0x30));
			rangeStyle.Background = Brushes.NavajoWhite;
			rangeStyle.BorderBottomBrush = rangeStyle.BorderLeftBrush =
				rangeStyle.BorderRightBrush = rangeStyle.BorderTopBrush = Brushes.Red;

			var rowStyle = sheet.Rows[0, 3].Style;
			rowStyle.Background = Brushes.MintCream;
			rowStyle.FontUnderline = true;
			rowStyle.FontSize = 14;
			rowStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			rowStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			rowStyle.BorderBottomBrush = rowStyle.BorderLeftBrush =
				rowStyle.BorderRightBrush = rowStyle.BorderTopBrush = Brushes.Blue;

			var columnStyle = sheet.Columns[0, 3].Style;
			columnStyle.Background = Brushes.OldLace;
			columnStyle.FontItalic = true;
			columnStyle.FontSize = 14;
			columnStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			columnStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			columnStyle.BorderBottomBrush = columnStyle.BorderLeftBrush =
				columnStyle.BorderRightBrush = columnStyle.BorderTopBrush = Brushes.Orange;

			sheet.Columns[0].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(55);
			sheet.Columns[1, 5].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(70);
			sheet.Rows[0].Height = MindFusion.Spreadsheet.Wpf.Measure.Point(25);
			sheet.Rows[1, 5].Height = MindFusion.Spreadsheet.Wpf.Measure.Point(18);

			sheet.BeginInit();

			for (int c = 1; c <= 5; )
				sheet.Cells[c, 0].Data = string.Format("Column {0}", c++);

			for (int r = 1; r <= 5; )
				sheet.Cells[0, r].Data = string.Format("Row {0}", r++);

			for (int c = 1; c <= 5; c++)
				for (int r = 1; r <= 5; r++)
					sheet.Cells[c, r].Data = string.Format("=ADDRESS({0},{1},4)", r + 1, c + 1);

			sheet.EndInit();
		}
	}
}
