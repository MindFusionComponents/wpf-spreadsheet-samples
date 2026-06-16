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
	/// Interaction logic for MergingPanel.xaml
	/// </summary>
	public partial class MergingPanel : UserControl
	{
		public MergingPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();

			sheet.BeginInit();

			var style = sheet.CellRanges[0, 0, 0x400, 0x100000].Style;
			style.FontName = "Calibri";
			style.FontBold = true;
			style.FontSize = 11;
			style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			style.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;

			var mergeBrush = Brushes.Yellow;

			sheet.Cells["A1"].Data = "A1:D1";
			sheet.CellRanges["A1:D1"].Merge();
			sheet.CellRanges["A1:D1"].Style.Background = mergeBrush;

			sheet.Cells["A3"].Data = "A3:A7";
			sheet.CellRanges["A3:A7"].Merge();
			sheet.CellRanges["A3:A7"].Style.Background = mergeBrush;

			sheet.Cells["A9"].Data = "A9:C12";
			sheet.CellRanges["A9:C12"].Merge();
			sheet.CellRanges["A9:C12"].Style.Background = mergeBrush;

			sheet.EndInit();

			workbookView.Selection.Set(0, 0, 3, 0);
		}

		private void mergeUnmergeButton_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			var selectedRange = workbookView.Selection.Range;

			var merge = activeSheet.GetMergedCell(selectedRange.Left, selectedRange.Top);
			if (merge == null || merge.Left != selectedRange.Left || merge.Top != selectedRange.Top ||
				merge.Right != selectedRange.Right || merge.Bottom != selectedRange.Bottom)
			{
				if (!activeSheet.Merge(selectedRange))
					activeSheet.Unmerge(selectedRange);
			}
			else
			{
				activeSheet.Unmerge(selectedRange);
			}
		}
	}
}
