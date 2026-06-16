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

using MindFusion.Spreadsheet.Wpf.StandardForms;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for NumberFormatsPanel.xaml
	/// </summary>
	public partial class NumberFormatsPanel : UserControl
	{
		public NumberFormatsPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();

			sheet.BeginInit();

			sheet.Rows[1].Style.FontBold = true;

			sheet.Cells[1, 1].Data = "Plain Value";
			sheet.Cells[2, 1].Data = "Format";
			sheet.Cells[3, 1].Data = "Formatted Value";

			sheet.Cells[2, 2].Data = "'" + workbook.DefaultNumericFormats.PercentFormats[1].Format;
			sheet.Cells[2, 3].Data = "'" + workbook.DefaultNumericFormats.NumberFormats[7].Format;
			sheet.Cells[2, 4].Data = "'" + workbook.DefaultNumericFormats.NumberFormats[4].Format;
			sheet.Cells[2, 5].Data = "'" + workbook.DefaultNumericFormats.CurrencyFormats[1].Format;
			sheet.Cells[2, 6].Data = "'" + workbook.DefaultNumericFormats.DateFormats[0].Format;
			sheet.Cells[2, 7].Data = "'" + workbook.DefaultNumericFormats.DateFormats[2].Format;
			sheet.Cells[2, 8].Data = "'" + workbook.DefaultNumericFormats.ScientificFormats[0].Format;
			sheet.Cells[2, 9].Data = "'" + workbook.DefaultNumericFormats.FractionFormats[1].Format;
			sheet.Cells[2, 10].Data = "???.???";
			sheet.Cells[2, 11].Data = "Boolean";

			for (int r = 2; r <= 11; r++)
			{
				sheet.Cells[1, r].Data = Math.Pow(r, 2) * r * 0.1;
				sheet.Cells[3, r].Data = "=B" + (r + 1).ToString();
				sheet.Cells[3, r].Style.Format = sheet.Cells[2, r].Value != null ? sheet.Cells[2, r].Value.ToString().TrimStart(new char[] { '\'' }) : "General";
			}

			sheet.Columns[2, 3].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(150);
			sheet.Columns[2].Style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Left;

			sheet.EndInit();
		}

		private void workbook_WorksheetCellChanged(object sender, CellChangedEventArgs e)
		{
			if (e.Cell.Column == 2 && e.Cell.Row >= 2 && e.Cell.Row <= 11)
			{
				workbook.Worksheets[0].Cells[3, e.Cell.Row].Style.Format = e.Cell.Value != null ? e.Cell.Value.ToString() : "General";
			}
		}

		private void workbook_WorksheetCellsCleared(object sender, WorksheetEventArgs e)
		{
			for (int r = 2; r <= 11; r++)
			{
				e.Worksheet.Cells[3, r].Style.Format = e.Worksheet.Cells[2, r].Value != null ? e.Worksheet.Cells[2, r].Value.ToString() : "General";
			}
		}

		private void showButton_Click(object sender, RoutedEventArgs e)
		{
			FormatForm form = new FormatForm(workbook, workbookView.Selection.Range.Style);
			if (form.ShowDialog() == true)
				form.Apply(workbookView.Selection.Range.Style);
		}
	}
}
