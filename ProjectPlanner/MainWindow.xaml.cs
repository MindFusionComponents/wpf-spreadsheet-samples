//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;

using Microsoft.Win32;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.ProjectPlanner
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			workbookView.AllowMoveCells = false;
			workbookView.AllowMoveHeaders = false;

			workbookView.InplaceEditStarting += workbookView_InplaceEditStarting;
			workbook.WorksheetCellChanged += workbook_WorksheetCellChanged;
			workbook.WorksheetCellChanging += workbook_WorksheetCellChanging;

			CreateProjectTemplate();
			projectStart = new DateTime(DateTime.Today.Year, 1, 1);
			projectEnd = new DateTime(DateTime.Today.Year, 12, 31);
			workbook.Worksheets[0].Cells["B3"].Data = projectStart.ToLongDateString();

			ApplyDateValidation(workbook.Worksheets[0].CellRanges[1, 5, 2, 0xfffff].Style);
		}

		private void CreateProjectTemplate()
		{
			workbook.Worksheets.Clear();
			workbook.Worksheets.Add("Project");

			ignoreChanges = true;
			projectPhases = 0;

			var sheet = workbook.Worksheets[0];

			sheet.BeginInit();

			sheet.Columns[0].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(120);
			sheet.Columns[1, 2].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(50);
			sheet.Columns[3, 14].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(50);

			sheet.Rows[0].Height = MindFusion.Spreadsheet.Wpf.Measure.Point(19);

			sheet.CellRanges[1, 0, 14, 0].Merge();
			sheet.CellRanges[1, 1, 14, 1].Merge();
			sheet.CellRanges[1, 2, 14, 2].Merge();
			sheet.CellRanges[1, 3, 2, 3].Merge();
			sheet.CellRanges[3, 3, 5, 3].Merge();
			sheet.CellRanges[6, 3, 8, 3].Merge();
			sheet.CellRanges[9, 3, 11, 3].Merge();
			sheet.CellRanges[12, 3, 14, 3].Merge();
			sheet.CellRanges[1, 7, 14, 7].Merge();

			sheet.Columns[15, 0x3ff].IsHidden = true;
			sheet.Rows[8, 0xfffff].IsHidden = true;
			sheet.Rows[6].IsHidden = true;

			IStyle style = sheet.Rows[0, 2].Style;
			style.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;

			style = sheet.Rows[3, 4].Style;
			style.FontBold = true;
			style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;

			style = sheet.CellRanges[1, 0, 1, 1].Style;
			style.Background = Brushes.AntiqueWhite;

			sheet.Rows[4].Style.Background = Brushes.LightGray;
			sheet.Rows[7].Style.Background = Brushes.LightGray;
			sheet.Rows[7].Style.FontBold = true;

			sheet.Cells["B1"].Style.FontBold = true;
			sheet.Cells["B1"].Style.FontSize = 12;
			sheet.Cells["B2"].Style.FontItalic = true;
			sheet.CellRanges[0, 0, 0, 3].Style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			sheet.CellRanges[0, 0, 0, 3].Style.PaddingRight = MindFusion.Spreadsheet.Wpf.Measure.Point(5);

			sheet.CellRanges[1, 0, 1, 2].Style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			sheet.Cells["A8"].Style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;
			sheet.Cells["A8"].Style.PaddingRight = MindFusion.Spreadsheet.Wpf.Measure.Point(5);
			sheet.Cells["B8"].Style.PaddingLeft = MindFusion.Spreadsheet.Wpf.Measure.Point(5);

			ApplySimpleBorders(sheet.CellRanges[0, 4, 14, 7].Style);

			sheet.Cells["A1"].Data = "Project name";
			sheet.Cells["A2"].Data = "Project author";
			sheet.Cells["A3"].Data = "Project start";

			sheet.Cells["D4"].Data = "I Quarter";
			sheet.Cells["G4"].Data = "II Quarter";
			sheet.Cells["J4"].Data = "III Quarter";
			sheet.Cells["M4"].Data = "IV Quarter";

			sheet.Cells["A5"].Data = "Project phase";
			sheet.Cells["B5"].Data = "Start date";
			sheet.Cells["C5"].Data = "End date";

			sheet.CellRanges[0, 5, 2, 6].Style.Background = Brushes.AntiqueWhite;

			sheet.Cells["A8"].Data = "Project duration";

			DateTime d = new DateTime(DateTime.Today.Year, 1, 1);
			for (int c = 3; c <= 15; c++)
			{
				var cell = sheet.Cells[c, 4];
				cell.Data = d;
				cell.Style.Format = "MMM yyyy";
				d = d.AddMonths(1);
			}

			for (int c = 3; c <= 15; c++)
			{
				var headerCell = sheet.Cells[c, 4];
				var cell = sheet.Cells[c, 5];
				cell.Data = " ";
				var cf = cell.Style.ConditionalFormats.Add();
				cf.Type = ConditionalFormatType.Expression;
				cf.First = string.Format("=AND(NOT(ISBLANK(B6)),NOT(ISBLANK(C6)),MONTH(B6)<={0},MONTH(C6)>={0})", ((DateTime)headerCell.Data).Month);
				cf.Style.Background = Brushes.SkyBlue;
			}
			for (int c = 3; c <= 15; c++)
			{
				var headerCell = sheet.Cells[c, 4];
				var cell = sheet.Cells[c, 6];
				cell.Data = " ";
				var cf = cell.Style.ConditionalFormats.Add();
				cf.Type = ConditionalFormatType.Expression;
				cf.First = string.Format("=AND(NOT(ISBLANK(B7)),NOT(ISBLANK(C7)),MONTH(B7)<={0},MONTH(C7)>={0})", ((DateTime)headerCell.Data).Month);
				cf.Style.Background = Brushes.SkyBlue;
			}

			sheet.EndInit();

			ignoreChanges = false;
		}

		private int GetProjectLength()
		{
			DateTime min = DateTime.MaxValue;
			DateTime max = DateTime.MinValue;
			foreach (var cell in workbook.Worksheets[0].CellRanges[1, 5, 1, 0x100000])
			{
				if (cell.Data == null || !(cell.Value is DateTime))
					continue;
				var date = (DateTime)cell.Value;
				if (date <= min)
					min = date;
			}
			foreach (var cell in workbook.Worksheets[0].CellRanges[2, 5, 2, 0x100000])
			{
				if (cell.Data == null || !(cell.Value is DateTime))
					continue;
				var date = (DateTime)cell.Value;
				if (date >= max)
					max = date;
			}
			return (int)(max - min).TotalDays;
		}

		private bool EntryIsValid(int rowIndex)
		{
			var sheet = workbook.Worksheets[0];
			if (sheet.Cells[0, rowIndex].Data != null && sheet.Cells[1, rowIndex].Data != null && sheet.Cells[2, rowIndex].Data != null)
				return true;
			return false;
		}

		private void ApplySimpleBorders(IStyle style)
		{
			style.BorderBottomBrush = style.BorderLeftBrush =
				style.BorderRightBrush = style.BorderTopBrush = Brushes.Black;
			style.BorderBottomSize = style.BorderLeftSize =
				style.BorderRightSize = style.BorderTopSize = MindFusion.Spreadsheet.Wpf.Measure.Point(1);
			style.BorderBottomStyle = style.BorderLeftStyle =
				style.BorderRightStyle = style.BorderTopStyle = DashStyles.Solid;
		}

		private void ApplyDateValidation(IStyle style)
		{
			var val = style.Validation;
			val.Type = ValidationType.Date;
			val.AllowBlankCells = false;
			val.ErrorAction = ValidationErrorAction.Stop;
			val.ErrorMessage = string.Format("Enter a valid date in the range {0} - {1} in the cell.", projectStart.ToShortDateString(), projectEnd.ToShortDateString());
			val.ErrorTitle = "Error";
			val.First = projectStart.ToShortDateString();
			val.Second = projectEnd.ToShortDateString();
			val.Operator = ComparisonOperator.Between;
			val.ShowError = true;
		}

		private void workbookView_InplaceEditStarting(object sender, InplaceEditValidationEventArgs e)
		{
			if (ignoreChanges)
				return;
			Cell cell = e.Item as Cell;
			if (cell == null)
				return;

			var sheet = workbook.Worksheets[0];

			if (cell.Column == 1 &&
				(cell.Row == 0 || cell.Row == 1))
				return;

			if (cell.Column <= 2 &&
				cell.Row >= 5 &&
				cell.Row <= 5 + projectPhases)
				return;

			e.Cancel = true;
		}

		private void workbook_WorksheetCellChanging(object sender, CellValidationEventArgs e)
		{
			if (ignoreChanges)
				return;
			var sheet = workbook.Worksheets[0];

			// If the currency changed 
			if (e.Cell.Column == 1 &&
				(e.Cell.Row == 0 || e.Cell.Row == 1 || e.Cell.Row == 2))
				return;

			if (e.NewData != null &&
				e.Cell.Column <= 2 &&
				e.Cell.Row >= 5 &&
				e.Cell.Row <= 5 + projectPhases)
				return;

			e.Cancel = true;

		}

		private void workbook_WorksheetCellChanged(object sender, CellChangedEventArgs e)
		{
			if (ignoreChanges)
				return;
			ignoreChanges = true;
			var sheet = workbook.Worksheets[0];
			if (projectPhases + 5 == e.Cell.Row && EntryIsValid(e.Cell.Row))
			{
				sheet.Rows[projectPhases + 6].IsHidden = false;
				projectPhases++;
				if (projectPhases > 1)
				{
					sheet.BeginInit();
					sheet.Rows.Insert(e.Cell.Row + 1);
					sheet.CellRanges[0, e.Cell.Row + 1, 2, e.Cell.Row + 1].Style.Background = Brushes.AntiqueWhite;
					for (int c = 3; c <= 15; c++)
					{
						var headerCell = sheet.Cells[c, 4];
						var cell = sheet.Cells[c, e.Cell.Row + 1];
						cell.Data = " ";
						cell.Style.ConditionalFormats.Clear();
						var cf = cell.Style.ConditionalFormats.Add();
						cf.Type = ConditionalFormatType.Expression;
						cf.First = string.Format("=AND(NOT(ISBLANK(B{1})),NOT(ISBLANK(C{1})),MONTH(B{1})<={0},MONTH(C{1})>={0},YEAR(B{1})={2},YEAR(C{1})={2})", ((DateTime)headerCell.Data).Month, e.Cell.Row + 2, ((DateTime)headerCell.Data).Year);
						cf.Style.Background = Brushes.SkyBlue;
					}
					sheet.EndInit();

				}
				workbookView.Selection.Set(0, e.Cell.Row);
			}
			if (sheet.CellRanges[1, 5, 2, 0x100000].Count() >= 2)
				sheet.Cells[1, sheet.CellRanges[0, 0, 2, 0x100000].Max(c => c.Row)].Data = string.Format("{0} days", GetProjectLength());
			ignoreChanges = false;
		}

		private void printMenu_Click(object sender, RoutedEventArgs e)
		{
			var sheet = workbook.Worksheets[0];

			workbook.PrintOptions.Scale = 75;
			workbook.PrintOptions.HeaderFormat = string.Format("%D{0}", sheet.Cells["B2"].Value != null ?
				string.Format("\nby {0}", sheet.Cells["B2"].Value.ToString()) : "");

			var printer = new WorkbookPrinter();

			printer.Print(workbook);
			//PrintDialog dlg = new PrintDialog();
			//dlg.AllowPrintToFile = false;
			//dlg.AllowSomePages = true;
			//dlg.UseEXDialog = true;
			//dlg.Document = new PrintDocument();
			//dlg.Document.DocumentName = sheet.Cells["B1"].Value != null ? sheet.Cells["B1"].Value.ToString() : "Untitled";
			//dlg.Document.DefaultPageSettings.Landscape = true;
			//dlg.Document.PrinterSettings.PrintRange = PrintRange.AllPages;

			//if (dlg.ShowDialog(this) == DialogResult.OK)
			//    printer.Print(workbook, dlg.Document);
		}

		private void exportMenu_Click(object sender, RoutedEventArgs e)
		{
			var dlg = new SaveFileDialog();
			dlg.Filter = "PDF (*.pdf)|*.pdf";
			dlg.Title = "Export Project";
			if (dlg.ShowDialog(this) == true)
			{
				string ext = dlg.FileName.Substring(dlg.FileName.LastIndexOf('.')).ToLowerInvariant();
				try
				{
#if DEBUG
					new PdfExporter() { Scale = 70, PageOrientation = Pdf.PageOrientation.Portrait }.Export(workbook.Worksheets[0], dlg.FileName);
#else
					new PdfExporter() { Scale = 100, PageOrientation = Pdf.PageOrientation.Landscape }.Export(workbook.Worksheets[0], dlg.FileName);
#endif
				}
				catch (Exception ex)
				{
					MessageBox.Show(
						string.Format("Export failed.{0}{1}", Environment.NewLine, ex.ToString()),
						"Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
		}

		private void exitMenu_Click(object sender, RoutedEventArgs e)
		{
			Close();
		}


		private int projectPhases;
		private DateTime projectStart;
		private DateTime projectEnd;
		private bool ignoreChanges;
	}
}
