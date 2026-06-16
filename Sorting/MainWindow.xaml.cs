//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

using Microsoft.Win32;
using MindFusion.Spreadsheet.Wpf.StandardForms;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Sorting
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			workbook.Worksheets.Add();
		}

		private void exitMenu_Click(object sender, RoutedEventArgs e)
		{
			Close();
		}

		private void generateMenu_Click(object sender, RoutedEventArgs e)
		{
			// Generate random data
			var random = new Random(DateTime.Now.Millisecond);
			var activeSheet = workbookView.ActiveWorksheet;

			activeSheet.BeginInit();
			for (int c = 1; c <= 20; c++)
			{
				for (int r = 1; r <= 50; r++)
					activeSheet.Cells[c, r].Data = random.Next() % 1000;
			}
			for (int c = 1; c <= 20; c++)
				activeSheet.Cells[c, 0].Data = string.Format("Heading {0}", c);
			for (int r = 1; r <= 50; r++)
				activeSheet.Cells[0, r].Data = string.Format("Entry {0}", r);
			activeSheet.CellRanges[1, 0, 20, 0].Style.Background = Brushes.LightSteelBlue;
			activeSheet.CellRanges[0, 1, 0, 50].Style.Background = Brushes.LightSteelBlue;
			activeSheet.EndInit();

			// Select the generated data
			workbookView.Selection.Set(1, 1, 20, 50);
		}

		private void importMenu_Click(object sender, RoutedEventArgs e)
		{
			var dialog = new OpenFileDialog();
			dialog.Filter = "Microsoft Excel (*.xlsx)|*.xlsx|OpenDocument (*.ods)|*.ods|Comma-separated values (*.csv)|*.csv|All files|*.*||";
			if (dialog.ShowDialog(this) == true)
			{
				string file = dialog.FileName;
				string extension = Path.GetExtension(file).ToLower();
				if (extension == ".xlsx")
				{
					var excelImporter = new ExcelImporter();
					excelImporter.Import(file, workbook);
				}
				else if (extension == ".ods")
				{
					var calcImporter = new CalcImporter();
					calcImporter.Import(file, workbook);
				}
				else if (extension == ".csv")
				{
					var activeSheet = workbookView.ActiveWorksheet;
					if (activeSheet == null)
						activeSheet = workbook.Worksheets.Add();

					var csvImporter = new CsvImporter();
					csvImporter.Import(file, activeSheet);
				}
				else
				{
					MessageBox.Show("Unsupported file format.");
				}
			}
		}

		private void pasteMenu_Click(object sender, RoutedEventArgs e)
		{
			//workbookView.Paste();
			ApplicationCommands.Paste.Execute(null, workbookView);
		}

		private void horizontallyMenu_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			if (activeSheet == null)
				activeSheet = workbook.Worksheets.Add();
			var sortOptions = new SortOptions();
			sortOptions.Direction = SortDirection.LeftToRight;
			sortOptions.Keys.Add(new SortKey(workbookView.ActiveCell.Row));
			activeSheet.Sort(workbookView.Selection.Range, sortOptions);
		}

		private void verticallyMenu_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			if (activeSheet == null)
				activeSheet = workbook.Worksheets.Add();
			var sortOptions = new SortOptions();
			sortOptions.Direction = SortDirection.TopToBottom;
			sortOptions.Keys.Add(new SortKey(workbookView.ActiveCell.Column));
			activeSheet.Sort(workbookView.Selection.Range, sortOptions);
		}

		private void optionsMenu_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			if (activeSheet == null)
				activeSheet = workbook.Worksheets.Add();
			var sortForm = new SortForm(activeSheet, workbookView.Selection.Range);
			if (sortForm.ShowDialog() == true)
				sortForm.Sort();
		}
	}
}
