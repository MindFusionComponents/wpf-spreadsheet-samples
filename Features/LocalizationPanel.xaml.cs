//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MindFusion.Spreadsheet.Wpf.StandardForms;
using Microsoft.Win32;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for LocalizationPanel.xaml
	/// </summary>
	public partial class LocalizationPanel : UserControl
	{
		public LocalizationPanel()
		{
			InitializeComponent();

			languagesFolder = FindDirectoryLocation("Localization");

			var sheet = workbook.Worksheets.Add();

			sheet.BeginInit();

			sheet.Columns[2, 3].Width = MindFusion.Spreadsheet.Wpf.Measure.Point(80);
			sheet.Rows[1].Style.FontBold = true;

			sheet.Cells[1, 1].Data = "Type";
			sheet.Cells[2, 1].Data = "Implicit format";
			sheet.Cells[3, 1].Data = "Explicit format";

			sheet.Cells[1, 2].Data = "Integer";
			sheet.Cells[2, 2].Data = 42;
			sheet.Cells[3, 2].Data = 42;
			sheet.Cells[3, 2].Style.Format = "0";

			sheet.Cells[1, 3].Data = "Float";
			sheet.Cells[2, 3].Data = 12.34f;
			sheet.Cells[3, 3].Data = 12.34f;
			sheet.Cells[3, 3].Style.Format = "0.00";

			sheet.Cells[1, 4].Data = "Double";
			sheet.Cells[2, 4].Data = 1234.56;
			sheet.Cells[3, 4].Data = 1234.56;
			sheet.Cells[3, 4].Style.Format = "#,##0.00";

			sheet.Cells[1, 5].Data = "Currency";
			sheet.Cells[2, 5].Data = 15.99.ToString("C");
			sheet.Cells[3, 5].Data = 15.99;
			sheet.Cells[3, 5].Style.Format = workbook.DefaultNumericFormats.CurrencyFormats[1].Format;

			sheet.Cells[1, 6].Data = "Percentage";
			sheet.Cells[2, 6].Data = "75%";
			sheet.Cells[3, 6].Data = 0.75;
			sheet.Cells[3, 6].Style.Format = workbook.DefaultNumericFormats.PercentFormats[1].Format;

			sheet.Cells[1, 7].Data = "Boolean";
			sheet.Cells[2, 7].Data = true;
			sheet.Cells[3, 7].Data = 1;
			sheet.Cells[3, 7].Style.Format = "Boolean";

			sheet.Cells[1, 8].Data = "Date";
			sheet.Cells[2, 8].Data = DateTime.Today;
			sheet.Cells[3, 8].Data = (DateTime.Today - new DateTime(1899, 12, 30)).TotalDays;
			sheet.Cells[3, 8].Style.Format = workbook.DefaultNumericFormats.DateFormats[0].Format;

			var now = DateTime.Now;
			sheet.Cells[1, 9].Data = "Time";
			sheet.Cells[2, 9].Data = new TimeSpan(now.Hour, now.Minute, now.Second);
			sheet.Cells[3, 9].Data = sheet.Cells[2, 9].Data;
			sheet.Cells[3, 9].Style.Format = workbook.DefaultNumericFormats.TimeFormats[1].Format;

			sheet.EndInit();

			var cultures = CultureInfo.GetCultures(CultureTypes.SpecificCultures).Where(ci => !ci.IsNeutralCulture).ToList();
			locale.ItemsSource = cultures;
			locale.SelectedItem = workbook.Locale;

			locale.SelectionChanged += new SelectionChangedEventHandler(locale_SelectionChanged);
		}

		private void locale_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (locale.SelectedIndex == -1)
				return;

			workbook.Locale = (CultureInfo)locale.SelectedItem;

			// Try to localize standard forms too
			var name = workbook.Locale.Name;
			int dashIndex = name.IndexOf('-');
			if (dashIndex != -1)
				name = name.Substring(0, dashIndex);
			name = name.ToUpper();

			var file = GetLocalizationFile(name);
			if (file != null && File.Exists(file))
				workbook.SetLocalizationInfo(file);
		}

		private void ContextMenu_Click(object sender, RoutedEventArgs e)
		{
			var menu = sender as ContextMenu;
			var index = menu.Items.IndexOf(e.OriginalSource);
			switch (index)
			{
				case 0:
					new ConditionalFormatForm(workbook, workbookView.Selection.Range.Style).ShowDialog();
					break;
				case 1:
					new DeleteCellsForm(workbook).ShowDialog();
					break;
				case 2:
					new FillSeriesForm(workbook).ShowDialog();
					break;
				case 3:
					new FindReplaceForm(workbook, workbookView.ActiveCell).ShowDialog();
					break;
				case 4:
					new FormatForm(workbook, workbookView.Selection.Range.Style).ShowDialog();
					break;
				case 5:
					new HyperlinkForm(workbook).ShowDialog();
					break;
				case 6:
					new InsertCellsForm(workbook).ShowDialog();
					break;
				case 7:
					new InsertWorksheetForm(workbook).ShowDialog();
					break;
				case 8:
					new ManageNamedRangesForm(workbook, workbookView.Selection.Range).ShowDialog();
					break;
				case 9:
					new SortForm(workbookView.ActiveWorksheet, workbookView.Selection.Range).ShowDialog();
					break;
				case 10:
					new ValidationForm(workbook, workbookView.Selection.Range.Style.Validation).ShowDialog();
					break;
				case 11:
					new WorksheetRenameForm(workbook, "New name").ShowDialog();
					break;
			}
		}

		static string FindDirectoryLocation(string dirName)
		{
			var dir = new DirectoryInfo(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location));
			return FindDirectoryLocation(dir, dirName);
		}

		static string FindDirectoryLocation(DirectoryInfo dir, string dirName)
		{
			var subdirs = dir.GetDirectories(dirName);
			if (subdirs == null || subdirs.Length == 0)
			{
				var parent = dir.Parent;
				if (parent != null)
					return FindDirectoryLocation(parent, dirName);
				return null;
			}

			return subdirs[0].FullName;
		}

		string GetLocalizationFile(string lang)
		{
			if (!Directory.Exists(languagesFolder))
			{
				if (MessageBox.Show("The localization file could not be found. Would you like to search for it manually?",
					"File not found", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
					return null;

				var dialog = new OpenFileDialog();
				dialog.Filter = "Localization Files (*.xml)|*.xml|All Files|*.*||";
				if (dialog.ShowDialog() != true)
					return null;

				return dialog.FileName;
			}

			return Path.Combine(languagesFolder, string.Format("Localization.{0}.xml", lang));
		}


		string languagesFolder;
	}
}
