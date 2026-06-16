//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Windows;
using System.Windows.Media;
using System.Globalization;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Calendar
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			var worksheet = workbook.Worksheets.Add();

			workbookView.DrawFilledCellBorders = true;
			workbookView.AllowMoveCells = false;
			workbookView.AllowMoveHeaders = false;

			worksheet.BeginInit();
			var heading = worksheet.CellRanges["B2:H2"];
			heading.Merge();
			var headingCell = worksheet.Cells["B2"];
			headingCell.Data = DateTime.Today;
			var headingStyle = headingCell.Style;
			headingStyle.Format = "MMMM, yyyy";
			headingStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			headingStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			headingStyle.Background = Brushes.SteelBlue;
			headingStyle.FontSize = 14;
			headingStyle.FontBold = true;
			headingStyle.TextBrush = Brushes.White;
			var daysOfWeek = new CultureInfo("en-US").DateTimeFormat.AbbreviatedDayNames;
			for (int i = 0; i < 7; i++)
				worksheet.Cells[i + 1, 2].Data = daysOfWeek[i];
			var subheading = worksheet.CellRanges["B3:H3"];
			var subheadingStyle = subheading.Style;
			subheadingStyle.Background = new SolidColorBrush(Color.FromArgb(255, 245, 230, 205));
			subheadingStyle.Background.Freeze();
			subheadingStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			subheadingStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			subheadingStyle.FontBold = true;
			var content = worksheet.CellRanges["B4:H9"];
			var contentStyle = content.Style;
			contentStyle.Background = new SolidColorBrush(Color.FromArgb(255, 255, 250, 220));
			contentStyle.Background.Freeze();
			contentStyle.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Center;
			contentStyle.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;
			contentStyle.Format = "D";
			var formula = "=IF(MONTH(DATE(YEAR(B2),MONTH(B2),1))<>MONTH(DATE(YEAR(B2),MONTH(B2),1)-(WEEKDAY(DATE(YEAR(B2),MONTH(B2),1))-1)+(ROW()-4)*7+(COLUMN()-1)-1),\"\",DATE(YEAR(B2),MONTH(B2),1)-(WEEKDAY(DATE(YEAR(B2),MONTH(B2),1))-1)+(ROW()-4)*7+(COLUMN()-1)-1)";
			for (int i = 0; i < 7; i++)
			{
				for (int j = 0; j < 6; j++)
					worksheet.Cells[content.Left + i, content.Top + j].Data = formula;
			}
			worksheet.Columns[1, 7].Width = "30pt";
			worksheet.Rows[1].Height = "20pt";
			worksheet.Rows[2, 8].Height = "16pt";
			worksheet.EndInit();
		}
	}
}
