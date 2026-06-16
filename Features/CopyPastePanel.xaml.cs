//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for CopyPastePanel.xaml
	/// </summary>
	public partial class CopyPastePanel : UserControl
	{
		public CopyPastePanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();
			sheet.BeginInit();
			var headingStyle = sheet.CellRanges["A1:E1"].Style;
			headingStyle.Background = Brushes.DarkSlateBlue;
			headingStyle.TextBrush = Brushes.White;
			headingStyle.FontBold = true;
			var contentStyle = sheet.CellRanges["A2:E10"].Style;
			contentStyle.Format = "0.##";
			contentStyle.Background = Brushes.LightSteelBlue;
			var random = new Random(DateTime.Now.Millisecond);
			for (int i = 0; i < 5; i++)
			{
				sheet.Cells[i, 0].Data = string.Format("Heading{0}", i + 1);
				for (int j = 1; j < 10; j++)
					sheet.Cells[i, j].Data = random.NextDouble();
			}
			sheet.EndInit();
		}

		private void copyButton_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			var selectionRange = workbookView.Selection.Range;
			var dataObject = new DataObject();

			// Copy data
			var cellData = activeSheet.CopyData(selectionRange);
			dataObject.SetData(cellData);

			// Copy styles
			var styleData = activeSheet.CopyStyles(selectionRange);
			dataObject.SetData(styleData);

			Clipboard.SetDataObject(dataObject, true);
		}

		private void pasteButton_Click(object sender, RoutedEventArgs e)
		{
			var activeSheet = workbookView.ActiveWorksheet;
			var dataObject = Clipboard.GetDataObject();

			// Paste data
			if (dataObject.GetDataPresent(typeof(CellData)))
			{
				var cellData = (CellData)dataObject.GetData(typeof(CellData));
				activeSheet.PasteData(workbookView.ActiveCell, cellData);
			}

			// Paste styles
			if (dataObject.GetDataPresent(typeof(StyleData)))
			{
				var styleData = (StyleData)dataObject.GetData(typeof(StyleData));
				activeSheet.PasteStyles(workbookView.ActiveCell, styleData);
			}
		}
	}
}
