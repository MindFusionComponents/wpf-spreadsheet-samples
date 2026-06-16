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
	/// Interaction logic for PrintingPanel.xaml
	/// </summary>
	public partial class PrintingPanel : UserControl
	{
		public PrintingPanel()
		{
			InitializeComponent();

			string path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Printing.xml");
			if (File.Exists(path))
				workbook.LoadFromXml(path);

			workbook.PrintOptions.DocumentName = "Prices";
			workbook.PrintOptions.HeaderFormat = "%D\nPage %P";
			workbook.PrintOptions.EnableGridLines = true;
			workbook.PrintOptions.Scale = 140f;
		}

		private void previewButton_Click(object sender, RoutedEventArgs e)
		{
			(new WorkbookPrinter()).PrintPreview(workbook);
		}

		private void printButton_Click(object sender, RoutedEventArgs e)
		{
			(new WorkbookPrinter()).Print(workbook);
		}
	}
}
