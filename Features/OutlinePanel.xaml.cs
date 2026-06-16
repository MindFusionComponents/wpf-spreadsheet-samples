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


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for OutlinePanel.xaml
	/// </summary>
	public partial class OutlinePanel : UserControl
	{
		public OutlinePanel()
		{
			InitializeComponent();

			string path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Outlines.xml");
			if (File.Exists(path))
				workbook.LoadFromXml(path);
		}

		private void groupRowsButton_Click(object sender, RoutedEventArgs e)
		{
			var selection = workbookView.Selection;
			using (workbook.StartChangeOperation(workbookView.ActiveWorksheet))
			{
				if (!workbookView.ActiveWorksheet.Rows[selection.Top, selection.Bottom].Group())
					MessageBox.Show("Unable to group the selected rows.", "Grouping", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}

		private void groupColumnsButton_Click(object sender, RoutedEventArgs e)
		{
			var selection = workbookView.Selection;
			using (workbook.StartChangeOperation(workbookView.ActiveWorksheet))
			{
				if (!workbookView.ActiveWorksheet.Columns[selection.Left, selection.Right].Group())
					MessageBox.Show("Unable to group the selected columns.", "Grouping", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}

		private void ungroupRowsButton_Click(object sender, RoutedEventArgs e)
		{
			var selection = workbookView.Selection;
			using (workbook.StartChangeOperation(workbookView.ActiveWorksheet))
			{
				if (!workbookView.ActiveWorksheet.Rows[selection.Top, selection.Bottom].Ungroup())
					MessageBox.Show("Unable to ungroup the selected rows.", "Grouping", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}

		private void ungroupColumnsButton_Click(object sender, RoutedEventArgs e)
		{
			var selection = workbookView.Selection;
			using (workbook.StartChangeOperation(workbookView.ActiveWorksheet))
			{
				if (!workbookView.ActiveWorksheet.Columns[selection.Left, selection.Right].Ungroup())
					MessageBox.Show("Unable to ungroup the selected columns.", "Grouping", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}
	}
}
