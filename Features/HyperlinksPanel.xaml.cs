//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;
using MindFusion.Spreadsheet.Wpf.StandardForms;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for HyperlinksPanel.xaml
	/// </summary>
	public partial class HyperlinksPanel : UserControl
	{
		public HyperlinksPanel()
		{
			InitializeComponent();

			workbookView.HyperlinkClicked += new EventHandler<CellEventArgs>(workbookView_HyperlinkClicked);
			workbookView.ActiveCellChanged += new EventHandler(workbookView_ActiveCellChanged);

			var worksheet = workbook.Worksheets.Add();

			var cell = worksheet.Cells[1, 1];
			cell.Data = "Created with Cell.Hyperlink";
			cell.SetHyperlink(HyperlinkType.WebAddress, "http://www.mindfusion.eu", "");

			cell = worksheet.Cells[1, 2];
			cell.Data = "=HYPERLINK(\"http://mindfusion.eu/Forum/YaBB.pl\", \"Created with the HYPERLINK function\")";
		}

		void workbookView_ActiveCellChanged(object sender, EventArgs e)
		{
			editButton.IsEnabled = workbookView.ActiveCell.Hyperlink != null;
		}

		void workbookView_HyperlinkClicked(object sender, CellEventArgs e)
		{
			if (Keyboard.Modifiers == ModifierKeys.Control)
			{
				try
				{
					System.Diagnostics.Process.Start(e.Cell.Hyperlink.Target);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
		}

		private void addRemoveButton_Click(object sender, RoutedEventArgs e)
		{
			var cell = workbookView.ActiveCell;
			if (cell == null)
				return;

			if (cell.Hyperlink == null)
			{
				HyperlinkForm hyperlinkForm = new HyperlinkForm(workbook, cell.Value != null ? cell.Value.ToString() : "");
				if (hyperlinkForm.ShowDialog() == true)
				{
					cell.SetHyperlink(hyperlinkForm.HyperlinkType, hyperlinkForm.Target, hyperlinkForm.SubTarget);
					cell.Hyperlink.ToolTip = hyperlinkForm.HyperlinkToolTip;
					editButton.IsEnabled = true;
				}
			}
			else
			{
				cell.RemoveHyperlink();
				editButton.IsEnabled = false;
			}
		}

		private void editButton_Click(object sender, RoutedEventArgs e)
		{
			var cell = workbookView.ActiveCell;
			if (cell == null || cell.Hyperlink == null)
				return;

			HyperlinkForm hyperlinkForm = new HyperlinkForm(workbook, cell.Hyperlink);
			if (hyperlinkForm.ShowDialog() == true)
			{
				cell.Hyperlink.Type = hyperlinkForm.HyperlinkType;
				cell.Hyperlink.Target = hyperlinkForm.Target;
				cell.Hyperlink.SubTarget = hyperlinkForm.SubTarget;
				cell.Hyperlink.ToolTip = hyperlinkForm.HyperlinkToolTip;
			}
		}
	}
}
