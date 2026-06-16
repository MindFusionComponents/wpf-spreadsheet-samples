//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using MindFusion.Spreadsheet.Wpf.StandardForms;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for DataValidationPanel.xaml
	/// </summary>
	public partial class DataValidationPanel : UserControl
	{
		public DataValidationPanel()
		{
			InitializeComponent();

			var worksheet = workbook.Worksheets.Add();

			var cell = worksheet.Cells[0, 0];
			cell.Data = "Select a value from the list in A2";

			worksheet.CellRanges[0, 0, 2, 0].Merge();

			cell = worksheet.Cells[0, 1];
			cell.Data = 1;
			cell.Validation.Type = ValidationType.List;
			cell.Validation.First = "1,2,3,4";
			cell.Validation.AllowBlankCells = false;
			cell.Validation.ShowDropdownList = true;
			cell.Validation.ShowInputMessage = true;
			cell.Validation.InputMessage = "Select a value from the list";
			cell.Validation.ShowError = true;
			cell.Validation.ErrorTitle = "Invalid input.";
			cell.Validation.ErrorMessage = "Enter a value from the specified list.";

			cell = worksheet.Cells[1, 1];
			cell.Data = "=IF(A2=1,\"one\",IF(A2=2,\"two\",IF(A2=3,\"three\",\"four\")))";
			cell.Style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Right;

			worksheet.CellRanges[1, 1, 2, 1].Merge();
		}

		private void validationButton_Click(object sender, RoutedEventArgs e)
		{
			var form = new ValidationForm(workbook, workbookView.Selection.Range.Validation);
			if (form.ShowDialog() == true)
			{
				form.Apply(workbookView.Selection.Range.Validation);
			}
		}
	}
}
