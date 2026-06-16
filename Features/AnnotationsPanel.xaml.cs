//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for AnnotationsPanel.xaml
	/// </summary>
	public partial class AnnotationsPanel : UserControl
	{
		public AnnotationsPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();
			sheet.BeginInit();

			var cell = sheet.Cells[3, 3];
			cell.Data = "Styles";
			cell.SetAnnotation("Annotation with applied styles.");
			var annotation = cell.Annotation.CellAnnotation;
			annotation.ShowAlways = true;
			annotation.AnchorPosition = ObjectAnchorPosition.TopLeft;
			annotation.HorizontalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(-75);
			annotation.VerticalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(-30);
			annotation.Height = MindFusion.Spreadsheet.Wpf.Measure.Point(50);
			annotation.BackBrush = Brushes.Navy;
			annotation.TextBrush = Brushes.Gold;
			annotation.LineBrush = Brushes.Gold;
			annotation.Margin = MindFusion.Spreadsheet.Wpf.Measure.Point(5);
			annotation.FontBold = true;

			cell = sheet.Cells[4, 3];
			cell.Data = "Default";
			cell.SetAnnotation("Annotation without customizations.");
			annotation = cell.Annotation.CellAnnotation;

			cell = sheet.Cells[4, 4];
			cell.Data = "Fading and animations";
			cell.SetAnnotation("This annotation uses animation and fading.");
			annotation = cell.Annotation.CellAnnotation;
			annotation.AnchorPosition = ObjectAnchorPosition.BottomRight;
			annotation.VerticalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(5);
			annotation.IndicatorBrush = Brushes.DarkGreen;

			cell = sheet.Cells[3, 4];
			cell.Data = "More styles";
			cell.SetAnnotation("C:\\> echo Hello, world!\r\nHello, world!\r\nC:\\>\u005F");
			annotation = cell.Annotation.CellAnnotation;
			annotation.ShowAlways = true;
			annotation.AnchorPosition = ObjectAnchorPosition.BottomLeft;
			annotation.HorizontalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(-140);
			annotation.VerticalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(5);
			annotation.Width = MindFusion.Spreadsheet.Wpf.Measure.Point(135);
			annotation.Height = MindFusion.Spreadsheet.Wpf.Measure.Point(40);
			annotation.LineStyle = DashStyles.Dash;
			annotation.LineSize = MindFusion.Spreadsheet.Wpf.Measure.Point(1.5);
			annotation.LineBrush = Brushes.Chartreuse;
			annotation.FontName = "Lucida Console";
			annotation.TextBrush = Brushes.Chartreuse;
			annotation.BackBrush = Brushes.Black;
			annotation.FontSize = 9;
			annotation.FontBold = true;
			workbookView.Selection.AddObject(annotation);

			sheet.CellRanges[4, 4, 5, 4].Merge();

			sheet.EndInit();

			workbookView.ActiveCellChanged += workbookView_ActiveCellChanged;
		}

		void workbookView_ActiveCellChanged(object sender, EventArgs e)
		{
			checkBox.IsEnabled = workbookView.ActiveCell.Annotation != null;
			if (checkBox.IsEnabled)
				checkBox.IsChecked = workbookView.ActiveCell.Annotation.ShowAlways;
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			var cell = workbookView.ActiveCell;
			if (cell == null)
				return;

			if (cell.Annotation == null)
			{
				// Add a new annotation
				cell.SetAnnotation("New comment.");
				cell.Annotation.ShowAlways = true;
				checkBox.IsEnabled = true;
				checkBox.IsChecked = true;
				workbookView.BeginEdit(cell.Annotation);
			}
			else
			{
				// Remove the annotation
				cell.RemoveAnnotation();
			}
		}

		private void CheckBox_Checked(object sender, RoutedEventArgs e)
		{
			var cell = workbookView.ActiveCell;
			if (cell == null || cell.Annotation == null)
				return;

			cell.Annotation.ShowAlways = checkBox.IsChecked == true;
		}

		private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
		{
			var cell = workbookView.ActiveCell;
			if (cell == null || cell.Annotation == null)
				return;

			cell.Annotation.ShowAlways = checkBox.IsChecked == true;
		}
	}
}
