//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for UndoRedoPanel.xaml
	/// </summary>
	public partial class UndoRedoPanel : UserControl
	{
		public UndoRedoPanel()
		{
			InitializeComponent();

			workbook.Worksheets.Add();

			workbook.UndoEnabled = true;
			workbook.PropertyValueChanged += (s, e) =>
			{
				if (e.PropertyName == "CanUndo")
					undoButton.IsEnabled = (bool)e.NewValue;
				if (e.PropertyName == "CanRedo")
					redoButton.IsEnabled = (bool)e.NewValue;
			};
			undoButton.IsEnabled = false;
			redoButton.IsEnabled = false;

			string path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "UndoRedo.csv");
			if (File.Exists(path))
			{
				using (workbook.StartChangeOperation(workbook.Worksheets[0]))
					new CsvImporter() { Separator = '\t', Quote = '"', Encoding = Encoding.Unicode }.
						Import(path, workbook.Worksheets[0]);
			}
		}

		private void undoButton_Click(object sender, RoutedEventArgs e)
		{
			workbook.Undo();
		}

		private void redoButton_Click(object sender, RoutedEventArgs e)
		{
			workbook.Redo();
		}
	}
}
