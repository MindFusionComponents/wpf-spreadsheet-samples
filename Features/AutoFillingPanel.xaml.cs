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
	/// Interaction logic for AutoFillingPanel.xaml
	/// </summary>
	public partial class AutoFillingPanel : UserControl
	{
		public AutoFillingPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();

			sheet.BeginInit();
			sheet.Cells[0, 0].Data = 1;
			sheet.Cells[1, 0].Data = 2;
			sheet.Cells[0, 1].Data = 4;
			sheet.Cells[1, 1].Data = 6;
			sheet.EndInit();
			workbookView.Selection.Set(0, 0, 1, 1);
		}
	}
}
