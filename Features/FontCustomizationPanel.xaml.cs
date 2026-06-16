//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using MindFusion.Spreadsheet.Wpf.StandardForms;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for FontCustomizationPanel.xaml
	/// </summary>
	public partial class FontCustomizationPanel : UserControl
	{
		public FontCustomizationPanel()
		{
			InitializeComponent();

			var sheet = workbook.Worksheets.Add();

			sheet.BeginInit();

			sheet.Rows[1].Style.FontBold = true;

			sheet.Cells[1, 1].Data = "Font";
			sheet.Cells[2, 1].Data = "Result";

			for (int i = 0; i < 9; i++)
			{
				var cell = sheet.Cells[2, i + 2];
				cell.Data = LOREM_IPSUM[i];
				GenerateRandomFont(ref cell);
				sheet.Cells[1, i + 2].Data = GetFontInfo(cell);
			}

			sheet.EndInit();

			workbookView.ResizeColumnsToFit(1, 2);
			workbookView.ResizeRowsToFit(2, 11);
		}

		private void GenerateRandomFont(ref Cell cell)
		{
			var style = cell.Style;

			style.FontName = fonts[rnd.Next(0, fonts.Count - 1)];
			style.FontBold = rnd.Next(1, 10) % 2 == 0;
			style.FontItalic = rnd.Next(1, 10) % 2 == 0;
			style.FontSize = rnd.Next(9, 14);
		}

		private string GetFontInfo(Cell cell)
		{
			var style = cell.Style;

			return string.Format("{0} {1}pt {2}", style.FontName,
				style.FontSize,
				!style.FontBold.Value && !style.FontItalic.Value ? "Regular" :
				style.FontItalic.Value && style.FontItalic.Value ? "Bold, Italic" :
				style.FontBold.Value ? "Bold" : "Italic");
		}

		private void fontButton_Click(object sender, RoutedEventArgs e)
		{
			FontPickerForm picker = new FontPickerForm(workbookView.Workbook, workbookView.ActiveCell.Style);
			if (picker.ShowDialog() == true)
			{
				var style = workbookView.Selection.Range.Style;
				style.FontName = picker.FontName;
				style.FontSize = picker.FontSize;
				style.FontBold = picker.Bold;
				style.FontItalic = picker.Italic;
				style.FontUnderline = picker.Underline;
				style.FontStrikeout = picker.Strikeout;
				style.TextBrush = picker.TextBrush;
			}
		}


		private static readonly string[] LOREM_IPSUM = new string[] 
		{
			"Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
			"Morbi ultrices sollicitudin lectus, sit amet tincidunt dui aliquet vel.",
			"Vestibulum accumsan quam volutpat pharetra porta.",
			"Donec in quam a mi eleifend tincidunt rutrum nec velit.",
			"Proin sit amet tortor felis.",
			"Morbi et nisi vel felis rutrum porta.",
			"Vivamus bibendum, orci eu vestibulum imperdiet, est mauris pulvinar libero, id tincidunt felis ligula nec massa.",
			"Ut a cursus leo. Duis molestie elit mi.",
			"Nullam volutpat dictum mi, vel aliquet augue eleifend ut.",
			"Donec rhoncus arcu vel nunc venenatis, eget semper neque placerat."
		};
		private static readonly Random rnd = new Random(DateTime.Now.Millisecond);
		private static List<string> fonts = Fonts.SystemFontFamilies.Select(f => f.ToString()).ToList();
	}
}
