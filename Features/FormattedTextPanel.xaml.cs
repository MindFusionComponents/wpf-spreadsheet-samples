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
	/// Interaction logic for FormattedTextPanel.xaml
	/// </summary>
	public partial class FormattedTextPanel : UserControl
	{
		public FormattedTextPanel()
		{
			InitializeComponent();

			var worksheet = workbook.Worksheets.Add();

			worksheet.Cells[0, 0].Data = new FormattedText(
				@"The formatted texts can contain <b>bold</b>, <i>italic</i>, <u>underlined</u> and " +
				@"<s>strikeout</s> texts as well as <color=""#ff0000"">text</color> <color=""#00ff00"">in</color> " +
				@"<color=""#0000ff"">various</color> <color=""#ff8000"">colors</color>, " +
				@"<fontname=""Courier New"">fonts</fontname> and <fontsize=""16"">sizes</fontsize>." + Environment.NewLine +
				@"<color=""#8000FF""><s>Or</s> <fontname=""Times New Roman""><fontsize=""14"">any</fontsize> <b>comb<i>inat</i>ion</b></color> <fontsize=""18"">of</fontname> <color=""#008080"">the</color></fontsize> a<u>bov</u>e.");

			workbookView.ResizeColumnToFit(0);
			workbookView.ResizeRowToFit(0);
		}
	}
}
