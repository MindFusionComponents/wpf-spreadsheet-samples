//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for AutoFilterPanel.xaml
	/// </summary>
	public partial class AutoFilterPanel : UserControl
	{
		public AutoFilterPanel()
		{
			InitializeComponent();

			var document = XDocument.Parse(Properties.Resources.AutoFilter);
			workbook.LoadFromXml(document);
		}
	}
}
