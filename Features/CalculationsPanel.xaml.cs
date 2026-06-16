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
	/// Interaction logic for CalculationsPanel.xaml
	/// </summary>
	public partial class CalculationsPanel : UserControl
	{
		public CalculationsPanel()
		{
			InitializeComponent();

			string path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Calculations.xml");
			if (File.Exists(path))
				workbook.LoadFromXml(path);
		}
	}
}
