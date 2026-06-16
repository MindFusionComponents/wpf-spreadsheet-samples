//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Media;

using Microsoft.Win32;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.ReadyWorkbooks
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			string path = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ReadyWorkbooks.xml");
			if (File.Exists(path))
				workbook.LoadFromXml(path);
		}
	}
}
