//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Windows;
using System.Windows.Media;
using System.Globalization;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Charts
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			var chartsSheet = workbook.Worksheets.Add("Charts");
			var dataSheet = workbook.Worksheets.Add("Data");

			dataSheet.BeginInit();
			for (int c = 1; c <= 5; c++)
				dataSheet.Cells[c, 0].Data = string.Format("Column {0}", c);
			for (int r = 1; r <= 5; r++)
				dataSheet.Cells[0, r].Data = string.Format("Row {0}", r);
			for (int c = 1; c <= 5; c++)
				for (int r = 1; r <= 5; r++)
					dataSheet.Cells[c, r].Data = rnd.Next(0, 100);
			dataSheet.EndInit();

			chartsSheet.BeginInit();
			var cell = chartsSheet.Cells[0, 0];
			cell.Data = "Column chart";
			var chart = chartsSheet.Drawing.AddChart(0, 1);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.Column;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);
			chart.YAxisSettings.ShowMajorGridLines = true;
			chart.GridFillBrush = new SolidColorBrush(Color.FromArgb(255, 250, 250, 250));
			chart.GridLineBrush = Brushes.LightGray;

			cell = chartsSheet.Cells[9, 0];
			cell.Data = "Bar chart";
			chart = chartsSheet.Drawing.AddChart(9, 1);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.Bar;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);
			chart.XAxisSettings.ShowMajorGridLines = true;
			chart.GridFillBrush = new SolidColorBrush(Color.FromArgb(255, 250, 250, 250));
			chart.GridLineBrush = Brushes.LightGray;

			cell = chartsSheet.Cells[0, 19];
			cell.Data = "Line chart with marks";
			chart = chartsSheet.Drawing.AddChart(0, 20);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.LineWithMarks;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);
			chart.DefaultLineSize = MindFusion.Spreadsheet.Wpf.Measure.Point(2);
			chart.YAxisSettings.ShowMajorGridLines = true;
			chart.GridFillBrush = new SolidColorBrush(Color.FromArgb(255, 250, 250, 250));
			chart.GridLineBrush = Brushes.LightGray;

			cell = chartsSheet.Cells[9, 19];
			cell.Data = "Stacked area chart";
			chart = chartsSheet.Drawing.AddChart(9, 20);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.AreaStacked;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);
			chart.YAxisSettings.ShowMajorGridLines = true;
			chart.GridFillBrush = new SolidColorBrush(Color.FromArgb(255, 250, 250, 250));
			chart.GridLineBrush = Brushes.LightGray;

			cell = chartsSheet.Cells[0, 38];
			cell.Data = "Scatter chart with smooth marks";
			chart = chartsSheet.Drawing.AddChart(0, 39);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.ScatterWithSmoothLines;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);
			chart.DefaultLineSize = MindFusion.Spreadsheet.Wpf.Measure.Point(2);
			chart.YAxisSettings.ShowMajorGridLines = true;
			chart.GridFillBrush = new SolidColorBrush(Color.FromArgb(255, 250, 250, 250));
			chart.GridLineBrush = Brushes.LightGray;

			cell = chartsSheet.Cells[9, 38];
			cell.Data = "Pie chart";
			chart = chartsSheet.Drawing.AddChart(9, 39);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.Pie;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);

			cell = chartsSheet.Cells[0, 57];
			cell.Data = "Filled radar chart";
			chart = chartsSheet.Drawing.AddChart(0, 58);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.RadarFilled;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);

			cell = chartsSheet.Cells[9, 57];
			cell.Data = "Bubble chart";
			chart = chartsSheet.Drawing.AddChart(9, 58);
			chart.Type = MindFusion.Spreadsheet.Wpf.Charts.ChartType.Bubble;
			chart.SetDataSource("Data!A1:F6", MindFusion.Spreadsheet.Wpf.Charts.PlotBy.Column, true, true);

			foreach (var c in chartsSheet.Drawing.Charts)
			{
				c.HorizontalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(2);
				c.VerticalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(1);
				c.Width = MindFusion.Spreadsheet.Wpf.Measure.Point(400);
				c.Height = MindFusion.Spreadsheet.Wpf.Measure.Point(250);
				c.LineBrush = Brushes.Navy;
				c.LineSize = MindFusion.Spreadsheet.Wpf.Measure.Point(1);
			}
			chartsSheet.EndInit();
		}


		private static readonly Random rnd = new Random(DateTime.Now.Millisecond);
	}
}
