//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Media;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.DatabaseFunctions
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			// Initialize the worksheet
			var worksheet = workbook.Worksheets.Add("Database Functions");
			worksheet.BeginInit();
			worksheet.Rows[0, worksheet.Rows.Count - 1].Height = "24pt";
			worksheet.Rows[0].Style.FontBold = true;
			worksheet.Columns[0].Width = "100pt";
			worksheet.Columns[1].Width = "100pt";
			worksheet.Columns[6].Width = "80pt";
			worksheet.Columns[7].Width = "200pt";
			worksheet.Columns[8].Width = "60pt";
			worksheet.Columns[4].Style.Format = "0.00";
			worksheet.Columns[8].Style.Format = "0.####";

			IStyle style;
			style = worksheet.CellRanges[0, 0, worksheet.Columns.Count - 1, worksheet.Rows.Count - 1].Style;
			style.FontName = "Segoe UI";
			style.FontSize = 10;
			style.PaddingLeft = "4px";
			style.PaddingTop = "4px";
			style.PaddingRight = "4px";
			style.PaddingBottom = "4px";
			style.HorizontalAlignment = MindFusion.Spreadsheet.Wpf.HorizontalAlignment.Left;
			style.VerticalAlignment = MindFusion.Spreadsheet.Wpf.VerticalAlignment.Middle;

			style = worksheet.CellRanges["E4:E10"].Style;
			style.BorderRightBrush = Brushes.SteelBlue;
			style.BorderRightSize = "1px";

			style = worksheet.CellRanges["A10:E10"].Style;
			style.BorderBottomBrush = Brushes.SteelBlue;
			style.BorderBottomSize = "1px";

			style = worksheet.CellRanges["A4:E4"].Style;
			style.BorderTopBrush = Brushes.SteelBlue;
			style.BorderTopSize = "1px";

			style = worksheet.CellRanges["A1:E1"].Style;
			style.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xF4, 0xFA, 0xFF));
			style.Background.Freeze();

			style = worksheet.CellRanges["A4:E4"].Style;
			style.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xF4, 0xFA, 0xFF));
			style.Background.Freeze();
			style.FontBold = true;

			style = worksheet.CellRanges["H2:H13"].Style;
			style.TextBrush = Brushes.DimGray;

			worksheet.Cells["A1"].Data = "NAME";
			worksheet.Cells["B1"].Data = "CATEGORY";
			worksheet.Cells["C1"].Data = "PRICE";
			worksheet.Cells["D1"].Data = "PRICE";
			worksheet.Cells["B2"].Data = "Beverages";
			worksheet.Cells["C2"].Data = ">10";
			worksheet.Cells["D2"].Data = "<15";
			worksheet.Cells["B3"].Data = "Confections";

			worksheet.Cells["A4"].Data = "Name";
			worksheet.Cells["B4"].Data = "Category";
			worksheet.Cells["C4"].Data = "Weight";
			worksheet.Cells["D4"].Data = "Qty";
			worksheet.Cells["E4"].Data = "Price";

			worksheet.Cells["A5"].Data = "Blueberry Tea";
			worksheet.Cells["B5"].Data = "Beverages";
			worksheet.Cells["C5"].Data = 400;
			worksheet.Cells["D5"].Data = 12;
			worksheet.Cells["E5"].Data = 12.5;

			worksheet.Cells["A6"].Data = "Ipoh Cofee";
			worksheet.Cells["B6"].Data = "Beverages";
			worksheet.Cells["C6"].Data = 500;
			worksheet.Cells["D6"].Data = 20;
			worksheet.Cells["E6"].Data = 18;

			worksheet.Cells["A7"].Data = "Chocolade";
			worksheet.Cells["B7"].Data = "Confections";
			worksheet.Cells["C7"].Data = 1250;
			worksheet.Cells["D7"].Data = 10;
			worksheet.Cells["E7"].Data = 22.45;

			worksheet.Cells["A8"].Data = "Steeleye Stout";
			worksheet.Cells["B8"].Data = "Beverages";
			worksheet.Cells["C8"].Data = 750;
			worksheet.Cells["D8"].Data = 6;
			worksheet.Cells["E8"].Data = 7.95;

			worksheet.Cells["A9"].Data = "Scottish Longbreads";
			worksheet.Cells["B9"].Data = "Confections";
			worksheet.Cells["C9"].Data = 450;
			worksheet.Cells["D9"].Data = 8;
			worksheet.Cells["E9"].Data = 11.5;

			worksheet.Cells["A10"].Data = "Mascarpone";
			worksheet.Cells["B10"].Data = "Dairy";
			worksheet.Cells["C10"].Data = 250;
			worksheet.Cells["D10"].Data = 2;
			worksheet.Cells["E10"].Data = 7.5;

			worksheet.Cells["G1"].Data = "Function";
			worksheet.Cells["H1"].Data = "Example";
			worksheet.Cells["I1"].Data = "Result";

			worksheet.Cells["G2"].Data = "DAVERAGE";
			worksheet.Cells["G3"].Data = "DCOUNT";
			worksheet.Cells["G4"].Data = "DCOUNTA";
			worksheet.Cells["G5"].Data = "DGET";
			worksheet.Cells["G6"].Data = "DMAX";
			worksheet.Cells["G7"].Data = "DMIN";
			worksheet.Cells["G8"].Data = "DPRODUCT";
			worksheet.Cells["G9"].Data = "DSTDEV";
			worksheet.Cells["G10"].Data = "DSTDEVP";
			worksheet.Cells["G11"].Data = "DSUM";
			worksheet.Cells["G12"].Data = "DVAR";
			worksheet.Cells["G13"].Data = "DVARP";

			worksheet.Cells["H2"].Data = @"'=DAVERAGE(A4:E10,""Qty"",B1:C2)";
			worksheet.Cells["H3"].Data = @"'=DCOUNT(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H4"].Data = @"'=DCOUNTA(A4:E10,""Name"",B1:B3)";
			worksheet.Cells["H5"].Data = @"'=DGET(A4:E10,""Price"",B1:D2)";
			worksheet.Cells["H6"].Data = @"'=DMAX(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H7"].Data = @"'=DMIN(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H8"].Data = @"'=DPRODUCT(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H9"].Data = @"'=DSTDEV(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H10"].Data = @"'=DSTDEVP(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H11"].Data = @"'=DSUM(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H12"].Data = @"'=DVAR(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["H13"].Data = @"'=DVARP(A4:E10,""Price"",B1:B2)";

			worksheet.Cells["I2"].Data = @"=DAVERAGE(A4:E10,""Qty"",B1:C2)";
			worksheet.Cells["I3"].Data = @"=DCOUNT(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I4"].Data = @"=DCOUNTA(A4:E10,""Name"",B1:B3)";
			worksheet.Cells["I5"].Data = @"=DGET(A4:E10,""Price"",B1:D2)";
			worksheet.Cells["I6"].Data = @"=DMAX(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I7"].Data = @"=DMIN(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I8"].Data = @"=DPRODUCT(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I9"].Data = @"=DSTDEV(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I10"].Data = @"=DSTDEVP(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I11"].Data = @"=DSUM(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I12"].Data = @"=DVAR(A4:E10,""Price"",B1:B2)";
			worksheet.Cells["I13"].Data = @"=DVARP(A4:E10,""Price"",B1:B2)";

			worksheet.Cells["I2"].SetAnnotation("Evaluates the average quantity of beverages with price over 10.");
			worksheet.Cells["I3"].SetAnnotation("Counts the number of beverages in the table.");
			worksheet.Cells["I4"].SetAnnotation("Counts the number of beverages and confections in the table.");
			worksheet.Cells["I5"].SetAnnotation("Returns the price of the beverage with price between 10 and 15.");
			worksheet.Cells["I6"].SetAnnotation("Returns the largest price of beverages in the table.");
			worksheet.Cells["I7"].SetAnnotation("Returns the smallest price of beverages in the table.");
			worksheet.Cells["I8"].SetAnnotation("Returns the product of all prices of beverages in the table.");
			worksheet.Cells["I9"].SetAnnotation("Returns the standard deviation of the beverage prices in the table.");
			worksheet.Cells["I10"].SetAnnotation("Returns the standard deviation of the beverage prices in the table.");
			worksheet.Cells["I11"].SetAnnotation("Returns the sum of all prices of beverages in the table.");
			worksheet.Cells["I12"].SetAnnotation("Returns the variance of the beverage prices in the table.");
			worksheet.Cells["I13"].SetAnnotation("Returns the variance of the beverage prices in the table.");

			worksheet.EndInit();
		}
	}
}
