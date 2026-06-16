//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;

using MindFusion.Spreadsheet.Wpf.Expressions;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Database
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			workbookView.AllowMoveCells = false;
			workbookView.AllowMoveHeaders = false;

			var dataSet = new nwindDataSet();
			var categoriesTableAdapter = new nwindDataSetTableAdapters.CategoriesTableAdapter();
			var customersTableAdapter = new nwindDataSetTableAdapters.CustomersTableAdapter();
			var employeesTableAdapter = new nwindDataSetTableAdapters.EmployeesTableAdapter();
			var ordersTableAdapter = new nwindDataSetTableAdapters.OrdersTableAdapter();
			var productsTableAdapter = new nwindDataSetTableAdapters.ProductsTableAdapter();
			var shippersTableAdapter = new nwindDataSetTableAdapters.ShippersTableAdapter();
			var suppliersTableAdapter = new nwindDataSetTableAdapters.SuppliersTableAdapter();

			loading = true;
			categoriesTableAdapter.Fill(dataSet.Categories);
			CreateWorksheetFromData("Categories", categoriesTableAdapter.GetData());

			customersTableAdapter.Fill(dataSet.Customers);
			CreateWorksheetFromData("Customers", customersTableAdapter.GetData());

			employeesTableAdapter.Fill(dataSet.Employees);
			CreateWorksheetFromData("Employees", employeesTableAdapter.GetData());

			ordersTableAdapter.Fill(dataSet.Orders);
			CreateWorksheetFromData("Orders", ordersTableAdapter.GetData());

			productsTableAdapter.Fill(dataSet.Products);
			CreateWorksheetFromData("Products", productsTableAdapter.GetData());

			shippersTableAdapter.Fill(dataSet.Shippers);
			CreateWorksheetFromData("Shippers", shippersTableAdapter.GetData());

			suppliersTableAdapter.Fill(dataSet.Suppliers);
			CreateWorksheetFromData("Suppliers", suppliersTableAdapter.GetData());
			loading = false;
		}

		private void workbookView_InplaceEditStarting(object sender, InplaceEditValidationEventArgs e)
		{
			// Auto-increment columns cannot be edited directly
			var sheet = workbookView.ActiveWorksheet;
			var cell = e.Item as Cell;
			if (cell == null)
				return;

			var column = sheet.Columns[cell.Column];
			if (column.Tag is bool && (bool)column.Tag)
				e.Cancel = true;
		}

		private void workbook_WorksheetCellChanging(object sender, CellValidationEventArgs e)
		{
			if (loading)
				return;

			// Auto-increment columns cannot be edited directly
			var sheet = e.Cell.Worksheet;
			var cell = e.Cell;
			var column = sheet.Columns[cell.Column];
			if (column.Tag is bool && (bool)column.Tag)
				e.Cancel = true;
		}

		private void workbook_WorksheetCellChanged(object sender, CellChangedEventArgs e)
		{
			if (loading)
				return;

			// If the bottommost row is edited, 'add' a new, empty row
			var worksheet = e.Cell.Worksheet;
			if (worksheet.Rows[e.Cell.Row + 1].IsHidden)
			{
				worksheet.Rows[e.Cell.Row + 1].IsHidden = false;
				worksheet.Rows[e.Cell.Row].Title = null;
				worksheet.Rows[e.Cell.Row + 1].Title = "";
			}

			// Also make sure to automatically update all auto-increment columns
			int columnCount = (int)worksheet.Tag;
			for (int i = 0; i < columnCount; i++)
			{
				var column = worksheet.Columns[i];
				if (column.Tag is bool && (bool)column.Tag)
				{
					var cell = worksheet.Cells[column.Index, e.Cell.Row];
					if (cell.Data == null)
					{
						if (e.Cell.Row > 0)
						{
							var previous = worksheet.Cells[column.Index, e.Cell.Row - 1];
							if (previous.Value is double)
								cell.Data = (double)previous.Value + 1;
							else
								cell.Data = 1;
						}
						else
						{
							cell.Data = 1;
						}
					}
				}
			}
		}

		private Worksheet CreateWorksheetFromData(string name, DataTable data)
		{
			var sheet = workbook.Worksheets.Add(name);

			sheet.BeginInit();
			foreach (DataColumn column in data.Columns)
			{
				var sheetColumn = sheet.Columns[column.Ordinal];
				sheetColumn.Title = column.Caption;
				if (column.AutoIncrement)
				{
					sheetColumn.Tag = true;
					sheetColumn.Style.Background = new SolidColorBrush(Color.FromArgb(255, 251, 246, 221));
					sheetColumn.Style.Background.Freeze();
				}
			}
			sheet.Tag = data.Columns.Count;
			sheet.Columns[data.Columns.Count, sheet.Columns.Count - 1].IsHidden = true;
			int i = 0;
			foreach (DataRow row in data.Rows)
			{
				foreach (DataColumn column in data.Columns)
					sheet.Cells[column.Ordinal, i].Data = row[column];
				i++;
			}
			sheet.Rows[i + 1, sheet.Rows.Count - 1].IsHidden = true;
			sheet.Rows[i].Title = "";

			workbookView.ActiveWorksheet = sheet;
			workbookView.ResizeColumnsToFit(0, data.Columns.Count - 1);
			Measure defaultColumnWidth = "48pt";
			for (int c = 0; c < data.Columns.Count; c++)
			{
				if (sheet.Columns[c].Width.Amount < defaultColumnWidth.Amount)
					sheet.Columns[c].Width = defaultColumnWidth;
			}
			sheet.EndInit();

			return sheet;
		}


		private bool loading;
	}
}
