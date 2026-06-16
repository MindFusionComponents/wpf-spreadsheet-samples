//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for EventsPanel.xaml
	/// </summary>
	public partial class EventsPanel : UserControl
	{
		public EventsPanel()
		{
			InitializeComponent();

			var worksheet = workbook.Worksheets.Add();

			workbookView.ActiveCellChanged += new EventHandler(workbookView_ActiveCellChanged);
			workbookView.ActiveWorksheetChanged += new EventHandler(workbookView_ActiveWorksheetChanged);
			workbookView.CellClicked += new EventHandler<CellMouseEventArgs>(workbookView_CellClicked);
			workbookView.CellSelectionChanged += new EventHandler(workbookView_CellSelectionChanged);
			workbookView.ColumnClicked += new EventHandler<ColumnMouseEventArgs>(workbookView_ColumnClicked);
			workbookView.HyperlinkClicked += new EventHandler<CellEventArgs>(workbookView_HyperlinkClicked);
			workbookView.InplaceEditEnded += new EventHandler<InplaceEditEventArgs>(workbookView_InplaceEditEnded);
			workbookView.InplaceEditEnding += new EventHandler<InplaceEditValidationEventArgs>(workbookView_InplaceEditEnding);
			workbookView.InplaceEditStarted += new EventHandler<InplaceEditEventArgs>(workbookView_InplaceEditStarted);
			workbookView.InplaceEditStarting += new EventHandler<InplaceEditValidationEventArgs>(workbookView_InplaceEditStarting);
			workbookView.ObjectClicked += new EventHandler<ObjectMouseEventArgs>(workbookView_ObjectClicked);
			workbookView.ObjectDeleted += new EventHandler<ObjectEventArgs>(workbookView_ObjectDeleted);
			workbookView.ObjectDeleting += new EventHandler<ObjectValidationEventArgs>(workbookView_ObjectDeleting);
			workbookView.ObjectDeselected += new EventHandler<ObjectEventArgs>(workbookView_ObjectDeselected);
			workbookView.ObjectModified += new EventHandler<ObjectInteractionEventArgs>(workbookView_ObjectModified);
			workbookView.ObjectModifying += new EventHandler<ObjectInteractionValidationEventArgs>(workbookView_ObjectModifying);
			workbookView.ObjectSelected += new EventHandler<ObjectEventArgs>(workbookView_ObjectSelected);
			workbookView.ObjectSelecting += new EventHandler<ObjectInteractionValidationEventArgs>(workbookView_ObjectSelecting);
			workbookView.ObjectStartModifying += new EventHandler<ObjectInteractionValidationEventArgs>(workbookView_ObjectStartModifying);
			workbookView.RowClicked += new EventHandler<RowMouseEventArgs>(workbookView_RowClicked);
			workbookView.WorksheetAdded += new EventHandler<WorksheetEventArgs>(workbookView_WorksheetAdded);
			workbookView.WorksheetTabClicked += new EventHandler<WorksheetTabMouseEventArgs>(workbookView_WorksheetTabClicked);

			var cell = worksheet.Cells["A1"];
			cell.Data = "Click Me!";
			cell.SetHyperlink(HyperlinkType.WebAddress, "", "");

			Picture pic = worksheet.Drawing.AddPicture();
			pic.Source = new BitmapImage(new Uri("/Features;component/App.ico", UriKind.RelativeOrAbsolute));
			pic.Anchor = ObjectAnchor.ToCell;
			pic.FromColumn = 1;
			pic.FromRow = 0;
			pic.HorizontalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(10);
			pic.VerticalOffset = MindFusion.Spreadsheet.Wpf.Measure.Point(10);
			pic.Width = MindFusion.Spreadsheet.Wpf.Measure.Point(100);
			pic.Height = MindFusion.Spreadsheet.Wpf.Measure.Point(100);

			workbook.ActionRedone += new EventHandler<UndoEventArgs>(workbook_ActionRedone);
			workbook.ActionUndone += new EventHandler<UndoEventArgs>(workbook_ActionUndone);
			workbook.WorksheetCellChanged += new EventHandler<CellChangedEventArgs>(workbook_WorksheetCellChanged);
			workbook.WorksheetCellChanging += new EventHandler<CellValidationEventArgs>(workbook_WorksheetCellChanging);
			workbook.WorksheetCellsCleared += new EventHandler<WorksheetEventArgs>(workbook_WorksheetCellsCleared);
			workbook.WorksheetChanged += new EventHandler<WorksheetChangedEventArgs>(workbook_WorksheetChanged);

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
		}

		void workbook_WorksheetChanged(object sender, WorksheetChangedEventArgs e)
		{
			OnEventHandled("WorksheetChanged");
		}

		void workbook_WorksheetCellsCleared(object sender, WorksheetEventArgs e)
		{
			OnEventHandled("WorksheetCellsCleared");
		}

		void workbook_WorksheetCellChanging(object sender, CellValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("WorksheetCellChanging");
		}

		void workbook_WorksheetCellChanged(object sender, CellChangedEventArgs e)
		{
			OnEventHandled("WorksheetCellChanged");
		}

		void workbook_ActionUndone(object sender, UndoEventArgs e)
		{
			OnEventHandled("ActionUndone");
		}

		void workbook_ActionRedone(object sender, UndoEventArgs e)
		{
			OnEventHandled("ActionRedone");
		}

		void workbookView_WorksheetTabClicked(object sender, WorksheetTabMouseEventArgs e)
		{
			OnEventHandled("WorksheetTabClicked");
		}

		void workbookView_WorksheetAdded(object sender, WorksheetEventArgs e)
		{
			OnEventHandled("WorksheetAdded");
		}

		void workbookView_CellSelectionChanged(object sender, EventArgs e)
		{
			OnEventHandled("CellSelectionChanged");
		}

		void workbookView_RowClicked(object sender, RowMouseEventArgs e)
		{
			OnEventHandled("RowClicked");
		}

		void workbookView_ObjectStartModifying(object sender, ObjectInteractionValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("ObjectStartModifying");
		}

		void workbookView_ObjectSelecting(object sender, ObjectInteractionValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("ObjectSelecting");
		}

		void workbookView_ObjectSelected(object sender, ObjectEventArgs e)
		{
			OnEventHandled("ObjectSelected");
		}

		void workbookView_ObjectModifying(object sender, ObjectInteractionValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("ObjectModifying");
		}

		void workbookView_ObjectModified(object sender, ObjectInteractionEventArgs e)
		{
			OnEventHandled("ObjectModified");
		}

		void workbookView_ObjectDeselected(object sender, ObjectEventArgs e)
		{
			OnEventHandled("ObjectDeselected");
		}

		void workbookView_ObjectDeleting(object sender, ObjectValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("ObjectDeleting");
		}

		void workbookView_ObjectDeleted(object sender, ObjectEventArgs e)
		{
			OnEventHandled("ObjectDeleted");
		}

		void workbookView_ObjectClicked(object sender, ObjectMouseEventArgs e)
		{
			OnEventHandled("ObjectClicked");
		}

		void workbookView_InplaceEditStarting(object sender, InplaceEditValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("InplaceEditStarting");
		}

		void workbookView_InplaceEditStarted(object sender, InplaceEditEventArgs e)
		{
			OnEventHandled("InplaceEditStarted");
		}

		void workbookView_InplaceEditEnding(object sender, InplaceEditValidationEventArgs e)
		{
			if (!checkValidation)
				return;
			OnEventHandled("InplaceEditEnding");
		}

		void workbookView_InplaceEditEnded(object sender, InplaceEditEventArgs e)
		{
			OnEventHandled("InplaceEditEnded");
		}

		void workbookView_HyperlinkClicked(object sender, CellEventArgs e)
		{
			OnEventHandled("HyperlinkClicked");
		}

		void workbookView_ColumnClicked(object sender, ColumnMouseEventArgs e)
		{
			OnEventHandled("ColumnClicked");
		}

		void workbookView_CellClicked(object sender, CellMouseEventArgs e)
		{
			OnEventHandled("CellClicked");
		}

		void workbookView_ActiveWorksheetChanged(object sender, EventArgs e)
		{
			OnEventHandled("ActiveWorksheetChanged");
		}

		void workbookView_ActiveCellChanged(object sender, EventArgs e)
		{
			OnEventHandled("ActiveCellChanged");
		}

		private void OnEventHandled(string eventName)
		{
			events.Items.Add(eventName);
			events.SelectedIndex = events.Items.Count - 1;
		}

		private void undoButton_Click(object sender, RoutedEventArgs e)
		{
			workbook.Undo();
		}

		private void redoButton_Click(object sender, RoutedEventArgs e)
		{
			workbook.Redo();
		}

		private void validation_Checked(object sender, RoutedEventArgs e)
		{
			checkValidation = true;
		}

		private void validation_Unchecked(object sender, RoutedEventArgs e)
		{
			checkValidation = false;
		}


		private bool checkValidation = true;
	}
}
