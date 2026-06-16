//
// Copyright (c) 2021, MindFusion LLC - Bulgaria.
//

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using Microsoft.Win32;


namespace MindFusion.Spreadsheet.Wpf.Samples.CS.Features
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			listView.Items.Add("Merging");
			listView.Items.Add("Freezing");
			listView.Items.Add("Grouping and Outlining");
			listView.Items.Add("Undo and redo");
			listView.Items.Add("Copy and paste");
			listView.Items.Add("Auto-filling");
			listView.Items.Add("Auto filter");
			listView.Items.Add("Annotations");
			listView.Items.Add("Hyperlinks");
			listView.Items.Add("Font customization");
			listView.Items.Add("Formatted text");
			listView.Items.Add("Localization");
			listView.Items.Add("Data validation");
			listView.Items.Add("Number formats");
			listView.Items.Add("Styles");
			listView.Items.Add("Events");
			listView.Items.Add("Printing");
			listView.Items.Add("Calculations");
		}

		private void listViewItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			var item = sender as ListViewItem;
			if (item == null)
				return;

			UserControl activePanel = null;

			contentPanel.Children.Clear();

			var feature = item.Content.ToString();
			switch (feature)
			{
				case "Annotations":
					if (annotationsPanel == null)
						annotationsPanel = new AnnotationsPanel();
					activePanel = annotationsPanel;
					break;

				case "Auto-filling":
					if (autoFillingPanel == null)
						autoFillingPanel = new AutoFillingPanel();
					activePanel = autoFillingPanel;
					break;

				case "Auto filter":
					if (autoFilterPanel == null)
						autoFilterPanel = new AutoFilterPanel();
					activePanel = autoFilterPanel;
					break;

				case "Calculations":
					if (calculationsPanel == null)
						calculationsPanel = new CalculationsPanel();
					activePanel = calculationsPanel;
					break;

				case "Copy and paste":
					if (copyPastePanel == null)
						copyPastePanel = new CopyPastePanel();
					activePanel = copyPastePanel;
					break;

				case "Data validation":
					if (dataValidationPanel == null)
						dataValidationPanel = new DataValidationPanel();
					activePanel = dataValidationPanel;
					break;

				case "Events":
					if (eventsPanel == null)
						eventsPanel = new EventsPanel();
					activePanel = eventsPanel;
					break;

				case "Font customization":
					if (fontCustomizationPanel == null)
						fontCustomizationPanel = new FontCustomizationPanel();
					activePanel = fontCustomizationPanel;
					break;

				case "Formatted text":
					if (formattedTextPanel == null)
						formattedTextPanel = new FormattedTextPanel();
					activePanel = formattedTextPanel;
					break;

				case "Freezing":
					if (freezingPanel == null)
						freezingPanel = new FreezingPanel();
					activePanel = freezingPanel;
					break;

				case "Grouping and Outlining":
					if (outlinePanel == null)
						outlinePanel = new OutlinePanel();
					activePanel = outlinePanel;
					break;

				case "Hyperlinks":
					if (hyperlinksPanel == null)
						hyperlinksPanel = new HyperlinksPanel();
					activePanel = hyperlinksPanel;
					break;

				case "Localization":
					if (localizationPanel == null)
						localizationPanel = new LocalizationPanel();
					activePanel = localizationPanel;
					break;

				case "Merging":
					if (mergingPanel == null)
						mergingPanel = new MergingPanel();
					activePanel = mergingPanel;
					break;

				case "Number formats":
					if (numberFormatsPanel == null)
						numberFormatsPanel = new NumberFormatsPanel();
					activePanel = numberFormatsPanel;
					break;

				case "Printing":
					if (printingPanel == null)
						printingPanel = new PrintingPanel();
					activePanel = printingPanel;
					break;

				case "Styles":
					if (stylesPanel == null)
						stylesPanel = new StylesPanel();
					activePanel = stylesPanel;
					break;

				case "Undo and redo":
					if (undoRedoPanel == null)
						undoRedoPanel = new UndoRedoPanel();
					activePanel = undoRedoPanel;
					break;
			}

			if (activePanel != null)
				contentPanel.Children.Add(activePanel);
		}


		private AnnotationsPanel annotationsPanel;
		private AutoFillingPanel autoFillingPanel;
		private AutoFilterPanel autoFilterPanel;
		private CalculationsPanel calculationsPanel;
		private CopyPastePanel copyPastePanel;
		private DataValidationPanel dataValidationPanel;
		private EventsPanel eventsPanel;
		private FontCustomizationPanel fontCustomizationPanel;
		private FormattedTextPanel formattedTextPanel;
		private FreezingPanel freezingPanel;
		private OutlinePanel outlinePanel;
		private HyperlinksPanel hyperlinksPanel;
		private LocalizationPanel localizationPanel;
		private MergingPanel mergingPanel;
		private NumberFormatsPanel numberFormatsPanel;
		private PrintingPanel printingPanel;
		private StylesPanel stylesPanel;
		private UndoRedoPanel undoRedoPanel;
	}
}
