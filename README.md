# WPF Spreadsheet Samples (`wpf-spreadsheet-samples`)

A comprehensive collection of C# WPF (Windows Presentation Foundation) sample applications demonstrating the features, formula engine, and data-management capabilities of the **MindFusion.Spreadsheet for WPF** control.

**Excel-Free Spreadsheet Operations:** MindFusion.Spreadsheet provides WPF applications with the ability to create, open, edit, calculate, and export spreadsheet documents natively **without requiring Microsoft Excel** to be installed on the machine.

All sample projects in this repository are pre-configured to reference the official **[MindFusion.Spreadsheet.Wpf](https://www.nuget.org/packages/MindFusion.Spreadsheet.Wpf/) NuGet package** directly rather than referencing local files. This enables automatic package restore and seamless out-of-the-box building.

---

## 🚀 Key Features

*   **Format Support:** Seamlessly load and save industry-standard formats including **Excel (XLSX), OpenDocument (ODS), Comma-Separated Values (CSV), and PDF**.
*   **Formula Engine:** Robust built-in calculation engine with hundreds of predefined functions (mathematical, statistical, string, database, logical, etc.) and complete support for writing **custom functions**.
*   **Grid Manipulation:** Advanced grid features such as cell merging, freezing panes, row/column resizing, hidden rows/columns, and multi-column sorting.
*   **Rich Formatting:** Dynamic conditional formatting, styles, borders, background patterns, custom fonts, and text alignment.
*   **Data Integration:** Ability to bind and display data tables from relational databases directly into workbooks as individual worksheets.

---

## 📂 Samples Demonstrated

This repository includes **10 specialized sample projects** illustrating various spreadsheet implementations:

*   📅 **Calendar** — Demonstrates a fully interactive, month-view calendar laid out entirely within a spreadsheet grid.
*   🎨 **Conditional Formats** — Highlights cell data dynamically with custom background styles, colors, and markers based on cell contents.
*   🧮 **Custom Functions** — Demonstrates how to register and evaluate your own custom math or domain-specific functions in the calculation engine.
*   🗄️ **Database** — Shows how to retrieve relational database tables and display them as separate, tabbed worksheets within a single workbook.
*   📊 **Database Functions** — Demonstrates the built-in database-oriented statistical and lookup functions available in the control.
*   🛠️ **Features** — An interactive playground illustrating basic grid manipulations: cell merging, freezing columns/panes, undo/redo stacks, and more.
*   🧾 **Invoice Template** — A clean, print-ready, interactive invoice that automatically calculates totals and can be printed or exported.
*   📅 **Project Planner** — Illustrates how to build a simplified project management planner and progress timeline on the spreadsheet grid.
*   📁 **Ready Workbooks** — Displays a collection of pre-made templates and complex spreadsheets showcasing layout and rendering quality.
*   🔼 **Sorting** — Demonstrates how to programmatically and interactively sort worksheet rows and multi-column datasets.

---

## ⚙️ Getting Started

### Prerequisites
*   **IDE:** Visual Studio 2022, 2026 or newer.
*   **Framework:** .NET Framework 8 SDK/Runtime.
*   **Package Manager:** NuGet (integrated natively into Visual Studio).

### How to Build & Run
1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/MindFusionComponents/wpf-spreadsheet-samples.git
    cd wpf-spreadsheet-samples
    ```
2.  **Open a Sample:**
    *   Navigate to any sample folder (e.g., `InvoiceTemplate`, `ConditionalFormats`, or `CustomFunctions`).
    *   Double-click the `.sln` or `.csproj` file to open it in Visual Studio.
3.  **Restore NuGet Packages:**
    *   When you build or debug the project, Visual Studio will automatically restore the missing `MindFusion.Spreadsheet.Wpf` package and its dependencies.
4.  **Run:**
    *   Press `F5` or click **Start** in Visual Studio to compile and run the sample!

---

## 📄 License and Product Info

*   These samples are designed to work with **MindFusion.Spreadsheet for WPF**.
*   A copy of MindFusion.Spreadsheet for WPF can be purchased with the control's **full C# source code**.
*   For product documentation, licensing, or evaluation licenses, visit the [MindFusion Official Website](https://mindfusion.dev) and the [official WPF Spreadsheet product page](https://mindfusion.dev/spreadsheet-wpf.html).
