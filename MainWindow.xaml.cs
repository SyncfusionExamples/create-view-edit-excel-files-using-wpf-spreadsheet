#region Copyright Syncfusion Inc. 2001 - 2015
// Copyright Syncfusion Inc. 2001 - 2015. All rights reserved.
// Use of this code is subject to the terms of our license.
// A copy of the current license can be obtained at any time by e-mailing
// licensing@syncfusion.com. Any infringement will be prosecuted under
// applicable laws. 
#endregion
using Syncfusion.Windows.Tools.Controls;
using System.Windows;
using System.IO;
using System;


namespace SpreadsheetDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            this.spreadsheetControl.Loaded += OnSpreadsheetControlLoaded;
            this.spreadsheetControl.WorksheetAdded += OnSpreadsheetControlWorksheetAdded;
        }

        private void OnSpreadsheetControlWorksheetAdded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorksheetAddedEventArgs args)
        {
            /// Editing a specific cell value.
            var range = this.spreadsheetControl.ActiveSheet.Range[2, 2];
            this.spreadsheetControl.ActiveGrid.SetCellValue(range, "Syncfusion");
            this.spreadsheetControl.ActiveGrid.InvalidateCell(2, 2);
        }

        private void OnSpreadsheetControlLoaded(object sender, RoutedEventArgs e)
        {
            /// View or read the existing the Excel file.
#if NETCORE
            using (var fileStream = new FileStream(@"..\..\..\Data\GettingStarted.xlsx", FileMode.Open))
#else
            using (var fileStream = new FileStream(@"..\..\Data\GettingStarted.xlsx", FileMode.Open))
#endif
            {
                this.spreadsheetControl.Open(fileStream);
            }
        }

        /// <summary>
        /// Provide support for Excel like closing operation when press the close button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RibbonWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.spreadsheetControl.Loaded -= OnSpreadsheetControlLoaded;
            this.spreadsheetControl.WorksheetAdded -= OnSpreadsheetControlWorksheetAdded;
            this.spreadsheetControl.Commands.FileClose.Execute(null);
            if (Application.Current.ShutdownMode != ShutdownMode.OnExplicitShutdown)
                e.Cancel = true;
        }
    }
}
