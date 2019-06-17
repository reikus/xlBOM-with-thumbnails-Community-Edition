using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Prism.Commands;
using Prism.Mvvm;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace ThumbnailedBOM.ViewModels
{
    public class MainViewModel : BindableBase
    {
        Window window;
        CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        CancellationToken token = default(CancellationToken);
        private string message = "Set the save location to start...";
        public string Message
        {
           
            get { return message; }
            set { SetProperty(ref message, value); }
        }
        private string saveLocation;
        public string SaveLocation
        {
            get { return saveLocation; }
            set { SetProperty(ref saveLocation, value); }
        }

        

        private bool isIdle = true;
        public bool IsIdle
        {
            get { return isIdle; }
            set { SetProperty(ref isIdle, value); }
        }

        private DelegateCommand donate;
        public DelegateCommand Donate =>
            donate ?? (donate = new DelegateCommand(ExecuteDonate, CanExecuteDonate));

       
        private DelegateCommand cancel;
        public DelegateCommand Cancel =>
            cancel ?? (cancel = new DelegateCommand(ExecuteCancel, CanExecuteCancel));
    
        private DelegateCommand start;
        public DelegateCommand Start =>
            start ?? (start = new DelegateCommand(ExecuteStart, CanExecuteStart));
       
        private DelegateCommand setSaveLocation;
        public DelegateCommand SetSaveLocation =>
            setSaveLocation ?? (setSaveLocation = new DelegateCommand(ExecuteSetSaveLocation, CanExecuteSetSaveLocation));

        
        public MainViewModel()
        {
            window = Application.Current.MainWindow;
            Start.ObservesProperty(() => this.SaveLocation);
            Start.ObservesProperty(() => this.IsIdle);
            SetSaveLocation.ObservesProperty(() => this.IsIdle);
            Donate.ObservesProperty(() => this.IsIdle);
            Cancel.ObservesProperty(() => this.IsIdle);
        }

        #region Execute and CanExecute
        void ExecuteSetSaveLocation()
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel files | *.xlsx";
            if (save.ShowDialog() == DialogResult.OK)
            {
                var ret = save.FileName;
                SaveLocation = ret;
            }
        }

        bool CanExecuteSetSaveLocation()
        {
            return IsIdle;
        }
        void ExecuteDonate()
        {
            Process.Start("https://www.paypal.me/AmenAllahJLILI");
        }

        bool CanExecuteDonate()
        {
            return IsIdle;
        }
        void ExecuteCancel()
        {
            cancellationTokenSource = new CancellationTokenSource();
            this.Message = "Cancel request received. Please wait..."; 
            
        }

        bool CanExecuteCancel()
        {
            return  !IsIdle;
        }
        async void ExecuteStart()
        {
            IsIdle = false;
            token = cancellationTokenSource.Token;
            var modelDoc = AddInContext.SOLIDWORKS.ActiveDoc as ModelDoc2; 
            if (modelDoc != null)
            {
                if (modelDoc.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                {
                    var selectionManager = modelDoc.SelectionManager as SelectionMgr;
                    int count = selectionManager.GetSelectedObjectCount2(-1);
                    if (count > 0)
                    {
                        bool found = false;
                        // Note: will traverse all selected tables and process last.
                        // needs to be extended to only process one table or all tables.
                        for (int i = 1; i < count+1; i++)
                        {
                            
                            Debug.Print(selectionManager.GetSelectedObjectType3(i, -1).ToString());

                            if (selectionManager.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelANNOTATIONTABLES)
                            {
                                found = true; 
                                var tableAnnotation = selectionManager.GetSelectedObject6(i, -1) as TableAnnotation;
                                BomTableAnnotation bomTableAnnotation;
                                bomTableAnnotation = tableAnnotation as BomTableAnnotation;

                                TableBoundryCondition tableBoundryConditions = new TableBoundryCondition();

                                if (bomTableAnnotation != null)
                                {
                                    swTableHeaderPosition_e tableHeaderPosition = (swTableHeaderPosition_e)tableAnnotation.GetHeaderStyle();


                                    tableBoundryConditions.RowHeaderIndex = 0;
                                    tableBoundryConditions.StartIndex = 1;
                                    tableBoundryConditions.EndIndex = tableAnnotation.RowCount - 1;
                                    tableBoundryConditions.HeaderPosition = swTableHeaderPosition_e.swTableHeader_Top;

                                    if (tableHeaderPosition == swTableHeaderPosition_e.swTableHeader_Bottom)
                                    {
                                        tableBoundryConditions.RowHeaderIndex = tableAnnotation.RowCount - 1;
                                        tableBoundryConditions.StartIndex = 0;
                                        tableBoundryConditions.EndIndex = tableAnnotation.RowCount - 2;
                                        tableBoundryConditions.HeaderPosition = swTableHeaderPosition_e.swTableHeader_Bottom;
                                    }

                                    var processRet = await ProcessTableAsync(bomTableAnnotation, tableAnnotation, tableBoundryConditions, token);

                                    if (processRet.Item1)
                                    {
                                        //var dialogRet = System.Windows.Forms.MessageBox.Show("Export completed successfully? Would you like to open the excel spreadsheet?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                        //if (dialogRet == DialogResult.Yes)
                                        //{
                                            Process.Start(saveLocation);
                                       // }
                                    }
                                    else
                                    {
                                        // case of an error 
                                        this.Message = processRet.Item2; 
                                    }

                                }
                                else
                                {
                                    Message = "Table is not of type Bill of Materials.";
                                }
                            }
                        }
                        if (found == false)
                            Message = "No table was selected.";
                    }
                    else
                    {
                        Message = "No table was selected.";
                    }
                }
            }
            Message = "Completed.";
            IsIdle = true;
        }
       
        Tuple<bool,string> ProcessTable(BomTableAnnotation bomTable, TableAnnotation table, TableBoundryCondition tableCondition, CancellationToken token)
        {
            try
            {
                using (var p = new ExcelPackage(new FileInfo(SaveLocation)))
                {
                    // create temporary folder to store images
                    var path = Path.Combine(Path.GetTempPath(),"thumbnailedTempFolder");
                    var tempDirectory = new DirectoryInfo(path);
                    if (tempDirectory.Exists == false)
                        tempDirectory.Create();

                    int Height = 30;
                    int Width = 30;

                    //Get the Worksheet created in the previous codesample. 
                    var ws = p.Workbook.Worksheets.Add("BOM");

                    for (int i = tableCondition.StartIndex; i <= tableCondition.EndIndex; i++)
                    {
                        if (token.IsCancellationRequested)
                        {
                            p.Save();
                            return new Tuple<bool, string>(false, "Cancelled by user.");
                        }

                        if (table.RowHidden[i])
                            continue;

                        string partNumber = string.Empty;
                        string itemNumber = string.Empty;
                        if (bomTable.GetComponentsCount2(i, string.Empty, out itemNumber, out partNumber)> 0)
                        {

                            var components = (object[])bomTable.GetComponents2(i,string.Empty);
                            var swComponent = components.First() as Component2;
                            var modelDoc = swComponent.GetModelDoc2() as ModelDoc2;
                            if (modelDoc != null)
                            {
                                
                                var modelDocTitle = Path.GetFileNameWithoutExtension(modelDoc.GetTitle()); 
                                SendMessageToUI($"{i}/{tableCondition.EndIndex} - Attempting to process {modelDocTitle}...");
                                var referencedConfiguration = swComponent.ReferencedConfiguration;
                                var configuration = modelDoc.GetActiveConfiguration() as Configuration;
                                if (configuration != null)
                                {
                                    string configurationName = configuration.Name;
                                    if (configurationName != referencedConfiguration)
                                        modelDoc.ShowConfiguration2(referencedConfiguration);
                                }
                                int er = 0; int wr = 0;
                                modelDoc.ViewZoomtofit2();
                                modelDoc.Visible = true;
                                var rowThumbnailFilePath = Path.Combine(tempDirectory.FullName, "thumbnail.bmp");
                                var saveRet = modelDoc.Extension.SaveAs(rowThumbnailFilePath, 0, 0, null, er, wr);
                                if (saveRet)
                                {
                                    
                                    Image img = Image.FromFile(rowThumbnailFilePath);
                                    ExcelPicture pic = ws.Drawings.AddPicture(i.ToString(), img);
                                    pic.SetPosition(i+1, Width, 1, Height);
                                    pic.SetSize(Height, Width);
                                }
                                else
                                {
                                    ws.Row(i + 1).Height = Height;
                                    ws.Cells[i+1, 1].Value = "N/A";
                                }

                                for (int j = 0; j < table.ColumnCount - 1; j++)
                                {
                                    if (table.ColumnHidden[j])
                                        continue;

                                    ws.Cells[i + 1, j + 2].Value = table.DisplayedText[i, j];
                                }

                                modelDoc.Visible = false; 
                            }
                            else
                            {
                                ws.Row(i + 1).Height = Height;
                                ws.Cells[i, 1].Value = "N/A";
                            }
                        }
                        else
                        {
                            ws.Row(i+1).Height = Height;
                            ws.Cells[i + 1, 1].Value = "N/A";
                        }
                       
                    }

                    // add row headers
                    for (int k = 0; k < table.ColumnCount -1; k++)
                    {
                        if (table.ColumnHidden[k])
                            continue;
                        ws.Cells[tableCondition.RowHeaderIndex+1, k + 2].Value = table.DisplayedText[tableCondition.RowHeaderIndex, k];
                        ws.Cells[tableCondition.RowHeaderIndex + 1, k + 2].Style.Font.Bold = true;
                    }
                    //Save and close the package.
                    p.Save();
                }
              
                
            }
            catch (Exception e)
            {

                return new Tuple<bool, string>(false, $"Fatal error: {e.Message} / {e.StackTrace}");
            }
          

            return new Tuple<bool, string>(true, "No error.");
        }

        /// <summary>
        /// Pumps message to the UI thread from another thread.
        /// </summary>
        /// <param name="message">Message.</param>
        void SendMessageToUI(string message)
        {
          
            window.Dispatcher.Invoke(() => {
                this.Message = message;
            });
        }

        Task<Tuple<bool, string>> ProcessTableAsync(BomTableAnnotation bomTable, TableAnnotation table, TableBoundryCondition tableCondition, CancellationToken token)
        {
            return Task<Tuple<bool, string>>.Run(() => {

                return ProcessTable(bomTable, table, tableCondition, token);
            });
        }

        bool CanExecuteStart()
        {
            if (IsIdle == true)
                if (string.IsNullOrWhiteSpace(SaveLocation) == false)
                    return true;

            return false;
        }
        #endregion 
    }

    struct TableBoundryCondition
    {
        public swTableHeaderPosition_e HeaderPosition { get; set; }
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }

        public int RowHeaderIndex { get; set; }
    }
}
