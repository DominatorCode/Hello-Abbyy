// � ABBYY. 2012.
// SAMPLES code is property of ABBYY, exclusive rights are reserved. 
// DEVELOPER is allowed to incorporate SAMPLES into his own APPLICATION and modify it 
// under the terms of License Agreement between ABBYY and DEVELOPER.

// ABBYY FlexiCapture Engine Samples
// This sample shows basic steps of ABBYY FlexiCapture Engine

using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using FCEngine;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Globalization;
using System.Reflection;
using System.Drawing;

namespace Hello
{
    public class HelloForm : System.Windows.Forms.Form
    {
        static IEngine _engine;
        static IEngineLoader _engineLoader;
        IFlexiCaptureProcessor _processor = null;
        IImageLoadingParams _imageLoadingParams = null;
        CustomPreprocessingImageSource _imageSource = null;

        private System.Windows.Forms.Button goButton;
        private System.Windows.Forms.StatusBar statusBar;
        private System.Windows.Forms.Button closeButton;
        private TextBox textBox1;
        private Label label1;
        private System.ComponentModel.Container components = null;
        private string exportFolder;
        private readonly List<string> _listExtensionsFilter = new List<string> {".tif",".png",".jpg",".bmp",".dcx",".pcx",".png",".jp2",".jpc",".jpeg",".jfif",".pdf",".tiff",".gif",".djvu",".djv",".jb2",".wdp"};

        private List<string> tmpGTDFoldersList = new List<string>();

        bool condHasAdditionalDocuments = false;
        bool condItemsListGDTNeedsInject = false;
        bool condItemsListKDTNeedsInject = false;
        string folderItemsListGDTName = "ItemsList";
        string folderItemsListKDTName = "ItemsList";
        string _errorText = "";
        string nameFolderKDT = "KDT";
        bool condKDTHasMainDocument = false;
        bool condKDTHasAdditionalDocument = false;

        bool condAllRecognizedCorrect = true;

        string _warningRecognizeResults = "";
        List<string> _listErrorsDocumentRecognizing = new List<string>();
        

        public HelloForm()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.goButton = new System.Windows.Forms.Button();
            this.statusBar = new System.Windows.Forms.StatusBar();
            this.closeButton = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // goButton
            // 
            this.goButton.Location = new System.Drawing.Point(24, 64);
            this.goButton.Name = "goButton";
            this.goButton.Size = new System.Drawing.Size(88, 28);
            this.goButton.TabIndex = 1;
            this.goButton.Text = "Go";
            this.goButton.Click += new System.EventHandler(this.goButton_Click);
            // 
            // statusBar
            // 
            this.statusBar.Location = new System.Drawing.Point(0, 101);
            this.statusBar.Name = "statusBar";
            this.statusBar.Size = new System.Drawing.Size(264, 22);
            this.statusBar.TabIndex = 8;
            // 
            // closeButton
            // 
            this.closeButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.closeButton.Location = new System.Drawing.Point(144, 64);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(87, 27);
            this.closeButton.TabIndex = 10;
            this.closeButton.Text = "Exit";
            this.closeButton.Click += new System.EventHandler(this.exit_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(24, 38);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(207, 20);
            this.textBox1.TabIndex = 11;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(155, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "������� ����� ��� ��������";
            // 
            // HelloForm
            // 
            this.AcceptButton = this.goButton;
            this.CancelButton = this.closeButton;
            this.ClientSize = new System.Drawing.Size(264, 123);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.statusBar);
            this.Controls.Add(this.goButton);
            this.Name = "HelloForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Hello C#";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        [STAThread]
        static void Main()
        {
            Application.Run(new HelloForm());
        }

        private void goButton_Click(object sender, System.EventArgs e)
        {
            closeButton.Enabled = false;
            goButton.Enabled = false;
            try
            {
                processImages();
                var test = GetImagesQualityInfo();

                //RunExternalExe(@"E:\VS Pro\DocNet Visual Editor v12\bin\Release\DocNet Visual Editor v12.exe", @"E:\VS Pro\Hello (C#) v12\bin\Debug\HelloFiles\FCEExport\���2018-10-31_11-35-08.tif");            

                //DeleteEmptyRows(@"E:\VS Pro\Hello (C#)\bin\Debug\HelloFiles\FCEExport");

                if (!string.IsNullOrEmpty(_warningRecognizeResults))
                    MessageBox.Show(_warningRecognizeResults);

                var errorsRecognize = GetRecognizingErrors();
                if (!String.IsNullOrEmpty(errorsRecognize))
                    MessageBox.Show(errorsRecognize);

                trace("��� � �������");
            }
            catch (Exception error)
            {
                showError(error.Message);
                MessageBox.Show(error.StackTrace);
                trace("����");
            }
            finally
            {
                goButton.Enabled = true;
                closeButton.Enabled = true;
                tmpGTDFoldersList.Clear();
            }
        }

        private void exit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        public void Scanning()
        {
            // Create an instance of ScanManager
            trace("Create an instance of ScanManager...");
            IScanManager scanManager = _engine.CreateScanManager();
            IStringsCollection sources = scanManager.ScanSources;

            string sourcesList = "������ ��������: ";
            string samplesFolder = AppDomain.CurrentDomain.BaseDirectory + "HelloFiles";
            
            if (sources.Count > 0)
            {
                for (int i = 0; i < sources.Count; i++)
                {
                    if (sourcesList.Length > 0)
                    {
                        sourcesList += ", ";
                    }
                    sourcesList += '\'' + sources[i] + '\'';
                }
            }
            else
            {
                sourcesList = "������� �� �������";
            }
            trace(sourcesList);
            // If at least one scan source found
            if (sources.Count > 0)
            {
                // In this sample we will be using the first found scan source
                string scanSource = sources[0];

                // You can optionally change the scan source settings or leave defaults
                // trace( "Configure scan source " + '\'' + scanSource + '\'' );
                // IScanSourceSettings sourceSettings = scanManager.get_ScanSourceSettings( scanSource );
                // sourceSettings.PictureMode = ScanPictureModeEnum.SPM_Grayscale;
                // scanManager.set_ScanSourceSettings( scanSource, sourceSettings );

                // Prepare a directory to store scanned images
                string scanFolder = samplesFolder + "\\FCEScanning";
                System.IO.Directory.CreateDirectory(scanFolder);
                bool IsInteractive = false;
                if (IsInteractive)
                {
                    // If the scenario is being run in interactive mode you can try scanning by uncommenting the lines below
                    trace("Scanning is disabled.");

                    //Get a single image from the scan source
                    IStringsCollection imageFiles = scanManager.Scan(scanSource, scanFolder, false);

                    System.IO.File.Delete(imageFiles[0]);

                }
                else
                {
                    trace("Running in non-interactive mode. Scanning skipped.");
                }
            }
        }

        private void processImages()
        {
            trace("Loading FlexiCapture Engine...");
            LoadEngine();
            
            if (condUsingCustomImageSource)
                _imageSource = new CustomPreprocessingImageSource(_engine);
            else
            {
                _imageLoadingParams = _engine.CreateImageLoadingParams();
                _imageLoadingParams.CorrectSkewByBlackSquares = false;
                _imageLoadingParams.CorrectSkewByText = true;
                _imageLoadingParams.CorrectSkewByBlackSeparators = true;

                _imageLoadingParams.AutocorrectResolution = true;

                // ������� ���������
                _imageLoadingParams.DiscardImageColor = true;
                _imageLoadingParams.UseFastBinarization = true;
                //_imageLoadingParams.ConvertToGray = true;
                //_imageLoadingParams.WhitenBackground = true;
                //_imageLoadingParams.RemoveGarbage = true;
                //_imageLoadingParams.UseAutocrop = true;
                //_imageLoadingParams.RemoveColorMarks = true;


                //_imageLoadingParams.OverwriteResolution = true;
                //_imageLoadingParams.XResolutionToOverwrite = 300;
                //_imageLoadingParams.YResolutionToOverwrite = 300;



                //_imageLoadingParams.SourceContentReuseMode = true;


                //_imageLoadingParams.TreatImageAsPhoto = true;
            }



            var isValid = CheckVersionValid();

           
            var version = _engine.Version;
            //_engine.StartLogging(@"C:\Users\seral\Desktop\FCEExport\Abbyy12\123.txt");
            try
            {
                string samplesFolder = AppDomain.CurrentDomain.BaseDirectory + "HelloFiles";

                _processor = _engine.CreateFlexiCaptureProcessor();
                // ��������� �������
                foreach (string file in Directory.EnumerateFiles(samplesFolder + @"\SampleProject\Templates", "*.fcdot"))
                {
                    _processor.AddDocumentDefinitionFile(file);                                      
                }
                foreach (string file in Directory.EnumerateFiles(samplesFolder + @"\SampleProject\Templates", "*.cfl"))
                {
                    _processor.AddClassificationTreeFile(file);
                    

                }
                SetTemplatesDirectory(samplesFolder + @"\SampleProject");
                
                var test = GetTemplateVersion();

                trace("Adding images to process...");


                string[] imagesForRecognize = Directory.GetFiles(samplesFolder + "\\SampleImages", "*");

                
                if (condUsingCustomImageSource)
                {
                    _imageSource.UseBinarization = true; _imageSource.ModeBinarization = BinarizationModeEnum.BM_Fast;
                    //_imageSource.IsColorBackImage = true;
                    _imageSource.UseCleanUpImage = true;
                }


                foreach (string file in imagesForRecognize)
                {
                    if (_listExtensionsFilter.Contains(Path.GetExtension(file)))
                    {
                        if (condUsingCustomImageSource)
                            _imageSource.AddImageFile(file);
                        else
                            _processor.AddImageFile(file);
                    }
                    
                }

                if (condUsingCustomImageSource)
                    _processor.SetCustomImageSource(_imageSource);
                else
                    _processor.SetImageLoadingParams(_imageLoadingParams);


                _processor.SetUseFirstMatchedDocumentDefinition(true);
                
                trace("Recognizing images and exporting results...");
                int count = 0;
                HashSet<string> listDeclarationRecognizeErrors = new HashSet<string>();

                while (true)
                {
                    

                    if (String.IsNullOrEmpty(this.textBox1.Text))
                        exportFolder = samplesFolder + "\\FCEExport";
                    else
                        exportFolder = new Uri(this.textBox1.Text).LocalPath;

                    IDocument document = _processor.RecognizeNextDocument();
                    
                    
                    IDocumentDefinition docDefinition = null;
                    CleanDirectories(exportFolder);
                    
                    if (document == null)
                    {
                        IProcessingError error = _processor.GetLastProcessingError();
                        if (error != null)
                        {
                            // Processing error
                            showError(error.MessageText());
                            Marshal.ReleaseComObject(error);
                            continue;
                        }
                        else {
                            // No more images
                            break;
                        }
                    }
                    else {
                        docDefinition = document.DocumentDefinition;

                        if (docDefinition == null)
                        {
                            // Couldn't find matching template for the image. In this sample this is an error.
                            // In other scenarios this might be normal
                            string undefinedFolder = exportFolder + "\\undefined";
                            Directory.CreateDirectory(undefinedFolder);
                            IPage page = document.Pages[0];
                            IImageDocument pageImageDocument = page.ReadOnlyImage;
                            IImage bwImage = pageImageDocument.BlackWhiteImage;

                            #region Numbering unrecognized pages

                            // ��������� �������������� ������� � ��������������� ���������
                            int pageIndex = page.SourceImageInfo.PageIndex;
                            int pageNumber;

                            if (pageIndex == 0)
                            {
                                pageNumber = 1;
                            }
                            else
                            {
                                pageIndex++;
                                pageNumber = pageIndex;
                            }

                            #endregion

                            string name = Path.GetFileNameWithoutExtension(page.OriginalImagePath);
                            IImage image = document.Pages[0].Image.BlackWhiteImage;

                            image.WriteToFile(
                                undefinedFolder + "\\" + name + "_p" + pageNumber + ".tif",
                                ImageFileFormatEnum.IFF_Tif, 
                                null, 
                                ImageCompressionTypeEnum.ICT_CcittGroup4, 
                                null);

                            _listErrorsDocumentRecognizing.Add("�� ������� ���������� �����������: " + document.Pages[0].OriginalImagePath);
                            continue;
                            
                        }
                    }

                    CreateExportParameters();
                   
                    _processor.ExportDocument(document, exportFolder);

                    // ���� ������������ ���������� ������ ��������� �����������,
                    // ��������� �� ��������
                    if (!condUsingCustomImageSource)
                        CheckIsImageSuitableForOcr(document);
                    // ������������ ������ � ��������� �����
                    var directory = new DirectoryInfo(exportFolder);
                    string myFile = directory.GetFiles()
                                 .OrderByDescending(f => f.LastWriteTime)
                                 .First().Name;
                    string nameDocument = Path.GetFileNameWithoutExtension(myFile);
                    Directory.CreateDirectory(exportFolder + "\\visualeditorfiles");

                    //document.AsCustomStorage.SaveToFile(exportFolder + "\\visualeditorfiles\\" + nameDocument + ".mydoc");

                    #region Extract blocks

                    // ������� ����������� ������
                    trace("Extract the field image...");

                    string fileDirectory = exportFolder + "\\blocks\\" + nameDocument;
                    Directory.CreateDirectory(fileDirectory);

                    int pagesCount = document.Pages.Count;
                    for (int p = 0; p < pagesCount; p++)
                    {
                        int blocksCount = document.Pages[p].Blocks.Count;

                        for (int i = 0; i < blocksCount - 1; i++)
                        {
                            IPage page = document.Pages[p];
                            IBlock block = page.Blocks.Item(i);

                            if (block.Field != null)
                            {
                                // ���� ���� - �������
                                if (block.Field.Name.Equals("Table1"))
                                {
                                    ExtractTableCells(page, block, docDefinition.Name, fileDirectory);
                                }
                                else
                                {
                                    string filename = docDefinition.Name + "_" + block.Field.Name + ".jpg";
                                    ExtractBlock(page, block, filename, fileDirectory);
                                }
                            }
                        }
                    }

                    #endregion

                    Marshal.ReleaseComObject(docDefinition);
                    Marshal.ReleaseComObject(document);

                    DeleteEmptyRows(exportFolder);

                    count++;                                     
                }
            }
            finally
            {
                UnloadEngine();
            }
        }

        // ��������� ����� � �����������
        public void ExtractBlock(IPage page, IBlock block, string filename, string fileDirectory)
        {
            IImageDocument pageImageDocument = page.ReadOnlyImage;
            IImage bwImage = pageImageDocument.BlackWhiteImage;

            IImageModification modification = _engine.CreateImageModification();
            modification.ClipRegion = block.Region;

            bwImage.WriteToFile
                (
                    fileDirectory + "\\" + filename,
                    ImageFileFormatEnum.IFF_Tif, 
                    modification, 
                    ImageCompressionTypeEnum.ICT_CcittGroup4, 
                    null
                );
        }

        // ������������ ����� ����� ������� � �����������
        public void ExtractTableCells(IPage page, IBlock block, string definitionName, string fileDirectory)
        {
            int rows = block.AsTableBlock().RowsCount;
            int columns = block.AsTableBlock().BoundColumnsCount;

            for (int c = 0; c < columns; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    ITableBlock table = block.AsTableBlock();
                    IBlock cell = table.Cell[c, r];
                    int pageNumber = page.Index + 1;
                    string cellName = definitionName + "_Table_" + cell.Field.Name + "_" + c + r + "_p" + pageNumber + ".jpg";
                    ExtractBlock(page, cell, cellName, fileDirectory);
                }
            }
        }

        /// <summary>�������������� ������ ������������� ABBYY</summary>
        /// <returns>������ - ������ ��������</returns>
        public bool LoadEngine()
        {
            if (_engine != null)
            {
                return true;
            }

            try
            {
                traceText += "�������� ��� Loader";


                _engineLoader = new OutprocLoader();
                IWorkProcessControl workProcess = (IWorkProcessControl)_engineLoader;
                workProcess.SetParentProcessId(System.Diagnostics.Process.GetCurrentProcess().Id);


                traceText += " ,��������� ������ _engineLoader.Load()";

                _engineLoader.CustomerProjectId = Seral;
                _engineLoader.LicensePath = "";
                _engineLoader.LicensePassword = "";

                _engine = _engineLoader.GetEngine();

                ILicense license = _engine.CurrentLicense;
                if (!license.IsActivated)
                {
                    errorText = "������ �������� �� ������������, ���� �������� �������� ����.";
                    //MessageBox.Show(errorText);
                    UnloadEngine();
                    return false;
                }

                // ���� ������������ ��������� �������� ��������� ��
                // � ������� ���� ��������
                if (license.RemainingUnits == 0)
                {
                    if (!ChangeLicense())
                    {
                        errorText = "������� �������� ����� 0 ������� ��� �������������, �� ������� ������� �������� �� ������";
                        throw new Exception(errorText);
                    }
                }

                Marshal.ReleaseComObject(license);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    case -2147221164:
                        {
                            errorText = "ABBYY FlexiCapture not registered in system with CLSID {C0003004-0000-48FF-9197-57B7554849BA}";
                            errorText = errorText + Environment.NewLine + "���������� ��������� ����������� ��������� (regsvr) � ������ ���������";
                            //MessageBox.Show(errorText);
                            break;
                        }
                    default:
                        {
                            errorText = "Failed to load ABBYY Engine (" + ex.ErrorCode + ", '" + ex.Message + "')";
                            //MessageBox.Show(errorText);
                            break;
                        }
                }
                return false;
            }
            catch (Exception e)
            {
                if (e.Message.Contains("80070005"))
                {
                    // To use LocalServer under a special account you must add this account to 
                    // the COM-object's launch permissions (using DCOMCNFG or OLE/COM object viewer)
                    errorText = "Launch permission for the work-process COM-object is not granted. Use DCOMCNFG to change security settings for the object. (" + e.Message + ")";
                    errorText = errorText + Environment.NewLine + "���������� ��������� ����� ������� ��� ��������� � ������ ���������";
                    //MessageBox.Show(errorText);
                    //throw new Exception(@"Launch permission for the work-process COM-object is not granted.
                    // Use DCOMCNFG to change security settings for the object. (" + e.Message + ")");
                }
                else
                {
                    errorText = "Failed to load ABBYY Engine " + e.Message;
                    //MessageBox.Show(errorText);
                }

                return false;
            }

            traceText += " ,������ ��������";

            return true;
        }

        /// <summary>��������� ������ ������������� ABBYY � �������</summary>
        public void UnloadEngine()
        {
            if (_engine != null)
            {
                UnloadProcessor();
                if (_engineLoader != null)
                {
                    _engineLoader.Unload();
                    Marshal.ReleaseComObject(_engineLoader);
                    _engineLoader = null;
                }
                Marshal.ReleaseComObject(_engine);
                _engine = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        void UnloadProcessor()
        {
            if (_imageSource != null)
            {
                _imageSource.Dispose();
                _imageSource = null;
            }

            if (_exportParams != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_exportParams);
                _exportParams = null;
            }

            if (_imageLoadingParams != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_imageLoadingParams);
                _imageLoadingParams = null;
            }

            if (_processor != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_processor);
                _processor = null;
            }

            _processorLoad = false;
        }

        public void assert(bool condition)
        {
            if (!condition)
            {
                StackTrace st = new StackTrace(1, true);
                StackFrame sf = st.GetFrame(0);
                throw new AssertionFailedException(sf);
            }
        }


        private void trace(string text)
        {
            statusBar.Text = text;
            statusBar.Update();
        }

        private void showError(string text)
        {
            MessageBox.Show(this, text, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            exportFolder = this.textBox1.Text;
        }

        // ���������� �������������� ������ �� 3� ������� ��� ����������� ������ ���������
        void AddRootNode(string paramPathExportFolder)
        {
            string[] xmlFiles = Directory.GetFiles(paramPathExportFolder, "*.xml");

            foreach (string xmlFile in xmlFiles)
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlFile);
                XmlNode itemNodes = xmlDocument.DocumentElement.FirstChild;

                // ����� �������� ���� ���������
                XmlNodeList nodeDocType = xmlDocument.GetElementsByTagName("_DocType");
                string nameNodeNew = "";
                if (nodeDocType.Count > 0)
                    nameNodeNew = nodeDocType[0].Value;

                // �������� �� ������� ����������m��� ������
                XmlNodeList nodeDocTypeExists = null;
                if (!String.IsNullOrEmpty(nameNodeNew))
                    nodeDocTypeExists = xmlDocument.GetElementsByTagName(nameNodeNew);

                if (nodeDocTypeExists != null)
                    if (nodeDocTypeExists.Count == 0)
                    {
                        XmlNode nodeNew = xmlDocument.CreateNode(XmlNodeType.Element, nameNodeNew, null);
                        itemNodes.InsertBefore(nodeNew, itemNodes.FirstChild);
                        xmlDocument.Save(xmlFile);
                    }               
            }
               
        }

        void DeleteEmptyRows(string paramPathExportFolder)
        {
            string[] xmlFiles = Directory.GetFiles(paramPathExportFolder, "*.xml");

            foreach (string xmlFile in xmlFiles)
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlFile);

                XmlNodeList itemNodes = xmlDocument.GetElementsByTagName("_Table");
                string nodeTaxIncluded = null;

                bool condTaxField = false;
                XmlNodeList listNodesTax = xmlDocument.GetElementsByTagName("_TaxIncluded");
                if (listNodesTax.Count == 1)
                {
                    nodeTaxIncluded = listNodesTax[0].InnerText;

                    if (String.IsNullOrEmpty(nodeTaxIncluded))
                        condTaxField = true;
                }
                                                   
                if (itemNodes.Count == 0)
                    itemNodes = xmlDocument.GetElementsByTagName("_Table1");

                if (itemNodes.Count > 0)
                {
                    
                    for (int i = itemNodes.Count - 1; i >= 0; i--)
                    {
                        string fieldDescript;
                        try
                        {
                            fieldDescript = itemNodes[i].SelectSingleNode("_Descript").InnerText;
                        }
                        catch
                        {
                            break;
                        }

                        if (String.IsNullOrEmpty(fieldDescript) || fieldDescript.Length < 3)
                        {
                            itemNodes[i].ParentNode.RemoveChild(itemNodes[i]);
                        }
                        else if (condTaxField) // ��������� ��� ������ � ��������� ��� ���
                        {                           
                            double Price = 0;
                            double Qty = 0;
                            double Cost = 0;
                            double SumWTax = 0;

                            var nodeSumWithTax = itemNodes[i].SelectSingleNode("_SumWithTax").InnerText;
                            nodeSumWithTax = nodeSumWithTax.Replace('.', ',');

                            if (nodeSumWithTax != null)
                            {
                                Double.TryParse(nodeSumWithTax, NumberStyles.Currency, null, out SumWTax);
                                string nodeCost = null;
                                try
                                {
                                    nodeCost = itemNodes[i].SelectSingleNode("_Cost").InnerText;
                                    nodeCost = nodeCost.Replace('.', ',');
                                }
                                catch
                                {

                                }
                                

                                if (String.IsNullOrEmpty(nodeCost))
                                {

                                    string nodePrice = null;
                                    string nodeQty = null;

                                    try
                                    {
                                        nodePrice = itemNodes[i].SelectSingleNode("_Price").InnerText;
                                        nodeQty = itemNodes[i].SelectSingleNode("_Qty").InnerText;
                                    }
                                    catch { }
                                    

                                    if (nodePrice != null & nodeQty != null)
                                    {
                                        nodePrice = nodePrice.Replace('.', ',');
                                        nodeQty = nodeQty.Replace('.', ',');
                                        Double.TryParse(nodePrice, NumberStyles.Currency, null, out Price);
                                        Double.TryParse(nodeQty, NumberStyles.Currency, null, out Qty);
                                    }

                                    if (Price > 0 & Qty > 0)
                                        Cost = Price * Qty;
                                }
                                else
                                    Double.TryParse(nodeCost, NumberStyles.Currency, null, out Cost);

                                if (Cost > 0)
                                {
                                    if (Math.Abs(SumWTax - Cost) > 0.1)
                                        nodeTaxIncluded = "false";
                                    else
                                        nodeTaxIncluded = "true";

                                    listNodesTax[0].InnerText = nodeTaxIncluded;
                                    condTaxField = false;
                                }
                            }

                            
                        }
                    }

                    xmlDocument.Save(xmlFile);

                }
            }
        }

        

        



        static string GetFieldValue(IDocument document, string fieldName)
        {
            IField resultField = findField(document, fieldName);

            if (resultField != null)
                return resultField.Value.AsString.ToLower();
            else
                return "";
        }


        static IField findField(IDocument document, string name)
        {
            IField root = document.AsField;
            return recursiveFindField(root, name);
        }

        static IField recursiveFindField(IField node, string name)
        {
            IFields children = node.Children; 
            if (children != null)
            {
                for (int i = 0; i < children.Count; i++)
                {
                    IField child = children[i]; 
                    if (child.Name == name)
                    {
                        return child;
                    }
                    else {
                        IField found = recursiveFindField(child, name);
                        if (found != null)
                        {
                            return found;
                        }
                    }
                }
            }
            return null;
        }

        public string GetRecognizingErrors()
        {
            if (_listErrorsDocumentRecognizing.Count > 0)
                return string.Join(System.Environment.NewLine, _listErrorsDocumentRecognizing);
            else
                return "";
        }

        string SetDeclarationFileName(string nameDocument, string numberPage, string pathFolderExport)
        {
            int numberPageCast = 0;            
            if (Int32.TryParse(numberPage, out numberPageCast))
            {
                numberPage = numberPage.PadLeft(5, '0');
                return nameDocument + "_" + numberPage;
            }               
            else
            {
                numberPageCast = Directory.GetFiles(pathFolderExport).Count();
                string nameFinal = nameDocument + "_" + (numberPageCast + 1).ToString().PadLeft(5, '0');
                if (File.Exists(nameFinal))
                {
                    _listErrorsDocumentRecognizing.Add("����� ��� ���������� �������� ����������");
                    Random rnd = new Random();
                    int rndnumber = rnd.Next(999);
                    return nameFinal + "_" + rndnumber;
                }
                else
                    return nameFinal;
            }

        }

        bool ChangeLicense()
        {
            bool condLicenseChanged = false;
            ILicenseCollection listLicenses = _engine.Licenses;
            List<int> listCountReminingUnits = new List<int>();

            if (listLicenses.Count > 1)
            {
                for (int i = 0; i < listLicenses.Count; i++)
                {
                    listCountReminingUnits.Add(listLicenses[i].RemainingUnits);
                }

                if (listCountReminingUnits.Max() > 0)
                {
                    try
                    {
                        _engine.SetCurrentLicense(listLicenses[listCountReminingUnits.IndexOf(listCountReminingUnits.Max())], FceConfig.GetCustomerProjectId());
                    }
                    catch (Exception Ex)
                    {
                        _errorText = "�� ������� ������� ��������: " + Ex.Message;
                        throw new Exception(_errorText);
                    }

                    condLicenseChanged = true;
                }
            }
            else
                return false;

            return condLicenseChanged;

        }

        bool CheckItemsListMissingDocuments(bool condILGTD, bool condILKDT, string pathILGDT, string pathILKDT)
        {
            // ���� � ��� ����������� ���� �����
            // �������� � ���
            if (condILGTD & !condILKDT) 
            {
                Directory.CreateDirectory(pathILKDT);

                // �������� ��� �����
                foreach (string newPath in Directory.GetFiles(pathILGDT, "*.*",
                    SearchOption.AllDirectories))
                    File.Copy(newPath, newPath.Replace(pathILGDT, pathILKDT), true);

                return true;
            }
            return
                false;
        }

        public bool RunExternalExe(string pathExe, string pathFile)
        {
            try
            {
                string fileName = Path.GetFileName(pathFile);
                string pathDirectory = Directory.GetParent(pathFile).FullName;

                Process p = new Process();
                p.StartInfo.FileName = pathExe;
                p.StartInfo.Arguments = "\"" + pathDirectory + "\"";
                p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;

                if (!String.IsNullOrEmpty(fileName))
                    p.StartInfo.Arguments += " \"" +  fileName + "\"";

                p.Start();
                //p.WaitForExit();
                return true;
            }
            catch (Exception ex)
            {
                _errorText = "������ ������� �������� exe: " + ex.Message + Environment.NewLine + ex.StackTrace;
                throw new Exception(_errorText);
            }
        }

        void startClean()
        {

        }

        public void CleanDirectories(string pathFolderClean, int countDaysStore = 3)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                IEnumerable<string> directoriesPath = Directory.GetDirectories(pathFolderClean);
                foreach (string dPath in directoriesPath)
                {
                    string[] files = Directory.GetFiles(dPath);

                    foreach (string file in files)
                    {
                        FileInfo fi = new FileInfo(file);
                        string time = fi.LastWriteTime.ToShortDateString();

                        if (fi.LastWriteTime < DateTime.Now.AddDays(-countDaysStore))
                            fi.Delete();
                    }
                }
            }).Start();            
        }

        
        string _pathDirectoryTemplates = "";
        List<string> _listTemplatesInfo = null;
        private string _versionValid = "12.1.24";
        private string errorText;
        private string traceText;
        private bool _processorLoad;
        private bool condUsingCustomImageSource = false;
        IFileExportParams _exportParams = null;

        public string Seral { get; private set; } = "aYJ5eBTzpHoXviTluPZp";

        public string GetSingleTemplateVersion(string pathFileTemplate)
        {
            if (_engine == null)
                throw new Exception("������ FlexiCapture �� ���������������");
            if (!File.Exists(pathFileTemplate))
                return "";

            // Load a Document Definition from file           
            if (Path.GetExtension(pathFileTemplate) == ".fcdot")
            {
                IDocumentDefinition documentDefinition = _engine.CreateDocumentDefinition();
                ICustomStorage customStorage = documentDefinition as ICustomStorage;
                customStorage.LoadFromFile(pathFileTemplate);
                var s = documentDefinition.GUID;
                return documentDefinition.Description;
            }

            if (Path.GetExtension(pathFileTemplate) == ".cfl")
            {
                /*IClassificationTree documentDefinition = engine.CreateClassificationTree();
                documentDefinition.LoadFromFile(pathFileTemplate);
                IStringsCollection collectionClasses = documentDefinition.GetClassNames();
               
                return collectionClasses.ToString() + " ������ ";*/
                return "";
            }

            return "";
        }

        public string GetTemplateVersion()
        {
            if (String.IsNullOrEmpty(_pathDirectoryTemplates))
            {
                _errorText = "ERROR: FC_Recognize: GetTemplateVersion: " + Environment.NewLine + "�� ����� ���� � ���������� � ���������";
                throw new Exception(_errorText);
            }


            if (_listTemplatesInfo == null)
            {
                _listTemplatesInfo = new List<string>();
                foreach (string file in Directory.EnumerateFiles(_pathDirectoryTemplates, "*.fcdot"))
                {
                    var version = GetSingleTemplateVersion(file);
                    if (!String.IsNullOrEmpty(version))
                        _listTemplatesInfo.Add(version);
                }
            }

            if (_listTemplatesInfo.Count > 0)
            {
                var res = _listTemplatesInfo.Last();
                _listTemplatesInfo.RemoveAt(_listTemplatesInfo.Count - 1);
                return res;
            }
            else
            {
                _listTemplatesInfo = null;
                return "";
            }
        }

        public void SetTemplatesDirectory(string pathDirectoryTemplates)
        {
            pathDirectoryTemplates = pathDirectoryTemplates + "\\Templates";
            if (!Directory.Exists(pathDirectoryTemplates))
                throw new Exception("���������� �����������: " + pathDirectoryTemplates);

            _pathDirectoryTemplates = pathDirectoryTemplates;
        }


        public bool SetLoggingPath(string pathLogging)
        {
            if (String.IsNullOrEmpty(pathLogging))
                AppParameters.LoggingMode = false;

            string pathDirectoryFileLogging = Path.GetDirectoryName(pathLogging);
            if (Directory.Exists(pathDirectoryFileLogging))
            {
                AppParameters.PathFileLog = pathLogging;
            }
            else
            {
                try
                {
                    Directory.CreateDirectory(pathDirectoryFileLogging);
                }
                catch (Exception ex)
                {
                    AppParameters.TextError.Add("SetLoggingPath: �� ������� ������� ����������: " + pathDirectoryFileLogging + Environment.NewLine + ex.Message);
                    return false;
                }
            }

            try
            {
                File.GetAccessControl(pathDirectoryFileLogging);
                AppParameters.LoggingMode = true;
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                AppParameters.TextError.Add("��� ���� �� ������ � ����������: " + pathDirectoryFileLogging);
                return false;
            }

        }

        bool CheckVersionValid()
        {
            bool isValid = false;
            string versionEngine;


            if (_engine != null)
            {
                versionEngine = _engine.Version;

                var offset = versionEngine.IndexOf('.');
                offset = versionEngine.IndexOf('.', offset + 1);
                var result = versionEngine.IndexOf('.', offset + 1);
                if (result != -1)
                    versionEngine = versionEngine.Substring(0, result);

                if (String.Compare(versionEngine, _versionValid) == 0)
                    isValid = true;

            }

            return isValid;
        }

        public string GetImagesQualityInfo()
        {
            AppParameters._listInfoQualityImages.Add("���������� ��������� ������ ��� ��������� �������� ������������� ����������");

            if (!String.IsNullOrEmpty(AppParameters._warnColorTypeBad))
            {
                AppParameters._listInfoQualityImages.Add(AppParameters._warnColorTypeBad);
                AppParameters._warnColorTypeBad = "";
            }


            if (!String.IsNullOrEmpty(AppParameters._warnResolutionImageBad))
            {
                AppParameters._listInfoQualityImages.Add(AppParameters._warnResolutionImageBad);
                AppParameters._warnResolutionImageBad = "";
            }


            string infoOCR = "";
            if (AppParameters._listInfoQualityImages.Count > 1)
            {
                infoOCR = string.Join(System.Environment.NewLine, AppParameters._listInfoQualityImages);

                AppParameters._isSuitableForOcr = true;
            }

            AppParameters._listInfoQualityImages.Clear();
            return infoOCR;
        }

        /// <summary>���������, ��� ����������� �� ������� � ����� ���������� 300 dpi</summary>
        /// <param name="pDocument">������������ ��������</param>
        void CheckIsImageSuitableForOcr(IDocument pDocument)
        {
            
            IImageDocument imageDoc = pDocument.Pages[0].ReadOnlyImage;

            if (AppParameters._isSuitableForOcr)
            {
                if (String.IsNullOrEmpty(AppParameters._warnResolutionImageBad))
                    if (pDocument.Pages[0].Image.SourceImageXResolution < 300 || pDocument.Pages[0].Image.SourceImageYResolution < 300)
                    {
                        AppParameters._isSuitableForOcr = false;
                        AppParameters._warnResolutionImageBad = "����������� ���������� ������� ��� ������������� ������ ���� �� 300 dpi" + Environment.NewLine +
                           "��� �� ��������� �������� ����� �������, ����������� - ������� ������ ��� �����-�����";
                    }                                          
            }
        }


        void CreateExportParameters()
        {
            _exportParams = _engine.CreateFileExportParams();
            _exportParams.FileFormat = FileExportFormatEnum.FEF_XML;
            //_exportParams.XMLParams.ExportErrors = true;
            _exportParams.FileNamePattern = "<DocumentDefinition><Time>";
            _exportParams.FileOverwriteRule = FileOverwriteRuleEnum.FOR_Rename;
            _exportParams.ExportOriginalImages = true;
        }
    }
}