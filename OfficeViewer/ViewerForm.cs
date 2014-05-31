using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage;

namespace OfficeViewer
{
    public partial class ViewerForm : Form
    {
        readonly List<string> _tempFolders = new List<string>(); 

        public ViewerForm()
        {
            InitializeComponent();
        }

        private void ViewerForm_Load(object sender, EventArgs e)
        {
            Closed += ViewerForm_Closed;
        }

        void ViewerForm_Closed(object sender, EventArgs e)
        {
            foreach (var tempFolder in _tempFolders)
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
        }

        private void SelectButton_Click(object sender, EventArgs e)
        {
            // Create an instance of the opeKn file dialog box.
            var openFileDialog1 = new OpenFileDialog
            {
                // ReSharper disable once LocalizableElement
                Filter = "Microsoft Office files(*.doc, *.dot, *.xls, *.ppt)|*.doc;*.dot;*.xls;*.ppt",
                FilterIndex = 1,
                Multiselect = false
            };

            // Process input if the user clicked OK.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Open the selected file to read.
                string tempFolder = null;

                try
                {
                    tempFolder = GetTemporaryFolder();
                    _tempFolders.Add(tempFolder);

                    var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
                    extractor.ExtractFromWord(openFileDialog1.FileName, tempFolder);
                    FilesListBox.Items.Clear();

                    //foreach (var file in files)
                    //    FilesListBox.Items.Add(file);
                }
                catch (Exception ex)
                {
                    if (tempFolder != null && Directory.Exists(tempFolder))
                        Directory.Delete(tempFolder, true);

                    MessageBox.Show(ex.Message);
                }
            }
        }

        public string GetTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }
}
