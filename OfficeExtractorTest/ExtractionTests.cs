using System;
using System.Collections.Generic;
using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeExtractorTest
{
    [TestClass]
    public class ExtractionTests
    {
        private readonly List<string> _tempFolders = new List<string>();

        [TestMethod]
        [ExpectedException(typeof(OEFileTypeNotSupported))]
        public void FileTypeNotSupported()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\filetypenotsupported.txt", outputFolder);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsCorrupt))]
        public void FileIsCorrupt()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A corrupt compound document.doc", outputFolder);
        }

        #region Microsoft Office Word tests
        [TestMethod]
        public void DocWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOC word document without embedded files.doc", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void DocWith7EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOC word document with 7 embedded files.doc", outputFolder);
            Assert.IsTrue(files.Count == 7);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void DocWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A DOC word document with password.doc", outputFolder);
        }

        [TestMethod]
        public void DocxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOCX word document without embedded files.docx", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOCX word document with 7 embedded files.docx", outputFolder);
            Assert.IsTrue(files.Count == 7);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void DocxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A DOCX word document with password.docx", outputFolder);
        }
        #endregion

        #region Microsoft Office Excel tests
        [TestMethod]
        public void XlsWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A XLS excel document without embedded files.xls", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void XlsWith2EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A XLS excel document with 2 embedded files.xls", outputFolder);
            Assert.IsTrue(files.Count == 2);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void XlsWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A XLS excel document with password.xls", outputFolder);
        }

        [TestMethod]
        public void XlsxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A XLSX excel document without embedded files.xlsx", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void XlsxWith2EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A XLSX excel document with 2 embedded files.xlsx", outputFolder);
            Assert.IsTrue(files.Count == 2);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void XlsxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A XLSX excel document with password.xlsx", outputFolder);
        }
        #endregion

        #region Microsoft Office PowerPoint tests
        [TestMethod]
        public void PptWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A PPT PowerPoint document without embedded files.ppt", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void PptWith3EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A PPT powerpoint document with 3 embedded files.ppt", outputFolder);
            Assert.IsTrue(files.Count == 3);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void PptWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A PPT PowerPoint document with password.ppt", outputFolder);
        }

        [TestMethod]
        public void PptxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A PPTX PowerPoint document without embedded files.pptx", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void PptxWith3EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A PPTX powerpoint document with 3 embedded files.pptx", outputFolder);
            Assert.IsTrue(files.Count == 3);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void PptxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A PPTX PowerPoint document with password.pptx", outputFolder);
        }
        #endregion

        #region Open Office Writer tests
        [TestMethod]
        public void OdtWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\An ODT document without embedded files.odt", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void OdtWith8EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\An ODT document with 8 embedded files.odt", outputFolder);
            Assert.IsTrue(files.Count == 8);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void OdtWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\An ODT document with password.odt", outputFolder);
        }
        #endregion
        
        #region Open Office Calc tests
        [TestMethod]
        public void OdsWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\An ODS document without embedded files.ods", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void OdsWith2EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\An ODS document with 2 embedded files.ods", outputFolder);
            Assert.IsTrue(files.Count == 2);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void OdsWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\An ODS document with password.ods", outputFolder);
        }
        #endregion

        #region Open Office Impress tests
        [TestMethod]
        public void OdpWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\An ODP document without embedded files.odp", outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void OdpWith3EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\An ODP document with 3 embedded files.odp", outputFolder);
            Assert.IsTrue(files.Count == 3);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void OdpWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\An ODP document with password.odp", outputFolder);
        }
        #endregion

        #region Helper methods
        /// <summary>
        /// Creates a new temporary folder and returns the path to it
        /// </summary>
        /// <returns></returns>
        private string CreateTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            _tempFolders.Add(tempDirectory);
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }

        private static void DeleteTemporaryFolder(string folder)
        {
            try
            {
                if (Directory.Exists(folder))
                    Directory.Delete(folder, true);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            { }
        }

        [TestCleanup]
        public void CleanUp()
        {
            foreach (var tempFolder in _tempFolders)
                DeleteTemporaryFolder(tempFolder);
        }
        #endregion
    }
}
