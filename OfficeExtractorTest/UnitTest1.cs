using System;
using System.IO;
using DocumentServices.Modules.Extractors.OfficeExtractor.Exceptions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeExtractorTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        [ExpectedException(typeof(OEFileTypeNotSupported))]
        public void FileTypeNotSupported()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\filetypenotsupported.txt", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }

        #region Word tests
        [TestMethod]
        public void DocWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOC word document without embedded files.doc", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void DocWithout7EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOC word document with 7 embedded files.doc", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 7);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void DocWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A DOC word document with password.doc", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }

        [TestMethod]
        public void DocxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOCX word document without embedded files.docx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A DOCX word document with 7 embedded files.docx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 7);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void DocxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A DOCX word document with password.docx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }
        #endregion

        #region Excel tests
        [TestMethod]
        public void XlsWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A XLS excel document without embedded files.xls", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void XlsWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A XLS excel document with password.xls", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }

        [TestMethod]
        public void XlsxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A XLSX excel document without embedded files.xlsx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void XlsxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A XLSX excel document with password.xlsx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }
        #endregion

        #region PowerPoint tests
        [TestMethod]
        public void PptWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A PPT PowerPoint document without embedded files.ppt", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void PptWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A PPT PowerPoint document with password.ppt", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }

        [TestMethod]
        public void PptxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            var files = extractor.ExtractToFolder("TestFiles\\A PPTX PowerPoint document without embedded files.pptx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
            Assert.IsTrue(files.Count == 0);
        }

        [TestMethod]
        [ExpectedException(typeof(OEFileIsPasswordProtected))]
        public void PptxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new DocumentServices.Modules.Extractors.OfficeExtractor.Extractor();
            extractor.ExtractToFolder("TestFiles\\A PPTX PowerPoint document with password.pptx", outputFolder);
            DeleteTemporaryFolder(outputFolder);
        }
        #endregion

        #region Helper methods
        /// <summary>
        /// Creates a new temporary folder and returns the path to it
        /// </summary>
        /// <returns></returns>
        private static string CreateTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
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
        #endregion
    }
}
