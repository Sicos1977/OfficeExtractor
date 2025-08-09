using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeExtractor;
using OfficeExtractor.Exceptions;
using PasswordProtectedChecker.Exceptions;

//
// ExtractionTest.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeExtractorTest
{
    [TestClass]
    public class ExtractionTests
    {
        private readonly List<string> _tempFolders = new List<string>();

        //[TestMethod]
        //[ExpectedException(typeof(FileTypeNotSupported))]
        //public void FileTypeNotSupported()
        //{
        //    var outputFolder = CreateTemporaryFolder();
        //    var extractor = new OfficeExtractor.Extractor();
        //    extractor.SaveToFolder("TestFiles\\filetypenotsupported.txt", outputFolder);
        //}

        [TestMethod]
        public void FileIsCorrupt()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsCorrupt>(() => extractor.Extract("TestFiles\\A corrupt compound document.doc", outputFolder));
        }

        #region Microsoft Office Word tests
        [TestMethod]
        public void DocWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOC word document without embedded files.doc", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void DocWithoutAttachmentsButWithEmbeddings_All()
        {
            DocWithoutAttachmentsButWithEmbeddings(false, 7);
        }

        [TestMethod]
        public void DocWithoutAttachmentsButWithEmbeddings_AttachmentsOnly()
        {
            DocWithoutAttachmentsButWithEmbeddings(true, 0);
        }

        private void DocWithoutAttachmentsButWithEmbeddings(bool attachmentsOnly, int expectedCount)
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOC file without attachments but with embeddings.doc",
                outputFolder, attachmentsOnly: attachmentsOnly);
            Assert.AreEqual(expectedCount, files.Count);
        }

        [TestMethod]
        public void DocWith2EmbeddedImages()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOC word document with embedded images.doc", outputFolder);
            Assert.HasCount(2, files);
        }

        [TestMethod]
        public void DocWith7EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOC word document with 7 embedded files.doc", outputFolder);
            Assert.HasCount(7, files);
        }

        [TestMethod]
        public void DocWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\A DOC word document with password.doc", outputFolder));
        }

        [TestMethod]
        public void DocxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOCX word document without embedded files.docx", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOCX word document with 7 embedded files.docx", outputFolder);
            Assert.HasCount(7, files);
        }

        [TestMethod]
        public void DocxWithUNICODE_EmbeddedZipFile()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            // New for replace not exceptions
            Extractor.ReplaceDisallowedCharsInFileNames = true;
            var files = extractor.Extract("TestFiles\\A DOCX word document with embedding Zip with UNICODE Sign.docx",
                outputFolder);
            Assert.HasCount(1, files);
            Assert.EndsWith("unicode-___.zip", files[0]);
        }

        [TestMethod]
        public void DocxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\A DOCX word document with password.docx", outputFolder));
        }

        [TestMethod]
        public void DocWithDocumentOleObjectAttached()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOC word document with document ole object attached.doc",
                outputFolder);
            Assert.HasCount(1, files);
            Assert.AreEqual("attachment.pdf", Path.GetFileName(files[0]));
        }

        [TestMethod]
        public void DocWithDocumentOleObjectAttachedPathRemoved()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A DOC word document with document ole object attached path removed.doc", outputFolder);
            Assert.HasCount(1, files);
            Assert.AreEqual("Embedded object.pdf", Path.GetFileName(files[0]));
        }


        [TestMethod]
        public void DocxWithEmbeddedMathTypeObjectSuccessfulExtractsNothing()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\MathType 5 Object.docx", outputFolder);
            Assert.IsEmpty(files);
        }
        #endregion

        #region Microsoft Office Excel tests
        [TestMethod]
        public void XlsWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A XLS excel document without embedded files.xls", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void XlsWith2EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A XLS excel document with 2 embedded files.xls", outputFolder);
            Assert.HasCount(2, files);
        }

        [TestMethod]
        public void XlsWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\A XLS excel document with password.xls", outputFolder));
        }

        [TestMethod]
        public void XlsxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A XLSX excel document without embedded files.xlsx", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void XlsxWith2EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A XLSX excel document with 2 embedded files.xlsx", outputFolder);
            Assert.HasCount(2, files);
        }

        [TestMethod]
        public void XlsxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\A XLSX excel document with password.xlsx", outputFolder));
        }
        #endregion

        #region Microsoft Office PowerPoint tests
        [TestMethod]
        public void PptWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A PPT PowerPoint document without embedded files.ppt",
                outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void PptWith3EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A PPT powerpoint document with 3 embedded files.ppt",
                outputFolder);
            Assert.HasCount(3, files);
        }

        /// <summary>
        ///     Test failed in NET6 prior to adaptions in reading <see cref="System.IO.Compression.DeflateStream" />
        /// </summary>
        [TestMethod]
        public void PptWith1EmbeddedFile()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A PPT PowerPoint document with 1 embedded file.ppt",
                outputFolder);
            Assert.HasCount(1, files);
        }

        [TestMethod]
        public void PptWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\A PPT PowerPoint document with password.ppt", outputFolder));
        }

        [TestMethod]
        public void PptxWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A PPTX PowerPoint document without embedded files.pptx",
                outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void PptxWith3EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A PPTX powerpoint document with 3 embedded files.pptx",
                outputFolder);
            Assert.HasCount(3, files);
        }

        [TestMethod]
        public void PptxWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\A PPTX PowerPoint document with password.pptx", outputFolder));
        }

        [TestMethod]
        public void PptWithEmbeddedMicrosoftClipArtGalleryObject()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract(
                "TestFiles\\A PPT PowerPoint document containing a Microsoft ClipArt Gallery object.ppt", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void PptWithEmbeddedMsClipArtGalleryObject()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract(
                "TestFiles\\A PPT PowerPoint document containing a MS ClipArt Gallery object.ppt", outputFolder);
            Assert.IsEmpty(files);
        }
        #endregion

        #region Open Office Writer tests
        [TestMethod]
        public void OdtWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\An ODT document without embedded files.odt", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void OdtWith8EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\An ODT document with 8 embedded files.odt", outputFolder);
            Assert.HasCount(8, files);
        }

        [TestMethod]
        public void OdtWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\An ODT document with password.odt", outputFolder));
        }
        #endregion

        #region Open Office Calc tests
        [TestMethod]
        public void OdsWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\An ODS document without embedded files.ods", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void OdsWith2EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\An ODS document with 2 embedded files.ods", outputFolder);
            Assert.HasCount(2, files);
        }

        [TestMethod]
        public void OdsWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\An ODS document with password.ods", outputFolder));
        }
        #endregion

        #region Open Office Impress tests
        [TestMethod]
        public void OdpWithoutEmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\An ODP document without embedded files.odp", outputFolder);
            Assert.IsEmpty(files);
        }

        [TestMethod]
        public void OdpWith3EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\An ODP document with 3 embedded files.odp", outputFolder);
            Assert.HasCount(3, files);
        }

        [TestMethod]
        public void OdpWithPassword()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            Assert.ThrowsExactly<OEFileIsPasswordProtected>(() => extractor.Extract("TestFiles\\An ODP document with password.odp", outputFolder));
        }
        #endregion

        #region RTF tests
        [TestMethod]
        public void RtfWith11EmbeddedFiles()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A RTF document with 11 embedded files.rtf", outputFolder);
            Assert.HasCount(11, files);
        }

        [TestMethod]
        public void RtfWith3EmbeddedFilesAndNoSpaceDelimiters()
        {
            var outputFolder = CreateTemporaryFolder();
            var extractor = new Extractor();
            var files = extractor.Extract("TestFiles\\A RTF document with 3 embedded files and no space delimiters.rtf",
                outputFolder);
            Assert.HasCount(3, files);
        }
        #endregion

        #region Helper methods
        /// <summary>
        ///     Creates a new temporary folder and returns the path to it
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
            catch (Exception)
            {
                // Ignore
            }
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