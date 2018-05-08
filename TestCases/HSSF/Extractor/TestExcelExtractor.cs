/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Extractor
{
    using System;
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.Extractor;
    using TestCases.HSSF;


    using Microsoft.VisualStudio.TestTools.UnitTesting;


    [TestClass]
    public class TestExcelExtractor
    {

        private static ExcelExtractor CreateExtractor(String sampleFileName)
        {

            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);

            try
            {
                return new ExcelExtractor(new POIFSFileSystem(is1));
            }
            catch (IOException)
            {
                throw;
            }
        }

        [TestMethod]
        public void TestSimple()
        {

            ExcelExtractor extractor = CreateExtractor("Simple.xls");

            Assert.AreEqual("Sheet1\nreplaceMe\nSheet2\nSheet3\n", extractor.Text);

            // Now turn off sheet names
            extractor.IncludeSheetNames=false;
            Assert.AreEqual("replaceMe\n", extractor.Text);
        }
        [TestMethod]
        public void TestNumericFormula()
        {

            ExcelExtractor extractor = CreateExtractor("sumifformula.xls");

            string expected = "Sheet1\n" +
                    "1000\t1\t5\n" +
                    "2000\t2\t\n" +
                    "3000\t3\t\n" +
                    "4000\t4\t\n" +
                    "5000\t5\t\n" +
                    "Sheet2\nSheet3\n";
            Assert.AreEqual(
                    expected,
                    extractor.Text
            );

            extractor.FormulasNotResults=(true);

            Assert.AreEqual(
                "Sheet1\n" +
                "1000\t1\tSUMIF(A1:A5,\">4000\",B1:B5)\n" +
                "2000\t2\t\n" +
                "3000\t3\t\n" +
                "4000\t4\t\n" +
                "5000\t5\t\n" +
                "Sheet2\nSheet3\n", 
                    extractor.Text
            );
        }
        [TestMethod]
        public void TestwithContinueRecords()
        {

            ExcelExtractor extractor = CreateExtractor("StringContinueRecords.xls");

            //extractor.Text;

            // Has masses of text
            // Until we fixed bug #41064, this would've
            //   failed by now
            Assert.IsTrue(extractor.Text.Length > 40960);
        }
        [TestMethod]
        public void TestStringConcat()
        {

            ExcelExtractor extractor = CreateExtractor("SimpleWithFormula.xls");

            // Comes out as NaN if treated as a number
            // And as XYZ if treated as a string
            Assert.AreEqual("Sheet1\nreplaceme\nreplaceme\nreplacemereplaceme\nSheet2\nSheet3\n", extractor.Text);

            extractor.FormulasNotResults = (true);

            Assert.AreEqual("Sheet1\nreplaceme\nreplaceme\nCONCATENATE(A1,A2)\nSheet2\nSheet3\n", extractor.Text);
        }
        [TestMethod]
        public void TestStringFormula()
        {

            ExcelExtractor extractor = CreateExtractor("StringFormulas.xls");

            // Comes out as NaN if treated as a number
            // And as XYZ if treated as a string
            Assert.AreEqual("Sheet1\nXYZ\nSheet2\nSheet3\n", extractor.Text);

            extractor.FormulasNotResults = (true);

            Assert.AreEqual("Sheet1\nUPPER(\"xyz\")\nSheet2\nSheet3\n", extractor.Text);
        }

        [TestMethod]
        public void TestEventExtractor()
        {
            EventBasedExcelExtractor extractor;

            // First up, a simple file with string
            //  based formulas in it
            extractor = new EventBasedExcelExtractor(
                    new POIFSFileSystem(
                            HSSFTestDataSamples.OpenSampleFileStream("SimpleWithFormula.xls")
                    )
            );
            extractor.IncludeSheetNames = (true);

            String text = extractor.Text;
            Assert.AreEqual("Sheet1\nreplaceme\nreplaceme\nreplacemereplaceme\nSheet2\nSheet3\n", text);

            extractor.IncludeSheetNames = (false);
            extractor.FormulasNotResults = (true);

            text = extractor.Text;
            Assert.AreEqual("replaceme\nreplaceme\nCONCATENATE(A1,A2)\n", text);


            // Now, a slightly longer file with numeric formulas
            extractor = new EventBasedExcelExtractor(
                    new POIFSFileSystem(
                            HSSFTestDataSamples.OpenSampleFileStream("sumifformula.xls")
                    )
            );
            extractor.IncludeSheetNames = (false);
            extractor.FormulasNotResults = (true);

            text = extractor.Text;
            Assert.AreEqual(
                    "1000\t1\tSUMIF(A1:A5,\">4000\",B1:B5)\n" +
                    "2000\t2\n" +
                    "3000\t3\n" +
                    "4000\t4\n" +
                    "5000\t5\n",
                    text
            );
        }
        [TestMethod]
        public void TestWithComments()
        {
            ExcelExtractor extractor = CreateExtractor("SimpleWithComments.xls");
            extractor.IncludeSheetNames = (false);

            // Check without comments
            Assert.AreEqual(
                    "1\tone\n" +
                    "2\ttwo\n" +
                    "3\tthree\n",
                    extractor.Text
            );

            // Now with
            extractor.IncludeCellComments = (true);
            Assert.AreEqual(
                    "1\tone Comment by Yegor Kozlov: Yegor Kozlov: first cell\n" +
                    "2\ttwo Comment by Yegor Kozlov: Yegor Kozlov: second cell\n" +
                    "3\tthree Comment by Yegor Kozlov: Yegor Kozlov: third cell\n",
                    extractor.Text
            );
        }
        [TestMethod]
        public void TestWithBlank()
        {
            ExcelExtractor extractor = CreateExtractor("MissingBits.xls");
            String def = extractor.Text;
            extractor.IncludeBlankCells = (true);
            String padded = extractor.Text;

            Assert.IsTrue(def.StartsWith(
                "Sheet1\n" +
                "&[TAB]\t\n" +
                "Hello\t\n" +
                "11\t23\t\n"
            ));

            Assert.IsTrue(padded.StartsWith(
                "Sheet1\n" +
                "&[TAB]\t\n" +
                "Hello\t\t\t\t\t\t\t\t\t\t\t\n" +
                "11\t\t\t23\t\t\t\t\t\t\t\t\n"
            ));
        }


        /**
         * Embded in a non-excel file
         */
        [TestMethod]
        public void TestWithEmbeded()
        {
            POIFSFileSystem fs = new POIFSFileSystem(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("word_with_embeded.doc")
            );

            DirectoryNode objPool = (DirectoryNode)fs.Root.GetEntry("ObjectPool");
            DirectoryNode dirA = (DirectoryNode)objPool.GetEntry("_1269427460");
            DirectoryNode dirB = (DirectoryNode)objPool.GetEntry("_1269427461");

            HSSFWorkbook wbA = new HSSFWorkbook(dirA, fs, true);
            HSSFWorkbook wbB = new HSSFWorkbook(dirB, fs, true);

            ExcelExtractor exA = new ExcelExtractor(wbA);
            ExcelExtractor exB = new ExcelExtractor(wbB);

            Assert.AreEqual("Sheet1\nTest excel file\nThis is the first file\nSheet2\nSheet3\n",
                    exA.Text);
            Assert.AreEqual("Sample Excel", exA.SummaryInformation.Title);

            Assert.AreEqual("Sheet1\nAnother excel file\nThis is the second file\nSheet2\nSheet3\n",
                    exB.Text);
            Assert.AreEqual("Sample Excel 2", exB.SummaryInformation.Title);
        }

        /**
         * Excel embeded in excel
         */
        [TestMethod]
        public void TestWithEmbededInOwn()
        {
            // TODO - encapsulate sys prop 'POIFS.Testdata.path' similar to HSSFTestDataSamples
            String pdirname = System.Configuration.ConfigurationSettings.AppSettings["POIFS.Testdata.path"];
            String filename = pdirname + "/excel_with_embeded.xls";
            POIFSFileSystem fs = new POIFSFileSystem(
                    new FileStream(filename, FileMode.Open, FileAccess.Read)
            );

            DirectoryNode dirA = (DirectoryNode)fs.Root.GetEntry("MBD0000A3B5");
            DirectoryNode dirB = (DirectoryNode)fs.Root.GetEntry("MBD0000A3B4");

            HSSFWorkbook wbA = new HSSFWorkbook(dirA, fs, true);
            HSSFWorkbook wbB = new HSSFWorkbook(dirB, fs, true);

            ExcelExtractor exA = new ExcelExtractor(wbA);
            ExcelExtractor exB = new ExcelExtractor(wbB);

            Assert.AreEqual("Sheet1\nTest excel file\nThis is the first file\nSheet2\nSheet3\n",
                    exA.Text);
            Assert.AreEqual("Sample Excel", exA.SummaryInformation.Title);

            Assert.AreEqual("Sheet1\nAnother excel file\nThis is the second file\nSheet2\nSheet3\n",
                    exB.Text);
            Assert.AreEqual("Sample Excel 2", exB.SummaryInformation.Title);

            // And the base file too
            ExcelExtractor ex = new ExcelExtractor(fs);
            Assert.AreEqual("Sheet1\nI have lots of embeded files in me\nSheet2\nSheet3\n",
                    ex.Text);
            Assert.AreEqual("Excel With Embeded", ex.SummaryInformation.Title);
        }

        /**
         * Test that we get text from headers and footers
         */
        [TestMethod]
        public void Test45538()
        {
            String[] files = {
			    "45538_classic_Footer.xls", "45538_form_Footer.xls",    
			    "45538_classic_Header.xls", "45538_form_Header.xls"
		    };
            for (int i = 0; i < files.Length; i++)
            {
                ExcelExtractor extractor = CreateExtractor(files[i]);
                String text = extractor.Text;
                Assert.IsTrue(text.IndexOf("testdoc") >= 0, "Unable to find expected word in text\n" + text);
                Assert.IsTrue(text.IndexOf("test phrase") >= 0, "Unable to find expected word in text\n" + text);
            }
        }
    }
}