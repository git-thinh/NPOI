/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.UserModel
{
    using System;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.UserModel;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;

    /**
     * Tests row shifting capabilities.
     *
     *
     * @author Shawn Laubach (slaubach at apache dot com)
     */
    [TestClass]
    public class TestHSSFHeaderFooter
    {

        /**
         * Tests that get header retreives the proper values.
         *
         * @author Shawn Laubach (slaubach at apache dot org)
         */
        [TestMethod]
        public void TestRetrieveCorrectHeader()
        {
            // Read initial file in
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("EmbeddedChartHeaderTest.xls");
            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Header head = s.Header;

            Assert.AreEqual("Top Left", head.Left);
            Assert.AreEqual("Top Center", head.Center);
            Assert.AreEqual("Top Right", head.Right);
        }
        [TestMethod]
        public void TestSpecialChars()
        {
            Assert.AreEqual("&U", HSSFHeader.StartUnderline);
            Assert.AreEqual("&U", HSSFHeader.EndUnderline);
            Assert.AreEqual("&P", HSSFHeader.Page);

            Assert.AreEqual("&22", HSSFFooter.FontSize((short)22));
            Assert.AreEqual("&\"Arial,bold\"", HSSFFooter.Font("Arial", "bold"));
        }
        [TestMethod]
        public void TestStripFields()
        {
            String simple = "I am a Test header";
            String withPage = "I am a&P Test header";
            String withLots = "I&A am&N a&P Test&T header&U";
            String withFont = "I&22 am a&\"Arial,bold\" Test header";
            String withOtherAnds = "I am a&P Test header&&";
            String withOtherAnds2 = "I am a&P Test header&a&b";

            Assert.AreEqual(simple, HSSFHeader.StripFields(simple));
            Assert.AreEqual(simple, HSSFHeader.StripFields(withPage));
            Assert.AreEqual(simple, HSSFHeader.StripFields(withLots));
            Assert.AreEqual(simple, HSSFHeader.StripFields(withFont));
            Assert.AreEqual(simple + "&&", HSSFHeader.StripFields(withOtherAnds));
            Assert.AreEqual(simple + "&a&b", HSSFHeader.StripFields(withOtherAnds2));

            // Now Test the default Strip flag
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("EmbeddedChartHeaderTest.xls");
            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Header head = s.Header;

            Assert.AreEqual("Top Left", head.Left);
            Assert.AreEqual("Top Center", head.Center);
            Assert.AreEqual("Top Right", head.Right);

            head.Left = ("Top &P&F&D Left");
            Assert.AreEqual("Top &P&F&D Left", head.Left);
            
            Assert.AreEqual("Top  Left", NPOI.HSSF.UserModel.HeaderFooter.StripFields(head.Left));
            // Now even more complex
            head.Center = ("HEADER TEXT &P&N&D&T&Z&F&F&A&G&X END");
            Assert.AreEqual("HEADER TEXT  END", NPOI.HSSF.UserModel.HeaderFooter.StripFields(head.Center));
        }

        /**
         * Tests that get header retreives the proper values.
         *
         * @author Shawn Laubach (slaubach at apache dot org)
         */
        [TestMethod]
        public void TestRetrieveCorrectFooter()
        {
            // Read initial file in
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("EmbeddedChartHeaderTest.xls");
            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Footer foot = s.Footer;

            Assert.AreEqual("Bottom Left", foot.Left);
            Assert.AreEqual("Bottom Center", foot.Center);
            Assert.AreEqual("Bottom Right", foot.Right);
        }

        /**
         * Testcase for Bug 17039 HSSFHeader  doesnot support DBCS 
         */
        [TestMethod]
        public void TestHeaderHas16bitCharacter()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = b.CreateSheet("Test");
            Header h = s.Header;
            h.Left = ("\u0391");
            h.Center = ("\u0392");
            h.Right = ("\u0393");

            HSSFWorkbook b2 = HSSFTestDataSamples.WriteOutAndReadBack(b);
            Header h2 = b2.GetSheet("Test").Header;

            Assert.AreEqual(h2.Left, "\u0391");
            Assert.AreEqual(h2.Center, "\u0392");
            Assert.AreEqual(h2.Right, "\u0393");
        }

        /**
         * Testcase for Bug 17039 HSSFFooter doesnot support DBCS 
         */
        [TestMethod]
        public void TestFooterHas16bitCharacter()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = b.CreateSheet("Test");
            Footer f = s.Footer;
            f.Left = ("\u0391");
            f.Center = ("\u0392");
            f.Right = ("\u0393");

            HSSFWorkbook b2 = HSSFTestDataSamples.WriteOutAndReadBack(b);
            Footer f2 = b2.GetSheet("Test").Footer;

            Assert.AreEqual(f2.Left, "\u0391");
            Assert.AreEqual(f2.Center, "\u0392");
            Assert.AreEqual(f2.Right, "\u0393");
        }
        [TestMethod]
        public void TestReadDBCSHeaderFooter()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DBCSHeader.xls");
            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Header h = s.Header;
            Assert.AreEqual(h.Left, "\u090f\u0915","Header Left ");
            Assert.AreEqual(h.Center, "\u0939\u094b\u0917\u093e", "Header Center ");
            Assert.AreEqual(h.Right, "\u091c\u093e", "Header Right ");

            Footer f = s.Footer;
            Assert.AreEqual(f.Left, "\u091c\u093e", "Footer Left ");
            Assert.AreEqual(f.Center, "\u091c\u093e", "Footer Center ");
            Assert.AreEqual(f.Right, "\u091c\u093e", "Footer Right ");
        }
    }

}