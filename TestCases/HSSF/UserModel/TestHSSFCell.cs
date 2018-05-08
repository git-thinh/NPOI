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
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.UserModel;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record;
    using System.Text;

    /**
     * Tests various functionity having to do with Cell.  For instance support for
     * paticular datatypes, etc.
     * @author Andrew C. Oliver (andy at superlinksoftware dot com)
     * @author  Dan Sherman (dsherman at isisph.com)
     * @author Alex Jacoby (ajacoby at gmail.com)
     */
    [TestClass]
    public class TestHSSFCell
    {

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }
        private static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack(original);
        }
        [TestMethod]
        public void TestSetValues()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = book.CreateSheet("Test");
            Row row = sheet.CreateRow(0);

            Cell cell = row.CreateCell(0);

            cell.SetCellValue(1.2);
            Assert.AreEqual(1.2, cell.NumericCellValue, 0.0001);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, cell.CellType);

            cell.SetCellValue(false);
            Assert.AreEqual(false, cell.BooleanCellValue);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BOOLEAN, cell.CellType);

            cell.SetCellValue(new HSSFRichTextString("Foo"));
            Assert.AreEqual("Foo", cell.RichStringCellValue.String);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, cell.CellType);

            cell.SetCellValue(new HSSFRichTextString("345"));
            Assert.AreEqual("345", cell.RichStringCellValue.String);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, cell.CellType);
        }

        /**
         * Test that Boolean and Error types (BoolErrRecord) are supported properly.
         */
        //[TestMethod]
        //public void TestBoolErr()
        //{

        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    NPOI.SS.UserModel.Sheet s = wb.CreateSheet("TestSheet1");
        //    Row r = null;
        //    Cell c = null;
        //    r = s.CreateRow(0);
        //    c = r.CreateCell(1);
        //    //c.SetCellType(NPOI.SS.UserModel.CellType.BOOLEAN);
        //    c.SetCellValue(true);

        //    c = r.CreateCell(2);
        //    //c.SetCellType(NPOI.SS.UserModel.CellType.BOOLEAN);
        //    c.SetCellValue(false);

        //    r = s.CreateRow(1);
        //    c = r.CreateCell(1);
        //    //c.SetCellType(NPOI.SS.UserModel.CellType.ERROR);
        //    c.SetCellErrorValue((byte)0);

        //    c = r.CreateCell(2);
        //    //c.SetCellType(NPOI.SS.UserModel.CellType.ERROR);
        //    c.SetCellErrorValue((byte)7);

        //    wb = WriteOutAndReadBack(wb);
        //    s = wb.GetSheetAt(0);
        //    r = s.GetRow(0);
        //    c = r.GetCell(1);
        //    Assert.IsTrue(c.BooleanCellValue, "boolean value 0,1 = true");
        //    c = r.GetCell(2);
        //    Assert.IsTrue(c.BooleanCellValue == false, "boolean value 0,2 = false");
        //    r = s.GetRow(1);
        //    c = r.GetCell(1);
        //    Assert.IsTrue(c.ErrorCellValue == 0, "boolean value 0,1 = 0");
        //    c = r.GetCell(2);
        //    Assert.IsTrue(c.ErrorCellValue == 7, "boolean value 0,2 = 7");
        //}

        /**
         * Checks that the recognition of files using 1904 date windowing
         *  is working properly. Conversion of the date is also an issue,
         *  but there's a separate unit Test for that.
         */
        [TestMethod]
        public void TestDateWindowingRead()
        {
            DateTime date = new DateTime(2000, 1, 1);

            // first Check a file with 1900 Date Windowing
            HSSFWorkbook workbook = OpenSample("1900DateWindowing.xls");
            NPOI.SS.UserModel.Sheet sheet = workbook.GetSheetAt(0);

            Assert.AreEqual(date, sheet.GetRow(0).GetCell(0).DateCellValue,
                               "Date from file using 1900 Date Windowing");

            // now Check a file with 1904 Date Windowing
            workbook = OpenSample("1904DateWindowing.xls");
            sheet = workbook.GetSheetAt(0);

            Assert.AreEqual(date, sheet.GetRow(0).GetCell(0).DateCellValue,
                             "Date from file using 1904 Date Windowing");
        }

        /**
         * Checks that dates are properly written to both types of files:
         * those with 1900 and 1904 date windowing.  Note that if the
         * previous Test ({@link #TestDateWindowingRead}) Assert.Fails, the
         * results of this Test are meaningless.
         */
        [TestMethod]
        public void TestDateWindowingWrite()
        {
            DateTime date = new DateTime(2000, 1, 1);

            // first Check a file with 1900 Date Windowing
            HSSFWorkbook wb;
            wb = OpenSample("1900DateWindowing.xls");

            SetCell(wb, 0, 1, date);
            wb = WriteOutAndReadBack(wb);

            Assert.AreEqual(date,
                            ReadCell(wb, 0, 1), "Date from file using 1900 Date Windowing");

            // now Check a file with 1904 Date Windowing
            wb = OpenSample("1904DateWindowing.xls");
            SetCell(wb, 0, 1, date);
            wb = WriteOutAndReadBack(wb);
            Assert.AreEqual(date,
                            ReadCell(wb, 0, 1), "Date from file using 1900 Date Windowing");
        }

        /**
 * Test for small bug observable around r736460 (prior to version 3.5).  POI fails to remove
 * the {@link StringRecord} following the {@link FormulaRecord} after the result type had been 
 * changed to number/boolean/error.  Excel silently ignores the extra record, but some POI
 * versions (prior to bug 46213 / r717883) crash instead.
 */
        [TestMethod]
        public void TestCachedTypeChange()
        {
            HSSFSheet sheet = (HSSFSheet)new HSSFWorkbook().CreateSheet("Sheet1");
            HSSFCell cell = (HSSFCell)sheet.CreateRow(0).CreateCell(0);
            cell.CellFormula = ("A1");
            cell.SetCellValue("abc");
            ConfirmStringRecord(sheet, true);
            cell.SetCellValue(123);
            NPOI.HSSF.Record.Record[] recs = RecordInspector.GetRecords(sheet, 0);
            if (recs.Length == 28 && recs[23] is StringRecord)
            {
                throw new AssertFailedException("Identified bug - leftover StringRecord");
            }
            ConfirmStringRecord(sheet, false);

            // string to error code
            cell.SetCellValue("abc");
            ConfirmStringRecord(sheet, true);
            cell.SetCellErrorValue((byte)ErrorConstants.ERROR_REF);
            ConfirmStringRecord(sheet, false);

            // string to boolean
            cell.SetCellValue("abc");
            ConfirmStringRecord(sheet, true);
            cell.SetCellValue(false);
            ConfirmStringRecord(sheet, false);
        }

        private static void ConfirmStringRecord(HSSFSheet sheet, bool isPresent)
        {
            Record[] recs = RecordInspector.GetRecords(sheet, 0);
            Assert.AreEqual(isPresent ? 31 : 30, recs.Length);
            int index = 24;
            Record fr = recs[index++];
            Assert.AreEqual(typeof(FormulaRecord), fr.GetType());
            if (isPresent)
            {
                Assert.AreEqual(typeof(StringRecord), recs[index++].GetType());
            }
            else
            {
                Assert.IsFalse(typeof(StringRecord) == recs[index].GetType());
            }
            Record dbcr = recs[index++];
            Assert.AreEqual(typeof(DBCellRecord), dbcr.GetType());
        }

        /**
         *  The maximum length of cell contents (text) is 32,767 characters.
         */
        [TestMethod]
        public void TestMaxTextLength()
        {
            HSSFSheet sheet = (HSSFSheet)new HSSFWorkbook().CreateSheet();
            HSSFCell cell = (HSSFCell)sheet.CreateRow(0).CreateCell(0);

            int maxlen = NPOI.SS.SpreadsheetVersion.EXCEL97.MaxTextLength;
            Assert.AreEqual(32767, maxlen);

            StringBuilder b = new StringBuilder();

            // 32767 is okay
            for (int i = 0; i < maxlen; i++)
            {
                b.Append("X");
            }
            cell.SetCellValue(b.ToString());

            b.Append("X");
            // 32768 produces an invalid XLS file
            try
            {
                cell.SetCellValue(b.ToString());
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The maximum length of cell contents (text) is 32,767 characters", e.Message);
            }
        }

        private static void SetCell(HSSFWorkbook workbook, int rowIdx, int colIdx, DateTime date)
        {
            NPOI.SS.UserModel.Sheet sheet = workbook.GetSheetAt(0);
            Row row = sheet.GetRow(rowIdx);
            Cell cell = row.GetCell(colIdx);

            if (cell == null)
            {
                cell = row.CreateCell(colIdx);
            }
            cell.SetCellValue(date);
        }

        private static DateTime ReadCell(HSSFWorkbook workbook, int rowIdx, int colIdx)
        {
            NPOI.SS.UserModel.Sheet sheet = workbook.GetSheetAt(0);
            Row row = sheet.GetRow(rowIdx);
            Cell cell = row.GetCell(colIdx);
            return cell.DateCellValue;
        }

        /**
         * Tests that the active cell can be correctly read and set
         */
        [TestMethod]
        public void TestActiveCell()
        {
            //read in sample
            HSSFWorkbook book = OpenSample("Simple.xls");

            //Check initial position
            HSSFSheet umSheet = (HSSFSheet)book.GetSheetAt(0);
            NPOI.HSSF.Model.Sheet s = umSheet.Sheet;
            Assert.AreEqual(0, s.ActiveCellCol, "Initial active cell should be in col 0");
            Assert.AreEqual(1, s.ActiveCellRow, "Initial active cell should be on row 1");

            //modify position through Cell
            Cell cell = umSheet.CreateRow(3).CreateCell(2);
            cell.SetAsActiveCell();
            Assert.AreEqual(2, s.ActiveCellCol, "After modify, active cell should be in col 2");
            Assert.AreEqual(3, s.ActiveCellRow, "After modify, active cell should be on row 3");

            //Write book to temp file; read and Verify that position is serialized
            book = WriteOutAndReadBack(book);

            umSheet = (HSSFSheet)book.GetSheetAt(0);
            s = umSheet.Sheet;

            Assert.AreEqual(2, s.ActiveCellCol, "After serialize, active cell should be in col 2");
            Assert.AreEqual(3, s.ActiveCellRow, "After serialize, active cell should be on row 3");
        }

        /**
         * Test that Cell Styles being applied to formulas remain intact
         */
        [TestMethod]
        public void TestFormulaStyle()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("TestSheet1");
            Row r = null;
            Cell c = null;
            NPOI.SS.UserModel.CellStyle cs = wb.CreateCellStyle();
            Font f = wb.CreateFont();
            f.FontHeightInPoints = ((short)20);
            f.Color = (HSSFColor.RED.index);
            f.Boldweight = (short)FontBoldWeight.BOLD;
            f.FontName = ("Arial Unicode MS");
            cs.FillBackgroundColor = ((short)3);
            cs.SetFont(f);
            cs.BorderTop = CellBorderType.THIN;
            cs.BorderRight = CellBorderType.THIN;
            cs.BorderLeft = CellBorderType.THIN;
            cs.BorderBottom = CellBorderType.THIN;

            r = s.CreateRow(0);
            c = r.CreateCell(0);
            c.CellStyle = (cs);
            c.CellFormula = ("2*3");

            wb = WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.IsTrue((c.CellType == NPOI.SS.UserModel.CellType.FORMULA), "Formula Cell at 0,0");
            cs = c.CellStyle;

            Assert.IsNotNull(cs, "Formula Cell Style");
            Assert.AreEqual(cs.FontIndex, f.Index, "Font Index Matches");
            Assert.AreEqual((short)cs.BorderTop, (short)1, "Top Border");
            Assert.AreEqual((short)cs.BorderLeft, (short)1, "Left Border");
            Assert.AreEqual((short)cs.BorderRight, (short)1, "Right Border");
            Assert.AreEqual((short)cs.BorderBottom, (short)1, "Bottom Border");
        }

        /**
         * Test reading hyperlinks
         */
        [TestMethod]
        public void TestWithHyperlink()
        {

            HSSFWorkbook wb = OpenSample("WithHyperlink.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);
            Cell cell = sheet.GetRow(4).GetCell(0);
            Hyperlink link = cell.Hyperlink;
            Assert.IsNotNull(link);

            Assert.AreEqual("Foo", link.Label);
            Assert.AreEqual(link.Address, "http://poi.apache.org/");
            Assert.AreEqual(4, link.FirstRow);
            Assert.AreEqual(0, link.FirstColumn);
        }

        /**
         * Test reading hyperlinks
         */
        [TestMethod]
        public void TestWithTwoHyperlinks()
        {

            HSSFWorkbook wb = OpenSample("WithTwoHyperLinks.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);

            Cell cell1 = sheet.GetRow(4).GetCell(0);
            Hyperlink link1 = cell1.Hyperlink;
            Assert.IsNotNull(link1);
            Assert.AreEqual("Foo", link1.Label);
            Assert.AreEqual("http://poi.apache.org/", link1.Address);
            Assert.AreEqual(4, link1.FirstRow);
            Assert.AreEqual(0, link1.FirstColumn);

            Cell cell2 = sheet.GetRow(8).GetCell(1);
            Hyperlink link2 = cell2.Hyperlink;
            Assert.IsNotNull(link2);
            Assert.AreEqual("Bar", link2.Label);
            Assert.AreEqual("http://poi.apache.org/hssf/", link2.Address);
            Assert.AreEqual(8, link2.FirstRow);
            Assert.AreEqual(1, link2.FirstColumn);
        }

        /*Tests the ToString() method of Cell*/
        [TestMethod]
        public void TestToString()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Row r = wb.CreateSheet("Sheet1").CreateRow(0);

            r.CreateCell(0).SetCellValue(true);
            r.CreateCell(1).SetCellValue(1.5);
            r.CreateCell(2).SetCellValue(new HSSFRichTextString("Astring"));
            r.CreateCell(3).SetCellErrorValue((byte)HSSFErrorConstants.ERROR_DIV_0);
            r.CreateCell(4).CellFormula = ("A1+B1");

            Assert.AreEqual("TRUE", r.GetCell(0).ToString(), "Boolean");
            Assert.AreEqual("1.5", r.GetCell(1).ToString(), "Numeric");
            Assert.AreEqual("Astring", r.GetCell(2).ToString(), "String");
            Assert.AreEqual("#DIV/0!", r.GetCell(3).ToString(), "Error");
            Assert.AreEqual("A1+B1", r.GetCell(4).ToString(), "Formula");

            //Write out the file, read it in, and then Check cell values
            wb = WriteOutAndReadBack(wb);

            r = wb.GetSheetAt(0).GetRow(0);
            Assert.AreEqual("TRUE", r.GetCell(0).ToString(), "Boolean");
            Assert.AreEqual("1.5", r.GetCell(1).ToString(), "Numeric");
            Assert.AreEqual("Astring", r.GetCell(2).ToString(), "String");
            Assert.AreEqual("#DIV/0!", r.GetCell(3).ToString(), "Error");
            Assert.AreEqual("A1+B1", r.GetCell(4).ToString(), "Formula");
        }
        [TestMethod]
        public void TestSetStringInFormulaCell_bug44606()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Cell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            cell.CellFormula = ("B1&C1");
            try
            {
                cell.SetCellValue(new HSSFRichTextString("hello"));
            }
            catch (InvalidCastException e)
            {
                throw new AssertFailedException("Identified bug 44606");
            }
        }
        [TestMethod]
        public void TestHSSFCellToStringWithDataFormat()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Cell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            cell.SetCellValue(new DateTime(2009, 8, 20));
            NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy");
            cell.CellStyle = cellStyle;
            Assert.AreEqual("8/20/09", cell.ToString());

            NPOI.SS.UserModel.CellStyle cellStyle2 = wb.CreateCellStyle();
            DataFormat format = wb.CreateDataFormat();
            cellStyle2.DataFormat = format.GetFormat("YYYY-mm/dd");
            cell.CellStyle = cellStyle2;
            Assert.AreEqual("2009-08/20", cell.ToString());
        }
        [TestMethod]
        public void TestGetDataFormatUniqueIndex()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            DataFormat format = wb.CreateDataFormat();
            short formatidx1 = format.GetFormat("YYYY-mm/dd");
            short formatidx2 = format.GetFormat("YYYY-mm/dd");
            Assert.AreEqual(formatidx1, formatidx2);
            short formatidx3 = format.GetFormat("000000.000");
            Assert.AreNotEqual(formatidx1, formatidx3);
        }
        /**
         * Test to ensure we can only assign cell styles that belong
         *  to our workbook, and not those from other workbooks.
         */
        [TestMethod]
        public void TestCellStyleWorkbookMatch()
        {
            HSSFWorkbook wbA = new HSSFWorkbook();
            HSSFWorkbook wbB = new HSSFWorkbook();

            HSSFCellStyle styA = (HSSFCellStyle)wbA.CreateCellStyle();
            HSSFCellStyle styB = (HSSFCellStyle)wbB.CreateCellStyle();

            styA.VerifyBelongsToWorkbook(wbA);
            styB.VerifyBelongsToWorkbook(wbB);
            try
            {
                styA.VerifyBelongsToWorkbook(wbB);
                Assert.Fail();
            }
            catch (ArgumentException e) { }
            try
            {
                styB.VerifyBelongsToWorkbook(wbA);
                Assert.Fail();
            }
            catch (ArgumentException e) { }

            Cell cellA = wbA.CreateSheet().CreateRow(0).CreateCell(0);
            Cell cellB = wbB.CreateSheet().CreateRow(0).CreateCell(0);

            cellA.CellStyle = (styA);
            cellB.CellStyle = (styB);
            try
            {
                cellA.CellStyle = (styB);
                Assert.Fail();
            }
            catch (ArgumentException) { }
            try
            {
                cellB.CellStyle = (styA);
                Assert.Fail();
            }
            catch (ArgumentException) { }
        }

    }

}