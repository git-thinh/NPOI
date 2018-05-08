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

namespace NPOI.SS.UserModel
{
    using System;
    using NPOI.SS;
    using NPOI.SS.Util;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.SS;
    using System.Collections;


    /**
     * Common superclass for testing {@link NPOI.XSSF.usermodel.XSSFCell}  and
     * {@link NPOI.HSSF.usermodel.HSSFCell}
     */
    public abstract class BaseTestSheet
    {

        /**
         * @return an object that provides test data in HSSF / XSSF specific way
         */
        protected abstract ITestDataProvider GetTestDataProvider();

        [TestMethod]
        public void TestCreateRow()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet();
            Assert.AreEqual(0, sheet.PhysicalNumberOfRows);

            //Test that we get null for undefined rownumber
            Assert.IsNull(sheet.GetRow(1));

            // Test row creation with consecutive indexes
            Row row1 = sheet.CreateRow(0);
            Row row2 = sheet.CreateRow(1);
            Assert.AreEqual(0, row1.RowNum);
            Assert.AreEqual(1, row2.RowNum);
            IEnumerator it = sheet.GetRowEnumerator();
            Assert.IsTrue(it.MoveNext());
            Assert.AreSame(row1, it.Current);
            Assert.IsTrue(it.MoveNext());
            Assert.AreSame(row2, it.Current);
            Assert.AreEqual(1, sheet.LastRowNum);

            // Test row creation with non consecutive index
            Row row101 = sheet.CreateRow(100);
            Assert.IsNotNull(row101);
            Assert.AreEqual(100, sheet.LastRowNum);
            Assert.AreEqual(3, sheet.PhysicalNumberOfRows);

            // Test overwriting an existing row
            Row row2_ovrewritten = sheet.CreateRow(1);
            Cell cell = row2_ovrewritten.CreateCell(0);
            cell.SetCellValue(100);
            IEnumerator it2 = sheet.GetRowEnumerator();
            Assert.IsTrue(it2.MoveNext());
            Assert.AreSame(row1, it2.Current);
            Assert.IsTrue(it2.MoveNext());
            Row row2_ovrewritten_ref = (Row)it2.Current;
            Assert.AreSame(row2_ovrewritten, row2_ovrewritten_ref);
            Assert.AreEqual(100.0, row2_ovrewritten_ref.GetCell(0).NumericCellValue, 0.0);
        }

        [TestMethod]
        public void TestRemoveRow()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet1 = workbook.CreateSheet();
            Assert.AreEqual(0, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(0, sheet1.FirstRowNum);
            Assert.AreEqual(0, sheet1.LastRowNum);

            Row row0 = sheet1.CreateRow(0);
            Assert.AreEqual(1, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(0, sheet1.FirstRowNum);
            Assert.AreEqual(0, sheet1.LastRowNum);
            sheet1.RemoveRow(row0);
            Assert.AreEqual(0, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(0, sheet1.FirstRowNum);
            Assert.AreEqual(0, sheet1.LastRowNum);

            sheet1.CreateRow(1);
            Row row2 = sheet1.CreateRow(2);
            Assert.AreEqual(2, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(1, sheet1.FirstRowNum);
            Assert.AreEqual(2, sheet1.LastRowNum);

            Assert.IsNotNull(sheet1.GetRow(1));
            Assert.IsNotNull(sheet1.GetRow(2));
            sheet1.RemoveRow(row2);
            Assert.IsNotNull(sheet1.GetRow(1));
            Assert.IsNull(sheet1.GetRow(2));
            Assert.AreEqual(1, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(1, sheet1.FirstRowNum);
            Assert.AreEqual(1, sheet1.LastRowNum);

            Row row3 = sheet1.CreateRow(3);
            Sheet sheet2 = workbook.CreateSheet();
            try
            {
                sheet2.RemoveRow(row3);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Specified row does not belong to this sheet", e.Message);
            }
        }
        [TestMethod]
        public void TestCloneSheet()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            CreationHelper factory = workbook.GetCreationHelper();
            Sheet sheet = workbook.CreateSheet("Test Clone");
            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            Cell cell2 = row.CreateCell(1);
            cell.SetCellValue(factory.CreateRichTextString("Clone_test"));
            cell2.CellFormula = "SIN(1)";

            Sheet ClonedSheet = workbook.CloneSheet(0);
            Row ClonedRow = ClonedSheet.GetRow(0);

            //Check for a good clone
            Assert.AreEqual(ClonedRow.GetCell(0).RichStringCellValue.String, "Clone_test");

            //Check that the cells are not somehow linked
            cell.SetCellValue(factory.CreateRichTextString("Difference Check"));
            cell2.CellFormula = "cos(2)";
            if ("Difference Check".Equals(ClonedRow.GetCell(0).RichStringCellValue.String))
            {
                Assert.Fail("string cell not properly Cloned");
            }
            if ("COS(2)".Equals(ClonedRow.GetCell(1).CellFormula))
            {
                Assert.Fail("formula cell not properly Cloned");
            }
            Assert.AreEqual(ClonedRow.GetCell(0).RichStringCellValue.String, "Clone_test");
            Assert.AreEqual(ClonedRow.GetCell(1).CellFormula, "SIN(1)");
        }

        /** tests that the sheet name for multiple Clones of the same sheet is unique
         * BUG 37416
         */
        [TestMethod]
        public void TestCloneSheetMultipleTimes()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            CreationHelper factory = workbook.GetCreationHelper();
            Sheet sheet = workbook.CreateSheet("Test Clone");
            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            cell.SetCellValue(factory.CreateRichTextString("Clone_test"));
            //Clone the sheet multiple times
            workbook.CloneSheet(0);
            workbook.CloneSheet(0);

            Assert.IsNotNull(workbook.GetSheet("Test Clone"));
            Assert.IsNotNull(workbook.GetSheet("Test Clone (2)"));
            Assert.AreEqual("Test Clone (3)", workbook.GetSheetName(2));
            Assert.IsNotNull(workbook.GetSheet("Test Clone (3)"));

            workbook.RemoveSheetAt(0);
            workbook.RemoveSheetAt(0);
            workbook.RemoveSheetAt(0);
            workbook.CreateSheet("abc ( 123)");
            workbook.CloneSheet(0);
            Assert.AreEqual("abc (124)", workbook.GetSheetName(1));
        }

        /**
         * Setting landscape and portrait stuff on new sheets
         */
        [TestMethod]
        public void TestPrintSetupLandscapeNew()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheetL = workbook.CreateSheet("LandscapeS");
            Sheet sheetP = workbook.CreateSheet("LandscapeP");

            // Check two aspects of the print Setup
            Assert.IsFalse(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetP.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(1, sheetP.PrintSetup.Copies);

            // Change one on each
            sheetL.PrintSetup.Landscape = true;
            sheetP.PrintSetup.Copies = (short)3;

            // Check taken
            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetP.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetP.PrintSetup.Copies);

            // Save and re-load, and check still there
            workbook = GetTestDataProvider().WriteOutAndReadBack(workbook);
            sheetL = workbook.GetSheet("LandscapeS");
            sheetP = workbook.GetSheet("LandscapeP");

            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetP.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetP.PrintSetup.Copies);
        }

        /**
         * Test Adding merged regions. If the region's bounds are outside of the allowed range
         * then an ArgumentException should be thrown
         *
         */
        [TestMethod]
        public void TestAddMerged()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = wb.CreateSheet();
            Assert.AreEqual(0, sheet.NumMergedRegions);
            SpreadsheetVersion ssVersion = GetTestDataProvider().GetSpreadsheetVersion();

            CellRangeAddress region = new CellRangeAddress(0, 1, 0, 1);
            sheet.AddMergedRegion(region);
            Assert.AreEqual(1, sheet.NumMergedRegions);

            try
            {
                region = new CellRangeAddress(-1, -1, -1, -1);
                sheet.AddMergedRegion(region);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {

            }
            try
            {
                region = new CellRangeAddress(0, 0, 0, ssVersion.LastColumnIndex + 1);
                sheet.AddMergedRegion(region);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {

            }
            try
            {
                region = new CellRangeAddress(0, ssVersion.LastRowIndex + 1, 0, 1);
                sheet.AddMergedRegion(region);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {

            }
            Assert.AreEqual(1, sheet.NumMergedRegions);

        }

        /**
         * When removing one merged region, it would break
         *
         */
        [TestMethod]
        public void TestRemoveMerged()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = wb.CreateSheet();
            CellRangeAddress region = new CellRangeAddress(0, 1, 0, 1);
            sheet.AddMergedRegion(region);
            region = new CellRangeAddress(1, 2, 0, 1);
            sheet.AddMergedRegion(region);

            sheet.RemoveMergedRegion(0);

            region = sheet.GetMergedRegion(0);
            Assert.AreEqual(1, region.FirstRow, "Left over region should be starting at row 1");

            sheet.RemoveMergedRegion(0);

            Assert.AreEqual(0, sheet.NumMergedRegions, "there should be no merged regions left!");

            //an, Add, Remove, Get(0) would null pointer
            sheet.AddMergedRegion(region);
            Assert.AreEqual(1, sheet.NumMergedRegions, "there should now be one merged region!");
            sheet.RemoveMergedRegion(0);
            Assert.AreEqual(0, sheet.NumMergedRegions, "there should now be zero merged regions!");
            //add it again!
            region.LastRow = 4;

            sheet.AddMergedRegion(region);
            Assert.AreEqual(1, sheet.NumMergedRegions, "there should now be one merged region!");

            //should exist now!
            Assert.IsTrue(1 <= sheet.NumMergedRegions, "there isn't more than one merged region in there");
            region = sheet.GetMergedRegion(0);
            Assert.AreEqual(4, region.LastRow, "the merged row to doesnt match the one we put in ");
        }
        [TestMethod]
        public void TestShiftMerged()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            CreationHelper factory = wb.GetCreationHelper();
            Sheet sheet = wb.CreateSheet();
            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            cell.SetCellValue(factory.CreateRichTextString("first row, first cell"));

            row = sheet.CreateRow(1);
            cell = row.CreateCell(1);
            cell.SetCellValue(factory.CreateRichTextString("second row, second cell"));

            CellRangeAddress region = new CellRangeAddress(1, 1, 0, 1);
            sheet.AddMergedRegion(region);

            sheet.ShiftRows(1, 1, 1);

            region = sheet.GetMergedRegion(0);
            Assert.AreEqual(2, region.FirstRow, "Merged region not moved over to row 2");
        }

        /**
         * Tests the display of gridlines, formulas, and rowcolheadings.
         * @author Shawn Laubach (slaubach at apache dot org)
         */
        [TestMethod]
        public void TestDisplayOptions()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = wb.CreateSheet();

            Assert.AreEqual(sheet.DisplayGridlines, true);
            Assert.AreEqual(sheet.DisplayRowColHeadings, true);
            Assert.AreEqual(sheet.DisplayFormulas, false);
            Assert.AreEqual(sheet.DisplayZeros, true);

            sheet.DisplayGridlines=(false);
            sheet.DisplayRowColHeadings=(false);
            sheet.DisplayFormulas=(true);
            sheet.DisplayZeros=(false);

            wb = GetTestDataProvider().WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);

            Assert.AreEqual(sheet.DisplayGridlines, false);
            Assert.AreEqual(sheet.DisplayRowColHeadings, false);
            Assert.AreEqual(sheet.DisplayFormulas, true);
            Assert.AreEqual(sheet.DisplayZeros, false);
        }
        [TestMethod]
        public void TestColumnWidth()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = wb.CreateSheet();

            //default column width measured in characters
            sheet.DefaultColumnWidth=(10);
            Assert.AreEqual(10, sheet.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth(0));
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth(1));
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth(2));

            //set custom width for D-F
            for (char i = 'D'; i <= 'F'; i++)
            {
                //Sheet#SetColumnWidth accepts the width in units of 1/256th of a character width
                int w = 256 * 12;
                sheet.SetColumnWidth(i, w);
                Assert.AreEqual(w, sheet.GetColumnWidth(i));
            }
            //reset the default column width, columns A-C Change, D-F still have custom width
            sheet.DefaultColumnWidth=(20);
            Assert.AreEqual(20, sheet.DefaultColumnWidth);
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(0));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(1));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(2));
            for (char i = 'D'; i <= 'F'; i++)
            {
                int w = 256 * 12;
                Assert.AreEqual(w, sheet.GetColumnWidth(i));
            }

            // check for 16-bit signed/unsigned error:
            sheet.SetColumnWidth(10, 40000);
            Assert.AreEqual(40000, sheet.GetColumnWidth(10));

            //The maximum column width for an individual cell is 255 characters
            try
            {
                sheet.SetColumnWidth(9, 256 * 256);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The maximum column width for an individual cell is 255 characters.", e.Message);
            }

            //serialize and read again
            wb = GetTestDataProvider().WriteOutAndReadBack(wb);

            sheet = wb.GetSheetAt(0);
            Assert.AreEqual(20, sheet.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(0));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(1));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(2));
            //columns D-F have custom width
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                Assert.AreEqual(w, sheet.GetColumnWidth(i));
            }
            Assert.AreEqual(40000, sheet.GetColumnWidth(10));
        }
        [TestMethod]
        public void TestDefaultRowHeight()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet();
            sheet.DefaultRowHeightInPoints = (15);
            Assert.AreEqual((short)300, sheet.DefaultRowHeight);
            Assert.AreEqual(15.0F, sheet.DefaultRowHeightInPoints, 0F);

            // Set a new default row height in twips and test Getting the value in points
            sheet.DefaultRowHeight = ((short)360);
            Assert.AreEqual(18.0f, sheet.DefaultRowHeightInPoints, 0F);
            Assert.AreEqual((short)360, sheet.DefaultRowHeight);

            // Test that defaultRowHeight is a tRuncated short: E.G. 360inPoints -> 18; 361inPoints -> 18
            sheet.DefaultRowHeight = ((short)361);
            Assert.AreEqual((float)361 / 20, sheet.DefaultRowHeightInPoints, 0F);
            Assert.AreEqual((short)361, sheet.DefaultRowHeight);

            // Set a new default row height in points and test Getting the value in twips
            sheet.DefaultRowHeightInPoints = (17.5f);
            Assert.AreEqual(17.5f, sheet.DefaultRowHeightInPoints, 0F);
            Assert.AreEqual((short)(17.5f * 20), sheet.DefaultRowHeight);
        }

        /** cell with formula becomes null on cloning a sheet*/
        [TestMethod]
        public void Test35084()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            Sheet s = wb.CreateSheet("Sheet1");
            Row r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(1);
            r.CreateCell(1).CellFormula = ("A1*2");
            Sheet s1 = wb.CloneSheet(0);
            r = s1.GetRow(0);
            Assert.AreEqual(r.GetCell(0).NumericCellValue, 1, 0, "double"); // sanity check
            Assert.IsNotNull(r.GetCell(1));
            Assert.AreEqual(r.GetCell(1).CellFormula, "A1*2", "formula");
        }

        /** test that new default column styles get applied */
        [TestMethod]
        public void TestDefaultColumnStyle()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            CellStyle style = wb.CreateCellStyle();
            Sheet sheet = wb.CreateSheet();
            sheet.SetDefaultColumnStyle(0, style);
            Assert.IsNotNull(sheet.GetColumnStyle(0));
            Assert.AreEqual(style.Index, sheet.GetColumnStyle(0).Index);

            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            CellStyle style2 = cell.CellStyle;
            Assert.IsNotNull(style2);
            Assert.AreEqual(style.Index, style2.Index, "style should match");
        }
        [TestMethod]
        public void TestOutlineProperties()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();

            Sheet sheet = wb.CreateSheet();

            //TODO defaults are different in HSSF and XSSF
            //Assert.IsTrue(sheet.RowSumsBelow);
            //Assert.IsTrue(sheet.RowSumsRight);

            sheet.RowSumsBelow = (false);
            sheet.RowSumsRight = (false);

            Assert.IsFalse(sheet.RowSumsBelow);
            Assert.IsFalse(sheet.RowSumsRight);

            sheet.RowSumsBelow = (true);
            sheet.RowSumsRight = (true);

            Assert.IsTrue(sheet.RowSumsBelow);
            Assert.IsTrue(sheet.RowSumsRight);

            wb = GetTestDataProvider().WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            Assert.IsTrue(sheet.RowSumsBelow);
            Assert.IsTrue(sheet.RowSumsRight);
        }

        /**
         * Test basic display properties
         */
        [TestMethod]
        public void TestSheetProperties()
        {
            Workbook wb = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = wb.CreateSheet();

            Assert.IsFalse(sheet.HorizontallyCenter);
            sheet.HorizontallyCenter = true;
            Assert.IsTrue(sheet.HorizontallyCenter);
            sheet.HorizontallyCenter = false;
            Assert.IsFalse(sheet.HorizontallyCenter);

            Assert.IsFalse(sheet.VerticallyCenter);
            sheet.VerticallyCenter = true;
            Assert.IsTrue(sheet.VerticallyCenter);
            sheet.VerticallyCenter = false;
            Assert.IsFalse(sheet.VerticallyCenter);

            Assert.IsFalse(sheet.IsPrintGridlines);
            sheet.IsPrintGridlines = true;
            Assert.IsTrue(sheet.IsPrintGridlines);

            Assert.IsFalse(sheet.DisplayFormulas);
            sheet.DisplayFormulas = true;
            Assert.IsTrue(sheet.DisplayFormulas);

            Assert.IsTrue(sheet.DisplayGridlines);
            sheet.DisplayGridlines = false;
            Assert.IsFalse(sheet.DisplayGridlines);

            //TODO: default "guts" is different in HSSF and XSSF
            //Assert.IsTrue(sheet.DisplayGuts);
            sheet.DisplayGuts = false;
            Assert.IsFalse(sheet.DisplayGuts);

            Assert.IsTrue(sheet.DisplayRowColHeadings);
            sheet.DisplayRowColHeadings = false;
            Assert.IsFalse(sheet.DisplayRowColHeadings);

            //TODO: default "autobreaks" is different in HSSF and XSSF
            //Assert.IsTrue(sheet.Autobreaks);
            sheet.Autobreaks = false;
            Assert.IsFalse(sheet.Autobreaks);

            Assert.IsFalse(sheet.ScenarioProtect);

            //TODO: default "fit-to-page" is different in HSSF and XSSF
            //Assert.IsFalse(sheet.FitToPage);
            sheet.FitToPage = (true);
            Assert.IsTrue(sheet.FitToPage);
            sheet.FitToPage = (false);
            Assert.IsFalse(sheet.FitToPage);
        }
        [TestMethod]
        public void BaseTestGetSetMargin(double[] defaultMargins)
        {
            double marginLeft = defaultMargins[0];
            double marginRight = defaultMargins[1];
            double marginTop = defaultMargins[2];
            double marginBottom = defaultMargins[3];
            double marginHeader = defaultMargins[4];
            double marginFooter = defaultMargins[5];

            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet("Sheet 1");
            Assert.AreEqual(marginLeft, sheet.GetMargin(MarginType.LeftMargin), 0.0);
            sheet.SetMargin(MarginType.LeftMargin, 10.0);
            //left margin is custom, all others are default
            Assert.AreEqual(10.0, sheet.GetMargin(MarginType.LeftMargin), 0.0);
            Assert.AreEqual(marginRight, sheet.GetMargin(MarginType.RightMargin), 0.0);
            Assert.AreEqual(marginTop, sheet.GetMargin(MarginType.TopMargin), 0.0);
            Assert.AreEqual(marginBottom, sheet.GetMargin(MarginType.BottomMargin), 0.0);
            sheet.SetMargin(MarginType.RightMargin, 11.0);
            Assert.AreEqual(11.0, sheet.GetMargin(MarginType.RightMargin), 0.0);
            sheet.SetMargin(MarginType.TopMargin, 12.0);
            Assert.AreEqual(12.0, sheet.GetMargin(MarginType.TopMargin), 0.0);
            sheet.SetMargin(MarginType.BottomMargin, 13.0);
            Assert.AreEqual(13.0, sheet.GetMargin(MarginType.BottomMargin), 0.0);

            //// incorrect margin constant
            //try {
            //    sheet.SetMargin((short) 65, 15);
            //    Assert.Fail("Expected exception");
            //} catch (ArgumentException e){
            //    Assert.AreEqual("Unknown margin constant:  65", e.GetMessage());
            //}
        }
        [TestMethod]
        public void TestRowBreaks()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet();
            //Sheet#RowBreaks returns an empty array if no row breaks are defined
            Assert.IsNotNull(sheet.RowBreaks);
            Assert.AreEqual(0, sheet.RowBreaks.Length);

            sheet.SetRowBreak(1);
            Assert.AreEqual(1, sheet.RowBreaks.Length);
            sheet.SetRowBreak(15);
            Assert.AreEqual(2, sheet.RowBreaks.Length);
            Assert.AreEqual(1, sheet.RowBreaks[0]);
            Assert.AreEqual(15, sheet.RowBreaks[1]);
            sheet.SetRowBreak(1);
            Assert.AreEqual(2, sheet.RowBreaks.Length);
            Assert.IsTrue(sheet.IsRowBroken(1));
            Assert.IsTrue(sheet.IsRowBroken(15));

            //now remove the Created breaks
            sheet.RemoveRowBreak(1);
            Assert.AreEqual(1, sheet.RowBreaks.Length);
            sheet.RemoveRowBreak(15);
            Assert.AreEqual(0, sheet.RowBreaks.Length);

            Assert.IsFalse(sheet.IsRowBroken(1));
            Assert.IsFalse(sheet.IsRowBroken(15));
        }
        [TestMethod]
        public void TestColumnBreaks()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet();
            Assert.IsNotNull(sheet.ColumnBreaks);
            Assert.AreEqual(0, sheet.ColumnBreaks.Length);

            Assert.IsFalse(sheet.IsColumnBroken(0));

            sheet.SetColumnBreak(11);
            Assert.IsNotNull(sheet.ColumnBreaks);
            Assert.AreEqual(11, sheet.ColumnBreaks[0]);
            sheet.SetColumnBreak(12);
            Assert.AreEqual(2, sheet.ColumnBreaks.Length);
            Assert.IsTrue(sheet.IsColumnBroken(11));
            Assert.IsTrue(sheet.IsColumnBroken(12));

            sheet.RemoveColumnBreak((short)11);
            Assert.AreEqual(1, sheet.ColumnBreaks.Length);
            sheet.RemoveColumnBreak((short)15); //remove non-existing
            Assert.AreEqual(1, sheet.ColumnBreaks.Length);
            sheet.RemoveColumnBreak((short)12);
            Assert.AreEqual(0, sheet.ColumnBreaks.Length);

            Assert.IsFalse(sheet.IsColumnBroken(11));
            Assert.IsFalse(sheet.IsColumnBroken(12));
        }
        [TestMethod]
        public void TestGetFirstLastRowNum()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet("Sheet 1");
            sheet.CreateRow(9);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            Assert.AreEqual(0, sheet.FirstRowNum);
            Assert.AreEqual(9, sheet.LastRowNum);
        }
        [TestMethod]
        public void TestFooter()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet("Sheet 1");
            Assert.IsNotNull(sheet.Footer);
            sheet.Footer.Center = ("test center footer");
            Assert.AreEqual("test center footer", sheet.Footer.Center);
        }
        [TestMethod]
        public void TestGetSetColumnHidden()
        {
            Workbook workbook = GetTestDataProvider().CreateWorkbook();
            Sheet sheet = workbook.CreateSheet("Sheet 1");
            sheet.SetColumnHidden(2, true);
            Assert.IsTrue(sheet.IsColumnHidden(2));
        }
    }
}






