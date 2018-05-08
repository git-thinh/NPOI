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
    using System.IO;
    using System.Text;
    using System.Collections;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;

    using NPOI.HSSF.UserModel;
    //using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.SS.UserModel;

    /**
     * Testcases for bugs entered in bugzilla
     * the Test name contains the bugzilla bug id
     * @author Avik Sengupta
     * @author Yegor Kozlov
     */
    [TestClass]
    public class TestBugs
    {

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        private static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack(original);
        }

        private static void WriteTestOutputFileForViewing(HSSFWorkbook wb, String simpleFileName)
        {
            if (true)
            { // set to false to output Test files
                return;
            }
            string file = TempFile.GetTempFilePath(simpleFileName + "#", ".xls");
            FileStream out1 = new FileStream(file, FileMode.Create);
            wb.Write(out1);
            out1.Close();

            if (!File.Exists(file))
            {
                throw new Exception("File was not written");
            }
            Console.WriteLine("Open file '" + Path.GetFullPath(file) + "' in Excel");
        }

        /** Test reading AND writing a complicated workbook
         *Test Opening resulting sheet in excel*/
        [TestMethod]
        public void Test15228()
        {
            HSSFWorkbook wb = OpenSample("15228.xls");
            Sheet s = wb.GetSheetAt(0);
            Row r = s.CreateRow(0);
            Cell c = r.CreateCell(0);
            c.SetCellValue(10);
            WriteTestOutputFileForViewing(wb, "Test15228");
        }
        [TestMethod]
        public void Test13796()
        {
            HSSFWorkbook wb = OpenSample("13796.xls");
            Sheet s = wb.GetSheetAt(0);
            Row r = s.CreateRow(0);
            Cell c = r.CreateCell(0);
            c.SetCellValue(10);
            WriteOutAndReadBack(wb);
        }
        /**Test writing a hyperlink
         * Open resulting sheet in Excel and Check that A1 contains a hyperlink*/
        [TestMethod]
        public void Test23094()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet s = wb.CreateSheet();
            Row r = s.CreateRow(0);
            r.CreateCell(0).CellFormula = ("HYPERLINK( \"http://jakarta.apache.org\", \"Jakarta\" )");

            WriteTestOutputFileForViewing(wb, "Test23094");
        }

        /** Test hyperlinks
         * Open resulting file in excel, and Check that there is a link to Google
         */
        [TestMethod]
        public void Test15353()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet = wb.CreateSheet("My sheet");

            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            cell.CellFormula = ("HYPERLINK(\"http://google.com\",\"Google\")");

            WriteOutAndReadBack(wb);
        }

        /** Test reading of a formula with a name and a cell ref in one
         **/
        [TestMethod]
        public void Test14460()
        {
            HSSFWorkbook wb = OpenSample("14460.xls");
            wb.GetSheetAt(0);
        }
        [TestMethod]
        public void Test14330()
        {
            HSSFWorkbook wb = OpenSample("14330-1.xls");
            wb.GetSheetAt(0);

            wb = OpenSample("14330-2.xls");
            wb.GetSheetAt(0);
        }

        private static void setCellText(Cell cell, String text)
        {
            cell.SetCellValue(new HSSFRichTextString(text));
        }

        /** Test rewriting a file with large number of unique strings
         *Open resulting file in Excel to Check results!*/
        [TestMethod]
        public void Test15375()
        {
            HSSFWorkbook wb = OpenSample("15375.xls");
            Sheet sheet = wb.GetSheetAt(0);

            Row row = sheet.GetRow(5);
            Cell cell = row.GetCell(3);
            if (cell == null)
                cell = row.CreateCell(3);

            // Write Test
            cell.SetCellType(CellType.STRING);
            setCellText(cell, "a Test");

            // change existing numeric cell value

            Row oRow = sheet.GetRow(14);
            Cell oCell = oRow.GetCell(4);
            oCell.SetCellValue(75);
            oCell = oRow.GetCell(5);
            setCellText(oCell, "0.3");

            WriteTestOutputFileForViewing(wb, "Test15375");
        }

        /** Test writing a file with large number of unique strings
         *Open resulting file in Excel to Check results!*/
        [TestMethod]
        public void Test15375_2()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet = wb.CreateSheet();

            String tmp1 = null;
            String tmp2 = null;
            String tmp3 = null;

            for (int i = 0; i < 6000; i++)
            {
                tmp1 = "Test1" + i;
                tmp2 = "Test2" + i;
                tmp3 = "Test3" + i;

                Row row = sheet.CreateRow(i);

                Cell cell = row.CreateCell(0);
                setCellText(cell, tmp1);
                cell = row.CreateCell(1);
                setCellText(cell, tmp2);
                cell = row.CreateCell(2);
                setCellText(cell, tmp3);
            }
            WriteTestOutputFileForViewing(wb, "Test15375-2");
        }
        /** another Test for the number of unique strings issue
         *Test Opening the resulting file in Excel*/
        [TestMethod]
        public void Test22568()
        {
            int r = 2000; int c = 3;

            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet = wb.CreateSheet("ExcelTest");

            int col_cnt = 0, rw_cnt = 0;

            col_cnt = c;
            rw_cnt = r;

            Row rw;
            rw = sheet.CreateRow(0);
            //Header row
            for (int j = 0; j < col_cnt; j++)
            {
                Cell cell = rw.CreateCell(j);
                setCellText(cell, "Col " + (j + 1));
            }

            for (int i = 1; i < rw_cnt; i++)
            {
                rw = sheet.CreateRow(i);
                for (int j = 0; j < col_cnt; j++)
                {
                    Cell cell = rw.CreateCell(j);
                    setCellText(cell, "Row:" + (i + 1) + ",Column:" + (j + 1));
                }
            }

            sheet.DefaultColumnWidth = 18;

            WriteTestOutputFileForViewing(wb, "Test22568");
        }

        /**Double byte strings*/
        [TestMethod]
        public void Test15556()
        {

            HSSFWorkbook wb = OpenSample("15556.xls");
            Sheet sheet = wb.GetSheetAt(0);
            Row row = sheet.GetRow(45);
            Assert.IsNotNull(row, "Read row fine!");
        }
        /**Double byte strings */
        [TestMethod]
        public void Test22742()
        {
            OpenSample("22742.xls");
        }
        /**Double byte strings */
        [TestMethod]
        public void Test12561_1()
        {
            OpenSample("12561-1.xls");
        }
        /** Double byte strings */
        [TestMethod]
        public void Test12561_2()
        {
            OpenSample("12561-2.xls");
        }
        /** Double byte strings
         File supplied by jubeson*/
        [TestMethod]
        public void Test12843_1()
        {
            OpenSample("12843-1.xls");
        }

        /** Double byte strings
         File supplied by Paul Chung*/
        [TestMethod]
        public void Test12843_2()
        {
            OpenSample("12843-2.xls");
        }

        /** Reference to Name*/
        [TestMethod]
        public void Test13224()
        {
            OpenSample("13224.xls");
        }

        /** Illegal argument exception - cannot store duplicate value in Map*/
        [TestMethod]
        public void Test19599()
        {
            OpenSample("19599-1.xls");
            OpenSample("19599-2.xls");
        }
        [TestMethod]
        public void Test24215()
        {
            HSSFWorkbook wb = OpenSample("24215.xls");

            for (int sheetIndex = 0; sheetIndex < wb.NumberOfSheets; sheetIndex++)
            {
                Sheet sheet = wb.GetSheetAt(sheetIndex);
                int rows = sheet.LastRowNum;

                for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                {
                    Row row = sheet.GetRow(rowIndex);
                    int cells = row.LastCellNum;

                    for (int cellIndex = 0; cellIndex < cells; cellIndex++)
                    {
                        row.GetCell(cellIndex);
                    }
                }
            }
        }
        [TestMethod]
        public void Test18800()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            book.CreateSheet("TEST");
            Sheet sheet = book.CloneSheet(0);
            book.SetSheetName(1, "CLONE");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Test"));

            book = WriteOutAndReadBack(book);
            sheet = book.GetSheet("CLONE");
            Row row = sheet.GetRow(0);
            Cell cell = row.GetCell(0);
            Assert.AreEqual("Test", cell.RichStringCellValue.String);
        }

        /**
         * Merged regions were being Removed from the parent in cloned sheets
         */
        [TestMethod]
        public void Test22720()
        {
            HSSFWorkbook workBook = new HSSFWorkbook();
            workBook.CreateSheet("TEST");
            Sheet template = workBook.GetSheetAt(0);

            template.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
            template.AddMergedRegion(new CellRangeAddress(1, 2, 0, 2));

            Sheet clone = workBook.CloneSheet(0);
            int originalMerged = template.NumMergedRegions;
            Assert.AreEqual(2, originalMerged, "2 merged regions");

            //        Remove merged regions from clone
            for (int i = template.NumMergedRegions - 1; i >= 0; i--)
            {
                clone.RemoveMergedRegion(i);
            }

            Assert.AreEqual(originalMerged, template.NumMergedRegions, "Original Sheet's Merged Regions were Removed");
            //        Check if template's merged regions are OK
            if (template.NumMergedRegions > 0)
            {
                // fetch the first merged region...EXCEPTION OCCURS HERE
                template.GetMergedRegion(0);
            }
            //make sure we dont exception

        }

        /**Tests read and Write of Unicode strings in formula results
         * bug and Testcase submitted by Sompop Kumnoonsate
         * The file contains THAI unicode characters.
         */
        [TestMethod]
        public void TestUnicodeStringFormulaRead()
        {

            HSSFWorkbook w = OpenSample("25695.xls");

            Cell a1 = w.GetSheetAt(0).GetRow(0).GetCell(0);
            Cell a2 = w.GetSheetAt(0).GetRow(0).GetCell(1);
            Cell b1 = w.GetSheetAt(0).GetRow(1).GetCell(0);
            Cell b2 = w.GetSheetAt(0).GetRow(1).GetCell(1);
            Cell c1 = w.GetSheetAt(0).GetRow(2).GetCell(0);
            Cell c2 = w.GetSheetAt(0).GetRow(2).GetCell(1);
            Cell d1 = w.GetSheetAt(0).GetRow(3).GetCell(0);
            Cell d2 = w.GetSheetAt(0).GetRow(3).GetCell(1);

            if (false)
            {
                // THAI code page
                Console.WriteLine("a1=" + unicodeString(a1));
                Console.WriteLine("a2=" + unicodeString(a2));
                // US code page
                Console.WriteLine("b1=" + unicodeString(b1));
                Console.WriteLine("b2=" + unicodeString(b2));
                // THAI+US
                Console.WriteLine("c1=" + unicodeString(c1));
                Console.WriteLine("c2=" + unicodeString(c2));
                // US+THAI
                Console.WriteLine("d1=" + unicodeString(d1));
                Console.WriteLine("d2=" + unicodeString(d2));
            }
            ConfirmSameCellText(a1, a2);
            ConfirmSameCellText(b1, b2);
            ConfirmSameCellText(c1, c2);
            ConfirmSameCellText(d1, d2);

            HSSFWorkbook rw = WriteOutAndReadBack(w);

            Cell ra1 = rw.GetSheetAt(0).GetRow(0).GetCell(0);
            Cell ra2 = rw.GetSheetAt(0).GetRow(0).GetCell(1);
            Cell rb1 = rw.GetSheetAt(0).GetRow(1).GetCell(0);
            Cell rb2 = rw.GetSheetAt(0).GetRow(1).GetCell(1);
            Cell rc1 = rw.GetSheetAt(0).GetRow(2).GetCell(0);
            Cell rc2 = rw.GetSheetAt(0).GetRow(2).GetCell(1);
            Cell rd1 = rw.GetSheetAt(0).GetRow(3).GetCell(0);
            Cell rd2 = rw.GetSheetAt(0).GetRow(3).GetCell(1);

            ConfirmSameCellText(a1, ra1);
            ConfirmSameCellText(b1, rb1);
            ConfirmSameCellText(c1, rc1);
            ConfirmSameCellText(d1, rd1);

            ConfirmSameCellText(a1, ra2);
            ConfirmSameCellText(b1, rb2);
            ConfirmSameCellText(c1, rc2);
            ConfirmSameCellText(d1, rd2);
        }

        private static void ConfirmSameCellText(Cell a, Cell b)
        {
            Assert.AreEqual(a.RichStringCellValue.String, b.RichStringCellValue.String);
        }
        private static String unicodeString(Cell cell)
        {
            String ss = cell.RichStringCellValue.String;
            char[] s = ss.ToCharArray();
            StringBuilder sb = new StringBuilder();
            for (int x = 0; x < s.Length; x++)
            {
                sb.Append("\\u").Append(StringUtil.ToHexString(s[x]));
            }
            return sb.ToString();
        }

        /** Error in Opening wb*/
        [TestMethod]
        public void Test32822()
        {
            OpenSample("32822.xls");
        }
        /**Assert.Fail to read wb with chart */
        [TestMethod]
        public void Test15573()
        {
            OpenSample("15573.xls");
        }

        /**names and macros */
        [TestMethod]
        public void Test27852()
        {
            HSSFWorkbook wb = OpenSample("27852.xls");

            for (int i = 0; i < wb.NumberOfNames; i++)
            {
                NPOI.SS.UserModel.Name name = wb.GetNameAt(i);
                //name.NameName();
                if (name.IsFunctionName)
                {
                    continue;
                }
                //name.Reference;
            }
        }
        [TestMethod]
        public void Test28031()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            String formulaText =
                "IF(ROUND(A2*B2*C2,2)>ROUND(B2*D2,2),ROUND(A2*B2*C2,2),ROUND(B2*D2,2))";
            cell.CellFormula = (formulaText);

            Assert.AreEqual(formulaText, cell.CellFormula);
            WriteTestOutputFileForViewing(wb, "output28031.xls");
        }
        [TestMethod]
        public void Test33082()
        {
            OpenSample("33082.xls");
        }
        [TestMethod]
        public void Test34775()
        {
            try
            {
                OpenSample("34775.xls");
            }
            catch (NullReferenceException e)
            {
                throw new AssertFailedException("identified bug 34775");
            }
        }

        /** Error when reading then writing ArrayValues in NameRecord's*/
        [TestMethod]
        public void Test37630()
        {
            HSSFWorkbook wb = OpenSample("37630.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 25183: org.apache.poi.hssf.usermodel.Sheet.SetPropertiesFromSheet
         */
        [TestMethod]
        public void Test25183()
        {
            HSSFWorkbook wb = OpenSample("25183.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 26100: 128-character message in IF statement cell causes HSSFWorkbook Open Assert.Failure
         */
        [TestMethod]
        public void Test26100()
        {
            HSSFWorkbook wb = OpenSample("26100.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 27933: Unable to use a template (xls) file containing a wmf graphic
         */
        [TestMethod]
        public void Test27933()
        {
            HSSFWorkbook wb = OpenSample("27933.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 29206:      NPE on Sheet.GetRow for blank rows
         */
        [TestMethod]
        public void Test29206()
        {
            //the first Check with blank workbook
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet = wb.CreateSheet();

            for (int i = 1; i < 400; i++)
            {
                Row row = sheet.GetRow(i);
                if (row != null)
                {
                    row.GetCell(0);
                }
            }

            //now Check on an existing xls file
            wb = OpenSample("Simple.xls");

            for (int i = 1; i < 400; i++)
            {
                Row row = sheet.GetRow(i);
                if (row != null)
                {
                    row.GetCell(0);
                }
            }
        }

        /**
         * Bug 29675: POI 2.5 corrupts output when starting workbook has a graphic
         */
        [TestMethod]
        public void Test29675()
        {
            HSSFWorkbook wb = OpenSample("29675.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 29942: Importing Excel files that have been Created by Open Office on Linux
         */
        [TestMethod]
        public void Test29942()
        {
            HSSFWorkbook wb = OpenSample("29942.xls");

            Sheet sheet = wb.GetSheetAt(0);
            int count = 0;
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                Row row = sheet.GetRow(i);
                if (row != null)
                {
                    Cell cell = row.GetCell(0);
                    Assert.AreEqual(CellType.STRING, cell.CellType);
                    count++;
                }
            }
            Assert.AreEqual(85, count); //should read 85 rows

            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 29982: Unable to read spreadsheet when dropdown list cell is selected -
         *  Unable to construct record instance
         */
        [TestMethod]
        public void Test29982()
        {
            HSSFWorkbook wb = OpenSample("29982.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 30540: Sheet.SetRowBreak throws NullPointerException
         */
        [TestMethod]
        public void Test30540()
        {
            HSSFWorkbook wb = OpenSample("30540.xls");

            Sheet s = wb.GetSheetAt(0);
            s.SetRowBreak(1);
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 31749: {Need help urgently}[This is critical] workbook.Write() corrupts the file......?
         */
        [TestMethod]
        public void Test31749()
        {
            HSSFWorkbook wb = OpenSample("31749.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 31979: {urgent help needed .....}poi library does not support form objects properly.
         */
        public void Test31979()
        {
            HSSFWorkbook wb = OpenSample("31979.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 35564: Cell.java: NullPtrExc in isGridsPrinted() and getProtect()
         *  when HSSFWorkbook is Created from file
         */
        [TestMethod]
        public void Test35564()
        {
            HSSFWorkbook wb = OpenSample("35564.xls");

            Sheet sheet = wb.GetSheetAt(0);
            Assert.AreEqual(false, sheet.IsPrintGridlines);
            Assert.AreEqual(false, sheet.Protect);

            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 35565: Cell.java: NullPtrExc in getColumnBreaks() when HSSFWorkbook is Created from file
         */
        [TestMethod]
        public void Test35565()
        {
            HSSFWorkbook wb = OpenSample("35565.xls");

            Sheet sheet = wb.GetSheetAt(0);
            Assert.IsNotNull(sheet);
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 37376: Cannot Open the saved Excel file if Checkbox controls exceed certain limit
         */
        [TestMethod]
        public void Test37376()
        {
            HSSFWorkbook wb = OpenSample("37376.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 40285:      CellIterator Skips First Column
         */
        [TestMethod]
        public void Test40285()
        {
            HSSFWorkbook wb = OpenSample("40285.xls");

            Sheet sheet = wb.GetSheetAt(0);
            int rownum = 0;
            for (IEnumerator it = sheet.GetRowEnumerator(); it.MoveNext(); rownum++)
            {
                Row row = (Row)it.Current;
                Assert.AreEqual(rownum, row.RowNum);
                int cellNum = 0;
                for (IEnumerator it2 = row.GetCellEnumerator(); it2.MoveNext(); cellNum++)
                {
                    Cell cell = (Cell)it2.Current;
                    Assert.AreEqual(cellNum, cell.ColumnIndex);
                }
            }
        }

        /**
         * Bug 40296:      Cell.SetCellFormula throws
         *   ClassCastException if cell is Created using Row.CreateCell(short column, int type)
         */
        [TestMethod]
        public void Test40296()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            HSSFWorkbook workBook = new HSSFWorkbook();
            Sheet workSheet = workBook.CreateSheet("Sheet1");
            Cell cell;
            Row row = workSheet.CreateRow(0);
            cell = row.CreateCell(0, CellType.NUMERIC);
            cell.SetCellValue(1.0);
            cell = row.CreateCell(1, CellType.NUMERIC);
            cell.SetCellValue(2.0);
            cell = row.CreateCell(2, CellType.FORMULA);
            cell.CellFormula = ("SUM(A1:B1)");

            WriteOutAndReadBack(wb);
        }

        /**
         * Test bug 38266: NPE when Adding a row break
         *
         * User's diagnosis:
         * 1. Manually (i.e., not using POI) Create an Excel Workbook, making sure it
         * contains a sheet that doesn't have any row breaks
         * 2. Using POI, Create a new HSSFWorkbook from the template in step #1
         * 3. Try Adding a row break (via sheet.SetRowBreak()) to the sheet mentioned in step #1
         * 4. Get a NullPointerException
         */
        [TestMethod]
        public void Test38266()
        {
            String[] files = { "Simple.xls", "SimpleMultiCell.xls", "duprich1.xls" };
            for (int i = 0; i < files.Length; i++)
            {
                HSSFWorkbook wb = OpenSample(files[i]);

                Sheet sheet = wb.GetSheetAt(0);
                int[] breaks = sheet.RowBreaks;
                Assert.AreEqual(0, breaks.Length);

                //Add 3 row breaks
                for (int j = 1; j <= 3; j++)
                {
                    sheet.SetRowBreak(j * 20);
                }
            }
        }
        [TestMethod]
        public void Test40738()
        {
            HSSFWorkbook wb = OpenSample("SimpleWithAutofilter.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 44200: Sheet not cloneable when Note Added to excel cell
         */
        [TestMethod]
        public void Test44200()
        {
            HSSFWorkbook wb = OpenSample("44200.xls");

            wb.CloneSheet(0);
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 44201: Sheet not cloneable when validation Added to excel cell
         */
        [TestMethod]
        public void Test44201()
        {
            HSSFWorkbook wb = OpenSample("44201.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 37684  : Unhandled Continue Record Error
         */
        [TestMethod]
        public void Test37684()
        {
            HSSFWorkbook wb = OpenSample("37684-1.xls");
            WriteOutAndReadBack(wb);


            wb = OpenSample("37684-2.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 41139: Constructing HSSFWorkbook is Assert.Failed,threw threw ArrayIndexOutOfBoundsException for creating UnknownRecord
         */
        [TestMethod]
        public void Test41139()
        {
            HSSFWorkbook wb = OpenSample("41139.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 41546: Constructing HSSFWorkbook is Assert.Failed,
         *  Unknown Ptg in Formula: 0x1a (26)
         */
        [TestMethod]
        public void Test41546()
        {
            HSSFWorkbook wb = OpenSample("41546.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
            wb = WriteOutAndReadBack(wb);
            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * Bug 42564: Some files from Access were giving a RecordFormatException
         *  when reading the BOFRecord
         */
        [TestMethod]
        public void Test42564()
        {
            HSSFWorkbook wb = OpenSample("42564.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 42564: Some files from Access also have issues
         *  with the NameRecord, once you get past the BOFRecord
         *  issue.
         */
        [TestMethod]
        public void Test42564Alt()
        {
            HSSFWorkbook wb = OpenSample("42564-2.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 42618: RecordFormatException reading a file containing
         *     =CHOOSE(2,A2,A3,A4)
         */
        [TestMethod]
        public void Test42618()
        {
            HSSFWorkbook wb = OpenSample("SimpleWithChoose.xls");
            wb = WriteOutAndReadBack(wb);
            // Check we detect the string properly too
            Sheet s = wb.GetSheetAt(0);

            // Textual value
            Row r1 = s.GetRow(0);
            Cell c1 = r1.GetCell(1);
            Assert.AreEqual("=CHOOSE(2,A2,A3,A4)", c1.RichStringCellValue.ToString());

            // Formula Value
            Row r2 = s.GetRow(1);
            Cell c2 = r2.GetCell(1);
            Assert.AreEqual(25, (int)c2.NumericCellValue);

            try
            {
                Assert.AreEqual("CHOOSE(2,A2,A3,A4)", c2.CellFormula);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.StartsWith("Too few arguments")
                        && e.Message.IndexOf("ConcatPtg") > 0)
                {
                    throw new AssertFailedException("identified bug 44306");
                }
            }
        }

        /**
         * Something up with the FileSharingRecord
         */
        [TestMethod]
        public void Test43251()
        {

            // Used to blow up with an ArgumentException
            //  when creating a FileSharingRecord
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("43251.xls");
            }
            catch (ArgumentException e)
            {
                throw new AssertFailedException("identified bug 43251");
            }

            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * Crystal reports generates files with short
         *  StyleRecords, which is against the spec
         */
        [TestMethod]
        public void Test44471()
        {

            // Used to blow up with an ArrayIndexOutOfBounds
            //  when creating a StyleRecord
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("OddStyleRecord.xls");
            }
            catch (IndexOutOfRangeException e)
            {
                throw new AssertFailedException("Identified bug 44471");
            }

            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * Files with "read only recommended" were giving
         *  grief on the FileSharingRecord
         */
        [TestMethod]
        public void Test44536()
        {

            // Used to blow up with an ArgumentException
            //  when creating a FileSharingRecord
            HSSFWorkbook wb = OpenSample("ReadOnlyRecommended.xls");

            // Check read only advised
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.IsTrue(wb.IsWriteProtected);

            // But also Check that another wb isn't
            wb = OpenSample("SimpleWithChoose.xls");
            Assert.IsFalse(wb.IsWriteProtected);
        }

        /**
         * Some files were having problems with the DVRecord,
         *  probably due to dropdowns
         */
        [TestMethod]
        public void Test44593()
        {

            // Used to blow up with an ArgumentException
            //  when creating a DVRecord
            // Now won't, but no idea if this means we have
            //  rubbish in the DVRecord or not...
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("44593.xls");
            }
            catch (ArgumentException e)
            {
                throw new AssertFailedException("Identified bug 44593");
            }

            Assert.AreEqual(2, wb.NumberOfSheets);
        }

        /**
         * Used to give problems due to trying to read a zero
         *  Length string, but that's now properly handled
         */
        [TestMethod]
        public void Test44643()
        {

            // Used to blow up with an ArgumentException
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("44643.xls");
            }
            catch (ArgumentException e)
            {
                throw new AssertFailedException("identified bug 44643");
            }

            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * User reported the wrong number of rows from the
         *  iterator, but we can't replicate that
         */
        [TestMethod]
        public void Test44693()
        {

            HSSFWorkbook wb = OpenSample("44693.xls");
            Sheet s = wb.GetSheetAt(0);

            // Rows are 1 to 713
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(712, s.LastRowNum);
            Assert.AreEqual(713, s.PhysicalNumberOfRows);

            // Now Check the iterator
            int rowsSeen = 0;
            for (IEnumerator i = s.GetRowEnumerator(); i.MoveNext(); )
            {
                Row r = (Row)i.Current;
                Assert.IsNotNull(r);
                rowsSeen++;
            }
            Assert.AreEqual(713, rowsSeen);
        }

        /**
         * Bug 28774: Excel will crash when Opening xls-files with images.
         */
        [TestMethod]
        public void Test28774()
        {
            HSSFWorkbook wb = OpenSample("28774.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        /**
         * Had a problem apparently, not sure what as it
         *  works just fine...
         */
        [TestMethod]
        public void Test44891()
        {
            HSSFWorkbook wb = OpenSample("44891.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        /**
         * Bug 44235: Ms Excel can't Open save as excel file
         *
         * Works fine with poi-3.1-beta1.
         */
        [TestMethod]
        public void Test44235()
        {
            HSSFWorkbook wb = OpenSample("44235.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }


        [TestMethod]
        public void Test36947()
        {
            HSSFWorkbook wb = OpenSample("36947.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        [TestMethod]
        public void Test39634()
        {
            HSSFWorkbook wb = OpenSample("39634.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        /**
         * Problems with extracting Check boxes from
         *  HSSFObjectData
         * @
         */
        [TestMethod]
        public void Test44840()
        {
            HSSFWorkbook wb = OpenSample("WithCheckBoxes.xls");

            // Take a look at the embedded objects
            IList objects = wb.GetAllEmbeddedObjects();
            Assert.AreEqual(1, objects.Count);

            HSSFObjectData obj = (HSSFObjectData)objects[0];
            Assert.IsNotNull(obj);

            // Peek inside the underlying record
            EmbeddedObjectRefSubRecord rec = obj.FindObjectRecord();
            Assert.IsNotNull(rec);

            //        Assert.AreEqual(32, rec.field_1_stream_id_offset);
            Assert.AreEqual(0, rec.StreamId); // WRONG!
            Assert.AreEqual("Forms.CheckBox.1", rec.OLEClassName);
            Assert.AreEqual(12, rec.ObjectData.Length);

            // Doesn't have a directory
            Assert.IsFalse(obj.HasDirectoryEntry());
            Assert.IsNotNull(obj.GetObjectData());
            Assert.AreEqual(12, obj.GetObjectData().Length);
            Assert.AreEqual("Forms.CheckBox.1", obj.OLE2ClassName);

            try
            {
                obj.GetDirectory();
                Assert.Fail();
            }
            catch (FileNotFoundException e)
            {
                // expected during successful Test
            }
        }

        /**
         * Test that we can delete sheets without
         *  breaking the build in named ranges
         *  used for printing stuff.
         */
        [TestMethod]
        public void Test30978()
        {
            HSSFWorkbook wb = OpenSample("30978-alt.xls");
            Assert.AreEqual(1, wb.NumberOfNames);
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Check all names fit within range, and use
            //  DeletedArea3DPtg
            NPOI.HSSF.Model.Workbook w = wb.Workbook;
            for (int i = 0; i < w.NumNames; i++)
            {
                NameRecord r = w.GetNameRecord(i);
                Assert.IsTrue(r.SheetNumber <= wb.NumberOfSheets);

                Ptg[] nd = r.NameDefinition;
                Assert.AreEqual(1, nd.Length);
                Assert.IsTrue(nd[0] is DeletedArea3DPtg);
            }


            // Delete the 2nd sheet
            wb.RemoveSheetAt(1);


            // Re-Check
            Assert.AreEqual(1, wb.NumberOfNames);
            Assert.AreEqual(2, wb.NumberOfSheets);

            for (int i = 0; i < w.NumNames; i++)
            {
                NameRecord r = w.GetNameRecord(i);
                Assert.IsTrue(r.SheetNumber <= wb.NumberOfSheets);

                Ptg[] nd = r.NameDefinition;
                Assert.AreEqual(1, nd.Length);
                Assert.IsTrue(nd[0] is DeletedArea3DPtg);
            }


            // Save and re-load
            wb = WriteOutAndReadBack(wb);
            w = wb.Workbook;

            Assert.AreEqual(1, wb.NumberOfNames);
            Assert.AreEqual(2, wb.NumberOfSheets);

            for (int i = 0; i < w.NumNames; i++)
            {
                NameRecord r = w.GetNameRecord(i);
                Assert.IsTrue(r.SheetNumber <= wb.NumberOfSheets);

                Ptg[] nd = r.NameDefinition;
                Assert.AreEqual(1, nd.Length);
                Assert.IsTrue(nd[0] is DeletedArea3DPtg);
            }
        }

        /**
         * Test that fonts get Added properly
         */
        [TestMethod]
        public void Test45338()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Assert.AreEqual(4, wb.NumberOfFonts);

            Sheet s = wb.CreateSheet();
            s.CreateRow(0);
            s.CreateRow(1);
            Cell c1 = s.GetRow(0).CreateCell(0);
            Cell c2 = s.GetRow(1).CreateCell(0);

            Assert.AreEqual(4, wb.NumberOfFonts);

            Font f1 = wb.GetFontAt((short)0);
            Assert.AreEqual(400, f1.Boldweight);

            // Check that asking for the same font
            //  multiple times gives you the same thing.
            // Otherwise, our Tests wouldn't work!
            Assert.AreEqual(
                    wb.GetFontAt((short)0),
                    wb.GetFontAt((short)0)
            );
            Assert.AreEqual(
                    wb.GetFontAt((short)2),
                    wb.GetFontAt((short)2)
            );
            Assert.IsTrue(
                    wb.GetFontAt((short)0)
                    !=
                    wb.GetFontAt((short)2)
            );

            // Look for a new font we have
            //  yet to Add
            Assert.IsNull(
                wb.FindFont(
                    (short)11, (short)123, (short)22,
                    "Thingy", false, true, (short)2, (byte)2
                )
            );

            Font nf = wb.CreateFont();
            Assert.AreEqual(5, wb.NumberOfFonts);

            Assert.AreEqual(5, nf.Index);
            Assert.AreEqual(nf, wb.GetFontAt((short)5));

            nf.Boldweight = ((short)11);
            nf.Color = ((short)123);
            nf.FontHeight = ((short)22);
            nf.FontName = ("Thingy");
            nf.IsItalic = (false);
            nf.IsStrikeout = (true);
            nf.TypeOffset = ((short)2);
            nf.Underline = ((byte)2);

            Assert.AreEqual(5, wb.NumberOfFonts);
            Assert.AreEqual(nf, wb.GetFontAt((short)5));

            // Find it now
            Assert.IsNotNull(
                wb.FindFont(
                    (short)11, (short)123, (short)22,
                    "Thingy", false, true, (short)2, (byte)2
                )
            );
            Assert.AreEqual(
                5,
                wb.FindFont(
                       (short)11, (short)123, (short)22,
                       "Thingy", false, true, (short)2, (byte)2
                   ).Index
            );
            Assert.AreEqual(nf,
                   wb.FindFont(
                       (short)11, (short)123, (short)22,
                       "Thingy", false, true, (short)2, (byte)2
                   )
            );
        }

        /**
         * From the mailing list - ensure we can handle a formula
         *  containing a zip code, eg ="70164"
         */
        [TestMethod]
        public void TestZipCodeFormulas()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet s = wb.CreateSheet();
            s.CreateRow(0);
            Cell c1 = s.GetRow(0).CreateCell(0);
            Cell c2 = s.GetRow(0).CreateCell(1);
            Cell c3 = s.GetRow(0).CreateCell(2);

            // As number and string
            c1.CellFormula = ("70164");
            c2.CellFormula = ("\"70164\"");
            c3.CellFormula = ("\"90210\"");

            // Check the formulas
            Assert.AreEqual("70164", c1.CellFormula);
            Assert.AreEqual("\"70164\"", c2.CellFormula);

            // And Check the values - blank
            ConfirmCachedValue(0.0, c1);
            ConfirmCachedValue(0.0, c2);
            ConfirmCachedValue(0.0, c3);

            // Try changing the cached value on one of the string
            //  formula cells, so we can see it updates properly
            c3.SetCellValue(new HSSFRichTextString("Test"));
            ConfirmCachedValue("Test", c3);
            try
            {
                double a = c3.NumericCellValue;
                throw new AssertFailedException("exception should have been thrown");
            }
            catch (InvalidOperationException e)
            {
                Assert.AreEqual("Cannot get a numeric value from a text formula cell", e.Message);
            }


            // Now Evaluate, they should all be changed
            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            eval.EvaluateFormulaCell(c1);
            eval.EvaluateFormulaCell(c2);
            eval.EvaluateFormulaCell(c3);

            // Check that the cells now contain
            //  the correct values
            ConfirmCachedValue(70164.0, c1);
            ConfirmCachedValue("70164", c2);
            ConfirmCachedValue("90210", c3);


            // Write and read
            HSSFWorkbook nwb = WriteOutAndReadBack(wb);
            HSSFSheet ns = (HSSFSheet)nwb.GetSheetAt(0);
            Cell nc1 = ns.GetRow(0).GetCell(0);
            Cell nc2 = ns.GetRow(0).GetCell(1);
            Cell nc3 = ns.GetRow(0).GetCell(2);

            // Re-Check
            ConfirmCachedValue(70164.0, nc1);
            ConfirmCachedValue("70164", nc2);
            ConfirmCachedValue("90210", nc3);

            CellValueRecordInterface[] cvrs = ns.Sheet.GetValueRecords();
            for (int i = 0; i < cvrs.Length; i++)
            {
                CellValueRecordInterface cvr = cvrs[i];
                if (cvr is FormulaRecordAggregate)
                {
                    FormulaRecordAggregate fr = (FormulaRecordAggregate)cvr;

                    if (i == 0)
                    {
                        Assert.AreEqual(70164.0, fr.FormulaRecord.Value, 0.0001);
                        Assert.IsNull(fr.StringRecord);
                    }
                    else if (i == 1)
                    {
                        Assert.AreEqual(0.0, fr.FormulaRecord.Value, 0.0001);
                        Assert.IsNotNull(fr.StringRecord);
                        Assert.AreEqual("70164", fr.StringRecord.String);
                    }
                    else
                    {
                        Assert.AreEqual(0.0, fr.FormulaRecord.Value, 0.0001);
                        Assert.IsNotNull(fr.StringRecord);
                        Assert.AreEqual("90210", fr.StringRecord.String);
                    }
                }
            }
            Assert.AreEqual(3, cvrs.Length);
        }

        private static void ConfirmCachedValue(double expectedValue, Cell cell)
        {
            Assert.AreEqual(CellType.FORMULA, cell.CellType);
            Assert.AreEqual(CellType.NUMERIC, cell.CachedFormulaResultType);
            Assert.AreEqual(expectedValue, cell.NumericCellValue, 0.0);
        }
        private static void ConfirmCachedValue(String expectedValue, Cell cell)
        {
            Assert.AreEqual(CellType.FORMULA, cell.CellType);
            Assert.AreEqual(CellType.STRING, cell.CachedFormulaResultType);
            Assert.AreEqual(expectedValue, cell.RichStringCellValue.String);
        }

        /**
         * Problem with "Vector Rows", eg a whole
         *  column which is set to the result of
         *  {=sin(B1:B9)}(9,1), so that each cell is
         *  shown to have the contents
         *  {=sin(B1:B9){9,1)[rownum][0]
         * In this sample file, the vector column
         *  is C, and the data column is B.
         *
         * For now, blows up with an exception from ExtPtg
         *  Expected ExpPtg to be converted from Shared to Non-Shared...
         */
        [TestMethod]
        public void Test43623()
        {
            HSSFWorkbook wb = OpenSample("43623.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);

            Sheet s1 = wb.GetSheetAt(0);

            Cell c1 = s1.GetRow(0).GetCell(2);
            Cell c2 = s1.GetRow(1).GetCell(2);
            Cell c3 = s1.GetRow(2).GetCell(2);

            // These formula contents are a guess...
            Assert.AreEqual("{=sin(B1:B9){9,1)[0][0]", c1.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[1][0]", c2.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[2][0]", c3.CellFormula);

            // Save and re-Open, ensure it still works
            HSSFWorkbook nwb = WriteOutAndReadBack(wb);
            Sheet ns1 = nwb.GetSheetAt(0);
            Cell nc1 = ns1.GetRow(0).GetCell(2);
            Cell nc2 = ns1.GetRow(1).GetCell(2);
            Cell nc3 = ns1.GetRow(2).GetCell(2);

            Assert.AreEqual("{=sin(B1:B9){9,1)[0][0]", nc1.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[1][0]", nc2.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[2][0]", nc3.CellFormula);
        }

        /**
         * People are all getting confused about the last
         *  row and cell number
         */
        [TestMethod]
        public void Test30635()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet s = wb.CreateSheet();

            // No rows, everything is 0
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(0, s.LastRowNum);
            Assert.AreEqual(0, s.PhysicalNumberOfRows);

            // One row, most things are 0, physical is 1
            s.CreateRow(0);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(0, s.LastRowNum);
            Assert.AreEqual(1, s.PhysicalNumberOfRows);

            // And another, things change
            s.CreateRow(4);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(4, s.LastRowNum);
            Assert.AreEqual(2, s.PhysicalNumberOfRows);


            // Now start on cells
            Row r = s.GetRow(0);
            Assert.AreEqual(-1, r.FirstCellNum);
            Assert.AreEqual(-1, r.LastCellNum);
            Assert.AreEqual(0, r.PhysicalNumberOfCells);

            // Add a cell, things move off -1
            r.CreateCell(0);
            Assert.AreEqual(0, r.FirstCellNum);
            Assert.AreEqual(1, r.LastCellNum); // last cell # + 1
            Assert.AreEqual(1, r.PhysicalNumberOfCells);

            r.CreateCell(1);
            Assert.AreEqual(0, r.FirstCellNum);
            Assert.AreEqual(2, r.LastCellNum); // last cell # + 1
            Assert.AreEqual(2, r.PhysicalNumberOfCells);

            r.CreateCell(4);
            Assert.AreEqual(0, r.FirstCellNum);
            Assert.AreEqual(5, r.LastCellNum); // last cell # + 1
            Assert.AreEqual(3, r.PhysicalNumberOfCells);
        }

        /**
         * Data Tables - ptg 0x2
         */
        [TestMethod]
        public void Test44958()
        {
            HSSFWorkbook wb = OpenSample("44958.xls");
            Sheet s;
            Row r;
            Cell c;

            // Check the contents of the formulas

            // E4 to G9 of sheet 4 make up the table
            s = wb.GetSheet("OneVariable Table Completed");
            r = s.GetRow(3);
            c = r.GetCell(4);
            Assert.AreEqual(CellType.FORMULA, c.CellType);

            // TODO - Check the formula once tables and
            //  arrays are properly supported


            // E4 to H9 of sheet 5 make up the table
            s = wb.GetSheet("TwoVariable Table Example");
            r = s.GetRow(3);
            c = r.GetCell(4);
            Assert.AreEqual(CellType.FORMULA, c.CellType);

            // TODO - Check the formula once tables and
            //  arrays are properly supported
        }

        /**
         * 45322: Sheet.autoSizeColumn Assert.Fails when style.GetDataFormat() returns -1
         */
        [TestMethod]
        public void Test45322()
        {
            HSSFWorkbook wb = OpenSample("44958.xls");
            Sheet sh = wb.GetSheetAt(0);
            for (short i = 0; i < 30; i++) sh.AutoSizeColumn(i);
        }

        /**
         * We used to Add too many UncalcRecords to sheets
         *  with diagrams on. Don't any more
         */
        [TestMethod]
        public void Test45414()
        {
            HSSFWorkbook wb = OpenSample("WithThreeCharts.xls");
            wb.GetSheetAt(0).ForceFormulaRecalculation = (true);
            wb.GetSheetAt(1).ForceFormulaRecalculation = (false);
            wb.GetSheetAt(2).ForceFormulaRecalculation = (true);

            // Write out and back in again
            // This used to break
            HSSFWorkbook nwb = WriteOutAndReadBack(wb);

            // Check now set as it should be
            Assert.IsTrue(nwb.GetSheetAt(0).ForceFormulaRecalculation);
            Assert.IsFalse(nwb.GetSheetAt(1).ForceFormulaRecalculation);
            Assert.IsTrue(nwb.GetSheetAt(2).ForceFormulaRecalculation);
        }

        /**
         * Very hidden sheets not displaying as such
         */
        [TestMethod]
        public void Test45761()
        {
            HSSFWorkbook wb = OpenSample("45761.xls");
            Assert.AreEqual(3, wb.NumberOfSheets);

            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));
            Assert.IsTrue(wb.IsSheetHidden(1));
            Assert.IsFalse(wb.IsSheetVeryHidden(1));
            Assert.IsFalse(wb.IsSheetHidden(2));
            Assert.IsTrue(wb.IsSheetVeryHidden(2));

            // Change 0 to be very hidden, and re-load
            wb.SetSheetHidden(0, 2);

            HSSFWorkbook nwb = WriteOutAndReadBack(wb);

            Assert.IsFalse(nwb.IsSheetHidden(0));
            Assert.IsTrue(nwb.IsSheetVeryHidden(0));
            Assert.IsTrue(nwb.IsSheetHidden(1));
            Assert.IsFalse(nwb.IsSheetVeryHidden(1));
            Assert.IsFalse(nwb.IsSheetHidden(2));
            Assert.IsTrue(nwb.IsSheetVeryHidden(2));
        }

        ///**
        // * header / footer text too long
        // */
        //[TestMethod]
        //public void Test45777()
        //{
        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    Sheet s = wb.CreateSheet();

        //    String s248 = "";
        //    for (int i = 0; i < 248; i++)
        //    {
        //        s248 += "x";
        //    }
        //    String s249 = s248 + "1";
        //    String s250 = s248 + "12";
        //    String s251 = s248 + "123";
        //    Assert.AreEqual(248, s248.Length);
        //    Assert.AreEqual(249, s249.Length);
        //    Assert.AreEqual(250, s250.Length);
        //    Assert.AreEqual(251, s251.Length);


        //    // Try on headers
        //    s.Header.Center = (s248);
        //    Assert.AreEqual(254, s.Header.RawHeader.Length);
        //    WriteOutAndReadBack(wb);

        //    s.Header.Center = (s249);
        //    Assert.AreEqual(255, s.Header.RawHeader.Length);
        //    WriteOutAndReadBack(wb);

        //    try
        //    {
        //        s.Header.Center = (s250); // 256
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }

        //    try
        //    {
        //        s.Header.Center = (s251); // 257
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }


        //    // Now try on footers
        //    s.Footer.Center = (s248);
        //    Assert.AreEqual(254, s.Footer.RawFooter.Length);
        //    WriteOutAndReadBack(wb);

        //    s.Footer.Center = (s249);
        //    Assert.AreEqual(255, s.Footer.RawFooter.Length);
        //    WriteOutAndReadBack(wb);

        //    try
        //    {
        //        s.Footer.Center = (s250); // 256
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }

        //    try
        //    {
        //        s.Footer.Center = (s251); // 257
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }
        //}

        /**
         * Charts with long titles
         */
        [TestMethod]
        public void Test45784()
        {
            // This used to break
            HSSFWorkbook wb = OpenSample("45784.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
          * Cell background colours
          */
        [TestMethod]
        public void Test45492()
        {
            HSSFWorkbook wb = OpenSample("45492.xls");
            Sheet s = wb.GetSheetAt(0);
            Row r = s.GetRow(0);
            HSSFPalette p = wb.GetCustomPalette();

            Cell auto = r.GetCell(0);
            Cell grey = r.GetCell(1);
            Cell red = r.GetCell(2);
            Cell blue = r.GetCell(3);
            Cell green = r.GetCell(4);

            Assert.AreEqual(64, auto.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, auto.CellStyle.FillBackgroundColor);
            Assert.AreEqual("0:0:0", p.GetColor(64).GetHexString());

            Assert.AreEqual(22, grey.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, grey.CellStyle.FillBackgroundColor);
            Assert.AreEqual("C0C0:C0C0:C0C0", p.GetColor(22).GetHexString());

            Assert.AreEqual(10, red.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, red.CellStyle.FillBackgroundColor);
            Assert.AreEqual("FFFF:0:0", p.GetColor(10).GetHexString());

            Assert.AreEqual(12, blue.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, blue.CellStyle.FillBackgroundColor);
            Assert.AreEqual("0:0:FFFF", p.GetColor(12).GetHexString());

            Assert.AreEqual(11, green.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, green.CellStyle.FillBackgroundColor);
            Assert.AreEqual("0:FFFF:0", p.GetColor(11).GetHexString());
        }
        /**
 * ContinueRecord after EOF
 */
        [TestMethod]
        public void Test46137()
        {
            // This used to break
            HSSFWorkbook wb = OpenSample("46137.xls");
            Assert.AreEqual(7, wb.NumberOfSheets);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }

        /**
         * Odd POIFS blocks issue:
         * block[ 44 ] already removed from org.apache.poi.poifs.storage.BlockListImpl.remove
         */
        [TestMethod]
        public void Test45290()
        {
            HSSFWorkbook wb = OpenSample("45290.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
        }
        /**
 * In POI-2.5 user reported exception when parsing a name with a custom VBA function:
 *  =MY_VBA_FUNCTION("lskdjflsk")
 */
        [TestMethod]
        public void Test30070()
        {
            HSSFWorkbook wb = OpenSample("30070.xls"); //contains custom VBA function 'Commission'
            HSSFSheet sh = (HSSFSheet)wb.GetSheetAt(0);
            HSSFCell cell = (HSSFCell)sh.GetRow(0).GetCell(1);

            //B1 uses VBA in the formula
            Assert.AreEqual("Commission(A1)", cell.CellFormula);

            //name sales_1 refers to Commission(Sheet0!$A$1)
            int idx = wb.GetNameIndex("sales_1");
            Assert.IsTrue(idx != -1);

            HSSFName name = (HSSFName)wb.GetNameAt(idx);
            Assert.AreEqual("Commission(Sheet0!$A$1)", name.RefersToFormula);

        }

        /**
 * The link formulas which is referring to other books cannot be taken (the bug existed prior to POI-3.2)
 * Expected:
 *
 * [link_sub.xls]Sheet1!$A$1
 * [link_sub.xls]Sheet1!$A$2
 * [link_sub.xls]Sheet1!$A$3
 *
 * POI-3.1 output:
 *
 * Sheet1!$A$1
 * Sheet1!$A$2
 * Sheet1!$A$3
 *
 */
        [TestMethod]
        public void Test27364()
        {
            HSSFWorkbook wb = OpenSample("27364.xls");
            Sheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual("[link_sub.xls]Sheet1!$A$1", sheet.GetRow(0).GetCell(0).CellFormula);
            Assert.AreEqual("[link_sub.xls]Sheet1!$A$2", sheet.GetRow(1).GetCell(0).CellFormula);
            Assert.AreEqual("[link_sub.xls]Sheet1!$A$3", sheet.GetRow(2).GetCell(0).CellFormula);
        }

        /**
         * Similar to bug#27364:
         * HSSFCell.getCellFormula() fails with references to external workbooks
         */
        [TestMethod]
        public void Test31661()
        {
            HSSFWorkbook wb = OpenSample("31661.xls");
            Sheet sheet = wb.GetSheetAt(0);
            Cell cell = sheet.GetRow(11).GetCell(10); //K11
            Assert.AreEqual("+'[GM Budget.xls]8085.4450'!$B$2", cell.CellFormula);
        }

        /**
         * Incorrect handling of non-ISO 8859-1 characters in Windows ANSII Code Page 1252
         */
        [TestMethod]
        public void Test27394()
        {
            HSSFWorkbook wb = OpenSample("27394.xls");
            Assert.AreEqual("\u0161\u017E", wb.GetSheetName(0));
            Assert.AreEqual("\u0161\u017E\u010D\u0148\u0159", wb.GetSheetName(1));
            Sheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual("\u0161\u017E", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("\u0161\u017E\u010D\u0148\u0159", sheet.GetRow(1).GetCell(0).StringCellValue);
        }

        /**
         * Multiple calls of HSSFWorkbook.Write result in corrupted xls
         */
        [TestMethod]
        public void Test32191()
        {
            HSSFWorkbook wb = OpenSample("27394.xls");

            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);
            long size1 = out1.Length;
            out1.Close();

            out1 = new MemoryStream();
            wb.Write(out1);
            long size2 = out1.Length;
            out1.Close();

            Assert.AreEqual(size1, size2);
            out1 = new MemoryStream();
            wb.Write(out1);
            long size3 = out1.Length;
            out1.Close();
            Assert.AreEqual(size2, size3);

        }

        /**
         * java.io.IOException: block[ 0 ] already removed
         * (is an excel 95 file though)
         */
        [TestMethod]
        public void Test46904()
        {
            try
            {
                OpenSample("46904.xls");
                Assert.Fail();
            }
            catch (NPOI.HSSF.OldExcelFormatException e)
            {
                Assert.IsTrue(e.Message.StartsWith(
                        "The supplied spreadsheet seems to be Excel"
                ));
            }
        }

        /**
         * java.lang.NegativeArraySizeException reading long
         *  non-unicode data for a name record
         */
        [TestMethod]
        public void Test47034()
        {
            HSSFWorkbook wb = OpenSample("47034.xls");
            Assert.AreEqual(893, wb.NumberOfNames);
            Assert.AreEqual("Matthew\\Matthew11_1\\Matthew2331_1\\Matthew2351_1\\Matthew2361_1___lab", wb.GetNameName(300));
        }

        /**
         * HSSFRichTextString.Length returns negative for really long strings.
         * The Test file was created in OpenOffice 3.0 as Excel does not allow cell text longer than 32,767 characters
         */
        [TestMethod]
        public void Test46368()
        {
            HSSFWorkbook wb = OpenSample("46368.xls");
            Sheet s = wb.GetSheetAt(0);
            Cell cell1 = s.GetRow(0).GetCell(0);
            Assert.AreEqual(32770, cell1.StringCellValue.Length);

            Cell cell2 = s.GetRow(2).GetCell(0);
            Assert.AreEqual(32766, cell2.StringCellValue.Length);
        }

        /**
         * Short records on certain sheets with charts in them
         */
        [TestMethod]
        public void Test48180()
        {
            HSSFWorkbook wb = OpenSample("48180.xls");

            Sheet s = wb.GetSheetAt(0);
            Cell cell1 = s.GetRow(0).GetCell(0);
            Assert.AreEqual("test ", cell1.StringCellValue.ToString());

            Cell cell2 = s.GetRow(0).GetCell(1);
            Assert.AreEqual(1.0, cell2.NumericCellValue);
        }

        /**
         * POI 3.5 beta 7 can not read excel file contain list box (Form Control)  
         */
        [TestMethod]
        public void Test47701()
        {
            OpenSample("47701.xls");
        }

        /// <summary>
        /// http://npoi.codeplex.com/WorkItem/View.aspx?WorkItemId=5010
        /// </summary>
        [TestMethod]
        public void TestNPOIBug5010()
        {
            try
            {
                OpenSample("NPOIBug5010.xls");
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Contains("Unable to construct record instance"))
                {
                    throw new AssertFailedException("identified NPOI bug 5010");
                }
            }
        }
        /// <summary>
        /// http://npoi.codeplex.com/WorkItem/View.aspx?WorkItemId=5139
        /// </summary>
        [TestMethod]
        public void TestNPOIBug5139()
        {
            try
            {
                OpenSample("NPOIBug5139.xls");
            }
            catch (LeftoverDataException e)
            {
                if (e.Message.StartsWith("Initialisation of record 0x862"))
                {
                    throw new AssertFailedException("identified NPOI bug 5139");
                }
            }
        }

    }
}