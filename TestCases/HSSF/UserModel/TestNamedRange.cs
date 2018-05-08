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
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * 
     * @author ROMANL
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Danny Mui (danny at muibros.com)
     * @author Amol S. Deshmukh &lt; amol at ap ache dot org &gt; 
     */
    [TestClass]
    public class TestNamedRange
    {

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        /** Test of TestCase method, of class Test.RangeTest. */
        [TestMethod]
        public void TestNamedRange1()
        {
            HSSFWorkbook wb = OpenSample("Simple.xls");

            //Creating new Named Range
            NPOI.SS.UserModel.Name newNamedRange = wb.CreateName();

            //Getting Sheet Name for the reference
            String sheetName = wb.GetSheetName(0);

            //Setting its name
            newNamedRange.NameName = ("RangeTest");
            //Setting its reference
            newNamedRange.RefersToFormula = (sheetName + "!$D$4:$E$8");

            //Getting NAmed Range
            NPOI.SS.UserModel.Name namedRange1 = wb.GetNameAt(0);
            //Getting it sheet name
            sheetName = namedRange1.SheetName;
            //Getting its reference
            String referece = namedRange1.RefersToFormula;

            // sanity Check
            SanityChecker c = new SanityChecker();
            c.CheckHSSFWorkbook(wb);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            NPOI.SS.UserModel.Name nm = wb.GetNameAt(wb.GetNameIndex("RangeTest"));
            Assert.IsTrue("RangeTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.AreEqual(wb.GetSheetName(0) + "!$D$4:$E$8", nm.RefersToFormula);
        }

        /**
         * Reads an excel file already containing a named range.
         * 
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/show_bug.cgi?id=9632" target="_bug">#9632</a>
         */
        [TestMethod]
        public void TestNamedRead()
        {
            HSSFWorkbook wb = OpenSample("namedinput.xls");

            //Get index of the namedrange with the name = "NamedRangeName" , which was defined in input.xls as A1:D10
            int NamedRangeIndex = wb.GetNameIndex("NamedRangeName");

            //Getting NAmed Range
            NPOI.SS.UserModel.Name namedRange1 = wb.GetNameAt(NamedRangeIndex);
            String sheetName = wb.GetSheetName(0);

            //Getting its reference
            String reference = namedRange1.RefersToFormula;

            Assert.AreEqual(sheetName + "!$A$1:$D$10", reference);

            NPOI.SS.UserModel.Name namedRange2 = wb.GetNameAt(1);

            Assert.AreEqual(sheetName + "!$D$17:$G$27", namedRange2.RefersToFormula);
            Assert.AreEqual(namedRange2.NameName, "SecondNamedRange");
        }

        /**
         * Reads an excel file already containing a named range and updates it
         * 
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/show_bug.cgi?id=16411" target="_bug">#16411</a>
         */
        [TestMethod]
        public void TestNamedReadModify()
        {
            HSSFWorkbook wb = OpenSample("namedinput.xls");

            NPOI.SS.UserModel.Name name = wb.GetNameAt(0);
            String sheetName = wb.GetSheetName(0);

            Assert.AreEqual(sheetName + "!$A$1:$D$10", name.RefersToFormula);

            name = wb.GetNameAt(1);
            String newReference = sheetName + "!$A$1:$C$36";

            name.RefersToFormula = (newReference);
            Assert.AreEqual(newReference, name.RefersToFormula);
        }

        /**
         * Test that multiple named ranges can be Added written and read
         */
        [TestMethod]
        public void TestMultipleNamedWrite()
        {
            HSSFWorkbook wb = new HSSFWorkbook();


            wb.CreateSheet("TestSheet1");
            String sheetName = wb.GetSheetName(0);

            Assert.AreEqual("TestSheet1", sheetName);

            //Creating new Named Range
            NPOI.SS.UserModel.Name newNamedRange = wb.CreateName();

            newNamedRange.NameName = ("RangeTest");
            newNamedRange.RefersToFormula = (sheetName + "!$D$4:$E$8");

            //Creating another new Named Range
            NPOI.SS.UserModel.Name newNamedRange2 = wb.CreateName();

            newNamedRange2.NameName = ("AnotherTest");
            newNamedRange2.RefersToFormula = (sheetName + "!$F$1:$G$6");


            NPOI.SS.UserModel.Name namedRange1 = wb.GetNameAt(0);
            String referece = namedRange1.RefersToFormula;

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            NPOI.SS.UserModel.Name nm = wb.GetNameAt(wb.GetNameIndex("RangeTest"));
            Assert.IsTrue("RangeTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.IsTrue((wb.GetSheetName(0) + "!$D$4:$E$8").Equals(nm.RefersToFormula), "Reference is " + nm.RefersToFormula);

            nm = wb.GetNameAt(wb.GetNameIndex("AnotherTest"));
            Assert.IsTrue("AnotherTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.IsTrue(newNamedRange2.RefersToFormula.Equals(nm.RefersToFormula), "Reference is " + nm.RefersToFormula);
        }

        /**
         * Test case provided by czhang@cambian.com (Chun Zhang)
         * 
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/show_bug.cgi?id=13775" target="_bug">#13775</a>
         */
        [TestMethod]
        public void TestMultiNamedRange()
        {

            // Create a new workbook
            HSSFWorkbook wb = new HSSFWorkbook();


            // Create a worksheet 'sheet1' in the new workbook
            wb.CreateSheet();
            wb.SetSheetName(0, "sheet1");

            // Create another worksheet 'sheet2' in the new workbook
            wb.CreateSheet();
            wb.SetSheetName(1, "sheet2");

            // Create a new named range for worksheet 'sheet1'
            NPOI.SS.UserModel.Name namedRange1 = wb.CreateName();

            // Set the name for the named range for worksheet 'sheet1'
            namedRange1.NameName = ("RangeTest1");

            // Set the reference for the named range for worksheet 'sheet1'
            namedRange1.RefersToFormula = ("sheet1" + "!$A$1:$L$41");

            // Create a new named range for worksheet 'sheet2'
            NPOI.SS.UserModel.Name namedRange2 = wb.CreateName();

            // Set the name for the named range for worksheet 'sheet2'
            namedRange2.NameName = ("RangeTest2");

            // Set the reference for the named range for worksheet 'sheet2'
            namedRange2.RefersToFormula = ("sheet2" + "!$A$1:$O$21");

            // Write the workbook to a file
            // Read the Excel file and verify its content
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            NPOI.SS.UserModel.Name nm1 = wb.GetNameAt(wb.GetNameIndex("RangeTest1"));
            Assert.IsTrue("RangeTest1".Equals(nm1.NameName), "Name is " + nm1.NameName);
            Assert.IsTrue((wb.GetSheetName(0) + "!$A$1:$L$41").Equals(nm1.RefersToFormula), "Reference is " + nm1.RefersToFormula);

            NPOI.SS.UserModel.Name nm2 = wb.GetNameAt(wb.GetNameIndex("RangeTest2"));
            Assert.IsTrue("RangeTest2".Equals(nm2.NameName), "Name is " + nm2.NameName);
            Assert.IsTrue((wb.GetSheetName(1) + "!$A$1:$O$21").Equals(nm2.RefersToFormula), "Reference is " + nm2.RefersToFormula);
        }
        [TestMethod]
        public void TestUnicodeNamedRange()
        {
            HSSFWorkbook workBook = new HSSFWorkbook();
            workBook.CreateSheet("Test");
            NPOI.SS.UserModel.Name name = workBook.CreateName();
            name.NameName = ("\u03B1");
            name.RefersToFormula = ("Test!$D$3:$E$8");


            HSSFWorkbook workBook2 = HSSFTestDataSamples.WriteOutAndReadBack(workBook);
            NPOI.SS.UserModel.Name name2 = workBook2.GetNameAt(0);

            Assert.AreEqual("\u03B1", name2.NameName);
            Assert.AreEqual("Test!$D$3:$E$8", name2.RefersToFormula);
        }

        /**
         * Test to see if the print areas can be retrieved/Created in memory
         */
        [TestMethod]
        public void TestSinglePrintArea()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea);
        }

        /**
         * For Convenience, dont force sheet names to be used
         */
        [TestMethod]
        public void TestSinglePrintAreaWOSheet()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!" + reference, retrievedPrintArea);
        }

        /**
         * Test to see if the print area can be retrieved from an excel Created file
         */
        [TestMethod]
        public void TestPrintAreaFileRead()
        {
            HSSFWorkbook workbook = OpenSample("SimpleWithPrintArea.xls");

            String sheetName = workbook.GetSheetName(0);
            String reference = sheetName + "!$A$1:$C$5";

            Assert.AreEqual(reference, workbook.GetPrintArea(0));
        }

        /**
         * Test to see if the print area made it to the file
         */
        [TestMethod]
        public void TestPrintAreaFile()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);


            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            String retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea, "References Match");
        }

        /**
         * Test to see if multiple print areas made it to the file
         */
        [TestMethod]
        public void TestMultiplePrintAreaFile()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();

            workbook.CreateSheet("Sheet1");
            workbook.CreateSheet("Sheet2");
            workbook.CreateSheet("Sheet3");
            String reference1 = "$A$1:$B$1";
            String reference2 = "$B$2:$D$5";
            String reference3 = "$D$2:$F$5";

            workbook.SetPrintArea(0, reference1);
            workbook.SetPrintArea(1, reference2);
            workbook.SetPrintArea(2, reference3);

            //Check Created print areas
            String retrievedPrintArea;

            retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 1)");
            Assert.AreEqual("Sheet1!"+reference1, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(1);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 2)");
            Assert.AreEqual("Sheet2!" + reference2, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(2);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 3)");
            Assert.AreEqual("Sheet3!" + reference3, retrievedPrintArea);

            // Check print areas after re-reading workbook
            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 1)");
            Assert.AreEqual("Sheet1!" + reference1, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(1);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 2)");
            Assert.AreEqual("Sheet2!" + reference2, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(2);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 3)");
            Assert.AreEqual("Sheet3!" + reference3, retrievedPrintArea);
        }

        /**
         * Tests the setting of print areas with coordinates (Row/Column designations)
         *
         */
        [TestMethod]
        public void TestPrintAreaCoords()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            String reference = sheetName + "!$A$1:$B$1";
            workbook.SetPrintArea(0, 0, 1, 0, 0);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea);
        }


        /**
         * Tests the parsing of union area expressions, and re-display in the presence of sheet names
         * with special characters.
         */
        [TestMethod]
        public void TestPrintAreaUnion()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);


            String reference = "$A$1:$B$1,$D$1:$F$2";
            workbook.SetPrintArea(0, reference);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'Test Print Area'!$A$1:$B$1,'Test Print Area'!$D$1:$F$2", retrievedPrintArea);
        }

        /**
         * Verifies an existing print area is deleted
         *
         */
        [TestMethod]
        public void TestPrintAreaRemove()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            String reference = sheetName + "!$A$1:$B$1";
            workbook.SetPrintArea(0, 0, 1, 0, 0);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");

            workbook.RemovePrintArea(0);
            Assert.IsNull(workbook.GetPrintArea(0), "PrintArea was not Removed");
        }

        /**
         * Verifies correct functioning for "single cell named range" (aka "named cell")
         */
        [TestMethod]
        public void TestNamedCell_1()
        {

            // setup for this Testcase
            String sheetName = "Test Named Cell";
            String cellName = "A name for a named cell";
            String cellValue = "TEST Value";
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = wb.CreateSheet(sheetName);
            sheet.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString(cellValue));

            // Create named range for a single cell using areareference
            NPOI.SS.UserModel.Name namedCell = wb.CreateName();
            namedCell.NameName = (cellName);
            String reference = "'" + sheetName + "'" + "!A1:A1";
            namedCell.RefersToFormula = (reference);

            // retrieve the newly Created named range
            int namedCellIdx = wb.GetNameIndex(cellName);
            NPOI.SS.UserModel.Name aNamedCell = wb.GetNameAt(namedCellIdx);
            Assert.IsNotNull(aNamedCell);

            // retrieve the cell at the named range and Test its contents
            AreaReference aref = new AreaReference(aNamedCell.RefersToFormula);
            Assert.IsTrue(aref.IsSingleCell, "Should be exactly 1 cell in the named cell :'" + cellName + "'");

            CellReference cref = aref.FirstCell;
            Assert.IsNotNull(cref);
            NPOI.SS.UserModel.Sheet s = wb.GetSheet(cref.SheetName);
            Assert.IsNotNull(s);
            Row r = sheet.GetRow(cref.Row);
            Cell c = r.GetCell(cref.Col);
            String contents = c.RichStringCellValue.String;
            Assert.AreEqual(contents, cellValue, "Contents of cell retrieved by its named reference");
        }

        /**
         * Verifies correct functioning for "single cell named range" (aka "named cell")
         */
        [TestMethod]
        public void TestNamedCell_2()
        {

            // setup for this Testcase
            String sname = "TestSheet", cname = "TestName", cvalue = "TestVal";
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = wb.CreateSheet(sname);
            sheet.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString(cvalue));

            // Create named range for a single cell using cellreference
            NPOI.SS.UserModel.Name namedCell = wb.CreateName();
            namedCell.NameName = (cname);
            String reference = sname + "!A1";
            namedCell.RefersToFormula = (reference);

            // retrieve the newly Created named range
            int namedCellIdx = wb.GetNameIndex(cname);
            NPOI.SS.UserModel.Name aNamedCell = wb.GetNameAt(namedCellIdx);
            Assert.IsNotNull(aNamedCell);

            // retrieve the cell at the named range and Test its contents
            CellReference cref = new CellReference(aNamedCell.RefersToFormula);
            Assert.IsNotNull(cref);
            NPOI.SS.UserModel.Sheet s = wb.GetSheet(cref.SheetName);
            Row r = sheet.GetRow(cref.Row);
            Cell c = r.GetCell(cref.Col);
            String contents = c.RichStringCellValue.String;
            Assert.AreEqual(contents, cvalue, "Contents of cell retrieved by its named reference");
        }
        [TestMethod]
        public void TestDeletedReference()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("24207.xls");
            Assert.AreEqual(2, wb.NumberOfNames);

            NPOI.SS.UserModel.Name name1 = wb.GetNameAt(0);
            Assert.AreEqual("a", name1.NameName);
            Assert.AreEqual("Sheet1!$A$1", name1.RefersToFormula);
            AreaReference ref1 = new AreaReference(name1.RefersToFormula);
            //Assert.IsTrue(true, "Successfully constructed first reference");

            NPOI.SS.UserModel.Name name2 = wb.GetNameAt(1);
            Assert.AreEqual("b", name2.NameName);
            Assert.AreEqual("Sheet1!#REF!", name2.RefersToFormula);
            Assert.IsTrue(name2.IsDeleted);
            try
            {
                AreaReference ref2 = new AreaReference(name2.RefersToFormula);
                Assert.Fail("attempt to supply an invalid reference to AreaReference constructor results in exception");
            }
            catch (IndexOutOfRangeException e)
            { // TODO - use a different exception for this condition
                // expected during successful Test
            }
        }
        [TestMethod]
        public void TestRepeatingRowsAndColumsNames()
        {
            // First Test that setting RR&C for same sheet more than once only Creates a 
            // single  Print_Titles built-in record
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = wb.CreateSheet("FirstSheet");

            // set repeating rows and columns twice for the first sheet
            for (int i = 0; i < 2; i++)
            {
                wb.SetRepeatingRowsAndColumns(0, 0, 0, 0, 3 - 1);
                sheet.CreateFreezePane(0, 3);
            }
            Assert.AreEqual(1, wb.NumberOfNames);
            NPOI.SS.UserModel.Name nr1 = wb.GetNameAt(0);

            Assert.AreEqual("Print_Titles", nr1.NameName);
            if (false)
            {
                // 	TODO - full column references not rendering properly, absolute markers not present either
                Assert.AreEqual("FirstSheet!$A:$A,FirstSheet!$1:$3", nr1.RefersToFormula);
            }
            else
            {
                Assert.AreEqual("FirstSheet!A:A,FirstSheet!$A$1:$IV$3", nr1.RefersToFormula);
            }

            // Save and re-Open
            HSSFWorkbook nwb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            Assert.AreEqual(1, nwb.NumberOfNames);
            nr1 = nwb.GetNameAt(0);

            Assert.AreEqual("Print_Titles", nr1.NameName);
            Assert.AreEqual("FirstSheet!A:A,FirstSheet!$A$1:$IV$3", nr1.RefersToFormula);

            // Check that setting RR&C on a second sheet causes a new Print_Titles built-in
            // name to be Created
            sheet = nwb.CreateSheet("SecondSheet");
            nwb.SetRepeatingRowsAndColumns(1, 1, 2, 0, 0);

            Assert.AreEqual(2, nwb.NumberOfNames);
            NPOI.SS.UserModel.Name nr2 = nwb.GetNameAt(1);

            Assert.AreEqual("Print_Titles", nr2.NameName);
            Assert.AreEqual("SecondSheet!B:C,SecondSheet!$A$1:$IV$1", nr2.RefersToFormula);

            if (false)
            {
                // In case you fancy Checking in excel, to ensure it
                //  won't complain about the file now
                try
                {
                    string tmppath = NPOI.Util.TempFile.GetTempFilePath("POI-45126-", ".xls");
                    FileStream fout = new FileStream(tmppath, FileMode.OpenOrCreate);
                    nwb.Write(fout);
                    fout.Close();
                    Console.WriteLine("Check out " + Path.GetFullPath(tmppath));
                }
                catch (IOException)
                {
                    throw;
                }
            }
        }
    }
}