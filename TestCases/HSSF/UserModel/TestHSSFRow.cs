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
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.SS.UserModel;

    /**
     * Test Row is okay.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class TestHSSFRow
    {

        public void TestLastAndFirstColumns()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            Row row = sheet.CreateRow(0);
            Assert.AreEqual(-1, row.FirstCellNum);
            Assert.AreEqual(-1, row.LastCellNum);

            row.CreateCell(2);
            Assert.AreEqual(2, row.FirstCellNum);
            Assert.AreEqual(3, row.LastCellNum);

            row.CreateCell(1);
            Assert.AreEqual(1, row.FirstCellNum);
            Assert.AreEqual(3, row.LastCellNum);

            // Check the exact case reported in 'bug' 43901 - notice that the cellNum is '0' based
            row.CreateCell(3);
            Assert.AreEqual(1, row.FirstCellNum);
            Assert.AreEqual(4, row.LastCellNum);
        }

        /**
         * Make sure that there is no cross-talk between rows especially with getFirstCellNum and getLastCellNum
         * This Test was Added in response to bug report 44987.
         */
        public void TestBoundsInMultipleRows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            Row rowA = sheet.CreateRow(0);

            rowA.CreateCell(10);
            rowA.CreateCell(5);
            Assert.AreEqual(5, rowA.FirstCellNum);
            Assert.AreEqual(11, rowA.LastCellNum);

            Row rowB = sheet.CreateRow(1);
            rowB.CreateCell(15);
            rowB.CreateCell(30);
            Assert.AreEqual(15, rowB.FirstCellNum);
            Assert.AreEqual(31, rowB.LastCellNum);

            Assert.AreEqual(5, rowA.FirstCellNum);
            Assert.AreEqual(11, rowA.LastCellNum);
            rowA.CreateCell(50);
            Assert.AreEqual(51, rowA.LastCellNum);

            Assert.AreEqual(31, rowB.LastCellNum);
        }
        [TestMethod]
        public void TestReMoveCell()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            Assert.AreEqual(-1, row.LastCellNum);
            Assert.AreEqual(-1, row.FirstCellNum);
            row.CreateCell(1);
            Assert.AreEqual(2, row.LastCellNum);
            Assert.AreEqual(1, row.FirstCellNum);
            row.CreateCell(3);
            Assert.AreEqual(4, row.LastCellNum);
            Assert.AreEqual(1, row.FirstCellNum);
            row.RemoveCell(row.GetCell(3));
            Assert.AreEqual(2, row.LastCellNum);
            Assert.AreEqual(1, row.FirstCellNum);
            row.RemoveCell(row.GetCell(1));
            Assert.AreEqual(-1, row.LastCellNum);
            Assert.AreEqual(-1, row.FirstCellNum);

            // all cells on this row have been Removed
            // so Check the row record actually Writes it out as 0's
            byte[] data = new byte[100];
            row.RowRecord.Serialize(0, data);
            Assert.AreEqual(0, data[6]);
            Assert.AreEqual(0, data[8]);

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = workbook.GetSheetAt(0);

            Assert.AreEqual(-1, sheet.GetRow(0).LastCellNum);
            Assert.AreEqual(-1, sheet.GetRow(0).FirstCellNum);
        }
        [TestMethod]
        public void TestMoveCell()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            Row row = sheet.CreateRow(0);
            Row rowB = sheet.CreateRow(1);

            Cell cellA2 = rowB.CreateCell(0);
            Assert.AreEqual(0, rowB.FirstCellNum);
            Assert.AreEqual(0, rowB.FirstCellNum);

            Assert.AreEqual(-1, row.LastCellNum);
            Assert.AreEqual(-1, row.FirstCellNum);
            Cell cellB2 = row.CreateCell(1);
            Cell cellB3 = row.CreateCell(2);
            Cell cellB4 = row.CreateCell(3);

            Assert.AreEqual(1, row.FirstCellNum);
            Assert.AreEqual(4, row.LastCellNum);

            // Try to move to somewhere else that's used
            try
            {
                row.MoveCell(cellB2, (short)3);
                Assert.Fail("ArgumentException should have been thrown");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
            }

            // Try to move one off a different row
            try
            {
                row.MoveCell(cellA2, (short)3);
                Assert.Fail("ArgumentException should have been thrown");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
            }

            // Move somewhere spare
            Assert.IsNotNull(row.GetCell(1));
            row.MoveCell(cellB2, (short)5);
            Assert.IsNull(row.GetCell(1));
            Assert.IsNotNull(row.GetCell(5));

            Assert.AreEqual(5, cellB2.ColumnIndex);
            Assert.AreEqual(2, row.FirstCellNum);
            Assert.AreEqual(6, row.LastCellNum);
        }
        [TestMethod]
        public void TestRowBounds()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            //Test low row bound
            sheet.CreateRow(0);
            //Test low row bound exception
            try
            {
                sheet.CreateRow(-1);
                Assert.Fail("IndexOutOfBoundsException should have been thrown");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                Assert.AreEqual("Invalid row number (-1) outside allowable range (0..65535)", e.Message);
            }

            //Test high row bound
            sheet.CreateRow(65535);
            //Test high row bound exception
            try
            {
                sheet.CreateRow(65536);
                Assert.Fail("IndexOutOfBoundsException should have been thrown");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                Assert.AreEqual("Invalid row number (65536) outside allowable range (0..65535)", e.Message);
            }
        }

        /**
         * Prior to patch 43901, POI was producing files with the wrong last-column
         * number on the row
         */
        [TestMethod]
        public void TestLastCellNumIsCorrectAfterAddCell_bug43901()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = book.CreateSheet("Test");
            Row row = sheet.CreateRow(0);

            // New row has last col -1
            Assert.AreEqual(-1, row.LastCellNum);
            if (row.LastCellNum == 0)
            {
                Assert.Fail("Identified bug 43901");
            }

            // Create two cells, will return one higher
            //  than that for the last number
            row.CreateCell(0);
            Assert.AreEqual(1, row.LastCellNum);
            row.CreateCell(255);
            Assert.AreEqual(256, row.LastCellNum);
        }

        /**
         * Tests for the missing/blank cell policy stuff
         */
        [TestMethod]
        public void TestGetCellPolicy()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = book.CreateSheet("Test");
            Row row = sheet.CreateRow(0);

            // 0 -> string
            // 1 -> num
            // 2 missing
            // 3 missing
            // 4 -> blank
            // 5 -> num
            row.CreateCell(0).SetCellValue(new HSSFRichTextString("Test"));
            row.CreateCell(1).SetCellValue(3.2);
            row.CreateCell(4, NPOI.SS.UserModel.CellType.BLANK);
            row.CreateCell(5).SetCellValue(4);

            // First up, no policy given, uses default
            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, row.GetCell(0).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(1).CellType);
            Assert.AreEqual(null, row.GetCell(2));
            Assert.AreEqual(null, row.GetCell(3));
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BLANK, row.GetCell(4).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(5).CellType);

            // RETURN_NULL_AND_BLANK - same as default
            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, row.GetCell(0, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);
            Assert.AreEqual(null, row.GetCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            Assert.AreEqual(null, row.GetCell(3, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BLANK, row.GetCell(4, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(5, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);

            // RETURN_BLANK_AS_NULL - nearly the same
            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, row.GetCell(0, MissingCellPolicy.RETURN_BLANK_AS_NULL).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(1, MissingCellPolicy.RETURN_BLANK_AS_NULL).CellType);
            Assert.AreEqual(null, row.GetCell(2, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.AreEqual(null, row.GetCell(3, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.AreEqual(null, row.GetCell(4, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(5, MissingCellPolicy.RETURN_BLANK_AS_NULL).CellType);

            // CREATE_NULL_AS_BLANK - Creates as needed
            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, row.GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BLANK, row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BLANK, row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BLANK, row.GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);

            // Check Created ones get the right column
            Assert.AreEqual(0, row.GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            Assert.AreEqual(1, row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            Assert.AreEqual(2, row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            Assert.AreEqual(3, row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            Assert.AreEqual(4, row.GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            Assert.AreEqual(5, row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);


            // Now change the cell policy on the workbook, Check
            //  that that is now used if no policy given
            book.MissingCellPolicy=(MissingCellPolicy.RETURN_BLANK_AS_NULL);

            Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, row.GetCell(0).CellType);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(1).CellType);
            Assert.AreEqual(null, row.GetCell(2));
            Assert.AreEqual(null, row.GetCell(3));
            Assert.AreEqual(null, row.GetCell(4));
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, row.GetCell(5).CellType);
        }
        [TestMethod]
        public void TestRowHeight()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            Row row1 = sheet.CreateRow(0);

            Assert.AreEqual(0xFF, row1.Height);
            Assert.AreEqual(sheet.DefaultRowHeight, row1.Height);

            Row row2 = sheet.CreateRow(1);
            row2.Height=((short)400);

            Assert.AreEqual(400, row2.Height);

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = workbook.GetSheetAt(0);

            row1 = sheet.GetRow(0);
            Assert.AreEqual(0xFF, row1.Height);
            Assert.AreEqual(sheet.DefaultRowHeight, row1.Height);

            row2 = sheet.GetRow(1);
            Assert.AreEqual(400, row2.Height);
        }

    }
}