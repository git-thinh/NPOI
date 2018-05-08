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
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.UserModel;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    /**
     * Tests TestHSSFCellComment.
     *
     * @author  Yegor Kozlov
     */
    [TestClass]
    public class TestHSSFComment
    {

        /**
         * Test that we can Create cells and Add comments to it.
         */
        //[TestMethod]
        //public void TestWriteComments()
        //{
        //    String cellText = "Hello, World";
        //    String commentText = "We can set comments in POI";
        //    String commentAuthor = "Apache Software Foundation";
        //    int cellRow = 3;
        //    short cellColumn = 1;

        //    HSSFWorkbook wb = new HSSFWorkbook();

        //    NPOI.SS.UserModel.Sheet sheet = wb.CreateSheet();

        //    Cell cell = sheet.CreateRow(cellRow).CreateCell(cellColumn);
        //    cell.SetCellValue(new HSSFRichTextString(cellText));
        //    Assert.IsNull(cell.CellComment);

        //    HSSFPalette patr = (HSSFPalette)sheet.CreateDrawingPatriarch();
        //    HSSFClientAnchor anchor = new HSSFClientAnchor();
        //    anchor.SetAnchor((short)4, 2, 0, 0, (short)6, 5, 0, 0);
        //    Comment comment = patr.CreateComment(anchor);
        //    HSSFRichTextString string1 = new HSSFRichTextString(commentText);
        //    comment.String = string1;
        //    comment.Author = (commentAuthor);
        //    cell.CellComment = (comment);

        //    //verify our settings
        //    Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_COMMENT, ((HSSFSimpleShape)comment).ShapeType);
        //    Assert.AreEqual(commentAuthor, comment.Author);
        //    Assert.AreEqual(commentText, comment.String.String);
        //    Assert.AreEqual(cellRow, comment.Row);
        //    Assert.AreEqual(cellColumn, comment.Column);

        //    //serialize the workbook and read it again
        //    MemoryStream out1 = new MemoryStream();
        //    wb.Write(out1);
        //    out1.Close();

        //    wb = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
        //    sheet = wb.GetSheetAt(0);
        //    cell = sheet.GetRow(cellRow).GetCell(cellColumn);
        //    comment = cell.CellComment;

        //    Assert.IsNotNull(comment);
        //    Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_COMMENT, ((HSSFSimpleShape)comment).ShapeType);
        //    Assert.AreEqual(commentAuthor, comment.Author);
        //    Assert.AreEqual(commentText, comment.String.String);
        //    Assert.AreEqual(cellRow, comment.Row);
        //    Assert.AreEqual(cellColumn, comment.Column);
        //}

        /**
         * Test that we can read cell comments from an existing workbook.
         */
        [TestMethod]
        public void TestReadComments()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithComments.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);

            Cell cell;
            Row row;
            Comment comment;

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(0);
                comment = cell.CellComment;
                Assert.IsNull(comment, "Cells in the first column are not commented");
                Assert.IsNull(sheet.GetCellComment(rownum, 0));
            }

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(1);
                comment = cell.CellComment;
                Assert.IsNotNull(comment, "Cells in the second column have comments");
                Assert.IsNotNull(sheet.GetCellComment(rownum, 1), "Cells in the second column have comments");

                Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_COMMENT, ((HSSFSimpleShape)comment).ShapeType);
                Assert.AreEqual("Yegor Kozlov", comment.Author);
                Assert.IsFalse(comment.String.String == string.Empty, "cells in the second column have not empyy notes");
                Assert.AreEqual(rownum, comment.Row);
                Assert.AreEqual(cell.ColumnIndex, comment.Column);
            }
        }

        /**
         * Test that we can modify existing cell comments
         */
        [TestMethod]
        public void TestModifyComments()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithComments.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);

            Cell cell;
            Row row;
            Comment comment;

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(1);
                comment = cell.CellComment;
                comment.Author = ("Mofified[" + rownum + "] by Yegor");
                comment.String = (new HSSFRichTextString("Modified comment at row " + rownum));
            }

            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);
            out1.Close();

            wb = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            sheet = wb.GetSheetAt(0);

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(1);
                comment = cell.CellComment;

                Assert.AreEqual("Mofified[" + rownum + "] by Yegor", comment.Author);
                Assert.AreEqual("Modified comment at row " + rownum, comment.String.String);
            }

        }
        [TestMethod]
        public void TestDeleteComments()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithComments.xls");
            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);

            // Zap from rows 1 and 3
            Assert.IsNotNull(sheet.GetRow(0).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(1).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(2).GetCell(1).CellComment);

            sheet.GetRow(0).GetCell(1).RemoveCellComment();
            sheet.GetRow(2).GetCell(1).CellComment = (null);

            // Check gone so far
            Assert.IsNull(sheet.GetRow(0).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(1).GetCell(1).CellComment);
            Assert.IsNull(sheet.GetRow(2).GetCell(1).CellComment);

            // Save and re-load
            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);
            out1.Close();
            wb = new HSSFWorkbook(new MemoryStream(out1.ToArray()));

            // Check
            Assert.IsNull(sheet.GetRow(0).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(1).GetCell(1).CellComment);
            Assert.IsNull(sheet.GetRow(2).GetCell(1).CellComment);

            //        FileOutputStream fout = new FileOutputStream("/tmp/c.xls");
            //        wb.Write(fout);
            //        fout.Close();
        }

        /**
 *  HSSFCell#findCellComment should NOT rely on the order of records
 * when matching cells and their cell comments. The correct algorithm is to map
 */
        [TestMethod]
        public void Test47924()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("47924.xls");
            Sheet sheet = wb.GetSheetAt(0);
            Cell cell;
            Comment comment;

            cell = sheet.GetRow(0).GetCell(0);
            comment = cell.CellComment;
            Assert.AreEqual("a1", comment.String.String);

            cell = sheet.GetRow(1).GetCell(0);
            comment = cell.CellComment;
            Assert.AreEqual("a2", comment.String.String);

            cell = sheet.GetRow(2).GetCell(0);
            comment = cell.CellComment;
            Assert.AreEqual("a3", comment.String.String);

            cell = sheet.GetRow(2).GetCell(2);
            comment = cell.CellComment;
            Assert.AreEqual("c3", comment.String.String);

            cell = sheet.GetRow(4).GetCell(1);
            comment = cell.CellComment;
            Assert.AreEqual("b5", comment.String.String);

            cell = sheet.GetRow(5).GetCell(2);
            comment = cell.CellComment;
            Assert.AreEqual("c6", comment.String.String);
        }
    }
}
