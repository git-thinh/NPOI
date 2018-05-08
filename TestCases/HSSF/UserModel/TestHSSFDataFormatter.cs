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
    using System.Text;
    using NPOI.HSSF.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;
    /**
     * Unit Tests for HSSFDataFormatter.java
     *
     * @author James May (james dot may at fmr dot com)
     *
     */
    [TestClass]
    public class TestHSSFDataFormatter
    {

        private DataFormatter formatter;
        private HSSFWorkbook wb;


        public TestHSSFDataFormatter()
        {
            // Create the formatter to Test
            formatter = new DataFormatter();

            // Create a workbook to Test with
            wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = wb.CreateSheet();
            DataFormat format = wb.CreateDataFormat();

            // Create a row and put some cells in it
            Row row = sheet.CreateRow(0);

            // date value for July 8 1901 1:19 PM
            double dateNum = 555.555;

            //valid date formats -- all should have "Jul" in output
            String[] goodDatePatterns ={
			    "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
			    "mmm/d/yy\\ h:mm PM;@",
			    "mmmm/d/yy\\ h:mm;@",
			    "mmmm/d;@",
			    "mmmm/d/yy;@",
			    "mmm/dd/yy;@",
			    "[$-409]d\\-mmm;@",
			    "[$-409]d\\-mmm\\-yy;@",
			    "[$-409]dd\\-mmm\\-yy;@",
			    "[$-409]mmm\\-yy;@",
			    "[$-409]mmmm\\-yy;@",
			    "[$-409]mmmm\\ d\\,\\ yyyy;@",
			    "[$-409]mmm/d/yy\\ h:mm:ss;@",
			    "[$-409]mmmm/d/yy\\ h:mm:ss am;@",
			    "[$-409]mmmmm;@",
			    "[$-409]mmmmm\\-yy;@",
			    "mmmm/d/yyyy;@",
			    "[$-409]d\\-mmm\\-yyyy;@"
		    };

            // valid number formats
            String[] goodNumPatterns = {
				    "#,##0.0000",
				    "#,##0;[Red]#,##0",
				    "(#,##0.00_);(#,##0.00)",
				    "($#,##0.00_);[Red]($#,##0.00)",
				    "$#,##0.00",
				    "[$�-809]#,##0.00",
				    "[$�-2] #,##0.00",
				    "0000.00000%",
				    "0.000E+00",
				    "0.00E+00",
		    };

            // invalid date formats -- will throw exception in DecimalFormat ctor
            String[] badNumPatterns = {
				    "#,#$'#0.0000",
				    "'#','#ABC#0;##,##0",
				    "000 '123 4'5'6 000",
				    "#''0#0'1#10L16EE"
		    };

            // Create cells with good date patterns
            for (int i = 0; i < goodDatePatterns.Length; i++)
            {
                Cell cell = row.CreateCell(i);
                cell.SetCellValue(dateNum);
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat(goodDatePatterns[i]));
                cell.CellStyle = (cellStyle);
            }
            row = sheet.CreateRow(1);

            // Create cells with num patterns
            for (int i = 0; i < goodNumPatterns.Length; i++)
            {
                Cell cell = row.CreateCell(i);
                cell.SetCellValue(-1234567890.12345);
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat(goodNumPatterns[i]));
                cell.CellStyle = (cellStyle);
            }
            row = sheet.CreateRow(2);

            // Create cells with bad num patterns
            for (int i = 0; i < badNumPatterns.Length; i++)
            {
                Cell cell = row.CreateCell(i);
                cell.SetCellValue(1234567890.12345);
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat(badNumPatterns[i]));
                cell.CellStyle = (cellStyle);
            }

            // Built in formats

            { // Zip + 4 format
                row = sheet.CreateRow(3);
                Cell cell = row.CreateCell(0);
                cell.SetCellValue(123456789);
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat("00000-0000"));
                cell.CellStyle = (cellStyle);
            }

            { // Phone number format
                row = sheet.CreateRow(4);
                Cell cell = row.CreateCell(0);
                cell.SetCellValue(5551234567D);
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat("[<=9999999]###-####;(###) ###-####"));
                cell.CellStyle = (cellStyle);
            }

            { // SSN format
                row = sheet.CreateRow(5);
                Cell cell = row.CreateCell(0);
                cell.SetCellValue(444551234);
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat("000-00-0000"));
                cell.CellStyle = (cellStyle);
            }

            { // formula cell
                row = sheet.CreateRow(6);
                Cell cell = row.CreateCell(0);
                cell.SetCellType(NPOI.SS.UserModel.CellType.FORMULA);
                cell.CellFormula = ("SUM(12.25,12.25)/100");
                NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (format.GetFormat("##.00%;"));
                cell.CellStyle = (cellStyle);
            }
        }

        /**
         * Test getting formatted values from numeric and date cells.
         */
        [TestMethod]
        public void TestGetFormattedCellValueHSSFCell()
        {
            // Valid date formats -- cell values should be date formatted & not "555.555"
            Row row = wb.GetSheetAt(0).GetRow(0);
            IEnumerator it = row.GetCellEnumerator();
            log("==== VALID DATE FORMATS ====");
            while (it.MoveNext())
            {
                Cell cell = (Cell)it.Current;
                log(formatter.FormatCellValue(cell));

                // should not be equal to "555.555"
                Assert.IsTrue(!"555.555".Equals(formatter.FormatCellValue(cell)));

                // should contain "Jul" in the String
                string result = formatter.FormatCellValue(cell);
                Assert.IsTrue(result.IndexOf("Jul") > -1 || result.IndexOf("七月") > -1);
            }

            // Test number formats
            row = wb.GetSheetAt(0).GetRow(1);
            it = row.GetCellEnumerator();
            log("\n==== VALID NUMBER FORMATS ====");
            while (it.MoveNext())
            {
                Cell cell = (Cell)it.Current;
                log(formatter.FormatCellValue(cell));

                // should not be equal to "1234567890.12345"
                Assert.IsTrue(!"1234567890.12345".Equals(formatter.FormatCellValue(cell)));
            }

            // Test bad number formats
            row = wb.GetSheetAt(0).GetRow(2);
            it = row.GetCellEnumerator();
            log("\n==== INVALID NUMBER FORMATS ====");
            while (it.MoveNext())
            {
                Cell cell = (Cell)it.Current;
                log(formatter.FormatCellValue(cell));
                // should be equal to "1234567890.12345"
                Assert.AreEqual("1234567890.12345", formatter.FormatCellValue(cell));
            }

            // Test Zip+4 format
            row = wb.GetSheetAt(0).GetRow(3);
            Cell cell2 = row.GetCell(0);
            log("\n==== ZIP FORMAT ====");
            log(formatter.FormatCellValue(cell2));
            Assert.AreEqual("12345-6789", formatter.FormatCellValue(cell2));

            // Test phone number format
            row = wb.GetSheetAt(0).GetRow(4);
            cell2 = row.GetCell(0);
            log("\n==== PHONE FORMAT ====");
            log(formatter.FormatCellValue(cell2));
            Assert.AreEqual("(555) 123-4567", formatter.FormatCellValue(cell2));

            // Test SSN format
            row = wb.GetSheetAt(0).GetRow(5);
            cell2 = row.GetCell(0);
            log("\n==== SSN FORMAT ====");
            log(formatter.FormatCellValue(cell2));
            Assert.AreEqual("444-55-1234", formatter.FormatCellValue(cell2));

            // null Test-- null cell should result in empty String
            Assert.AreEqual(formatter.FormatCellValue(null), "");

            // null Test-- null cell should result in empty String
            Assert.AreEqual(formatter.FormatCellValue(null), "");
        }
        [TestMethod]
        public void TestGetFormattedCellValueHSSFCellHSSFFormulaEvaluator()
        {
            // Test formula format
            Row row = wb.GetSheetAt(0).GetRow(6);
            Cell cell = row.GetCell(0);
            log("\n==== FORMULA CELL ====");

            // first without a formula evaluator
            log(formatter.FormatCellValue(cell) + "\t (without evaluator)");
            Assert.AreEqual("SUM(12.25,12.25)/100", formatter.FormatCellValue(cell));

            // now with a formula evaluator
            HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(wb);
            log(formatter.FormatCellValue(cell, evaluator) + "\t\t\t (with evaluator)");
            Assert.AreEqual("24.50%", formatter.FormatCellValue(cell, evaluator));
        }

        /**
         * Test using a default number format. The format should be used when a
         * format pattern cannot be parsed by DecimalFormat.
         */
        [TestMethod]
        public void TestSetDefaultNumberFormat()
        {
            Row row = wb.GetSheetAt(0).GetRow(2);
            IEnumerator it = row.GetCellEnumerator();
            FormatBase defaultFormat = new DecimalFormat("Balance $#,#00.00 USD;Balance -$#,#00.00 USD");
            formatter.SetDefaultNumberFormat(defaultFormat);

            log("\n==== DEFAULT NUMBER FORMAT ====");

            Random rand = new Random((int)DateTime.Now.ToFileTime());
            while (it.MoveNext())
            {
                Cell cell = (Cell)it.Current;
                cell.SetCellValue(cell.NumericCellValue * rand.Next() / 1000000 - 1000);
                string result = formatter.FormatCellValue(cell);
                log(result);
                Assert.IsTrue(result.StartsWith("Balance "));
                Assert.IsTrue(result.EndsWith(" USD"));
            }
        }

        private static void log(String msg)
        {
            //if (false)
            //{ // successful Tests should be silent
            Console.WriteLine(msg);
            //}
        }
    }
}