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

    using TestCases.HSSF;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.UserModel;
    /**
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Avik Sengupta
     */
    [TestClass]
    public class TestFormulas
    {
        public TestFormulas()
        { 
        }

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        /**
         * Add 1+1 -- WHoohoo!
         */
        [TestMethod]
        public void TestBasicAddIntegers()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;

            //get our minimum values
            r = s.CreateRow(1);
            c = r.CreateCell(1);
            c.CellFormula = (1 + "+" + 1);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(1);
            c = r.GetCell(1);

            Assert.IsTrue("1+1".Equals(c.CellFormula), "Formula is as expected");
        }

        /**
         * Add various integers
         */
        [TestMethod]
        public void TestAddIntegers()
        {
            BinomialOperator("+");
        }

        /**
         * Multiply various integers
         */
        [TestMethod]
        public void TestMultplyIntegers()
        {
            BinomialOperator("*");
        }

        /**
         * Subtract various integers
         */
        [TestMethod]
        public void TestSubtractIntegers()
        {
            BinomialOperator("-");
        }

        /**
         * Subtract various integers
         */
        [TestMethod]
        public void TestDivideIntegers()
        {
            BinomialOperator("/");
        }

        /**
         * Exponentialize various integers;
         */
        [TestMethod]
        public void TestPowerIntegers()
        {
            BinomialOperator("^");
        }

        /**
         * Concatenate two numbers 1&2 = 12
         */
        [TestMethod]
        public void TestConcatIntegers()
        {
            BinomialOperator("&");
        }

        /**
         * Tests 1*2+3*4
         */
        [TestMethod]
        public void TestOrderOfOperationsMultiply()
        {
            OrderTest("1*2+3*4");
        }

        /**
         * Tests 1*2+3^4
         */
        [TestMethod]
        public void TestOrderOfOperationsPower()
        {
            OrderTest("1*2+3^4");
        }

        /**
         * Tests that parenthesis are obeyed
         */
        [TestMethod]
        public void TestParenthesis()
        {
            OrderTest("(1*3)+2+(1+2)*(3^4)^5");
        }
        [TestMethod]
        public void TestReferencesOpr()
        {
            String[] operation = new String[] {
                            "+", "-", "*", "/", "^", "&"
                           };
            for (int k = 0; k < operation.Length; k++)
            {
                OperationRefTest(operation[k]);
            }
        }

        /**
         * Tests creating a file with floating point in a formula.
         *
         */
        [TestMethod]
        public void TestFloat()
        {
            floatTest("*");
            floatTest("/");
        }

        private static void floatTest(String operator1)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;

            //get our minimum values

            r = s.CreateRow(0);
            c = r.CreateCell(1);
            c.CellFormula = ("" + float.MinValue + operator1 + float.MinValue);

            for (int x = 1; x < short.MaxValue && x > 0; x = (short)(x * 2))
            {
                r = s.CreateRow(x);

                for (int y = 1; y < 256 && y > 0; y = (short)(y + 2))
                {

                    c = r.CreateCell(y);
                    c.CellFormula = ("" + x + "." + y + operator1 + y + "." + x);


                }
            }
            if (s.LastRowNum < short.MaxValue)
            {
                r = s.CreateRow(0);
                c = r.CreateCell(0);
                c.CellFormula = ("" + float.MaxValue + operator1 + float.MaxValue);
            }
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            floatVerify(operator1, wb);
        }

        private static void floatVerify(String operator1, HSSFWorkbook wb)
        {

            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);

            // don't know how to Check correct result .. for the moment, we just verify that the file can be read.

            for (int x = 1; x < short.MaxValue && x > 0; x = (short)(x * 2))
            {
                Row r = s.GetRow(x);

                for (int y = 1; y < 256 && y > 0; y = (short)(y + 2))
                {

                    Cell c = r.GetCell(y);
                    Assert.IsTrue(c.CellFormula != null, "got a formula");

                    Assert.IsTrue(
                    ("" + x + "." + y + operator1 + y + "." + x).Equals(c.CellFormula),
                    "loop Formula is as expected " + x + "." + y + operator1 + y + "." + x + "!=" + c.CellFormula);
                }
            }
        }
        [TestMethod]
        public void TestAreaSum()
        {
            AreaFunctionTest("SUM");
        }
        [TestMethod]
        public void TestAreaAverage()
        {
            AreaFunctionTest("AVERAGE");
        }
        [TestMethod]
        public void TestRefArraySum()
        {
            RefArrayFunctionTest("SUM");
        }
        [TestMethod]
        public void TestAreaArraySum()
        {
            RefAreaArrayFunctionTest("SUM");
        }

        private static void OperationRefTest(String operator1)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;

            //get our minimum values
            r = s.CreateRow(0);
            c = r.CreateCell(1);
            c.CellFormula = ("A2" + operator1 + "A3");

            for (int x = 1; x < short.MaxValue && x > 0; x = (short)(x * 2))
            {
                r = s.CreateRow(x);

                for (int y = 1; y < 256 && y > 0; y++)
                {

                    String ref1 = null;
                    String ref2 = null;
                    short refx1 = 0;
                    short refy1 = 0;
                    short refx2 = 0;
                    short refy2 = 0;
                    if (x + 50 < short.MaxValue)
                    {
                        refx1 = (short)(x + 50);
                        refx2 = (short)(x + 46);
                    }
                    else
                    {
                        refx1 = (short)(x - 4);
                        refx2 = (short)(x - 3);
                    }

                    if (y + 50 < 255)
                    {
                        refy1 = (short)(y + 50);
                        refy2 = (short)(y + 49);
                    }
                    else
                    {
                        refy1 = (short)(y - 4);
                        refy2 = (short)(y - 3);
                    }

                    c = r.GetCell(y);
                    CellReference cr = new CellReference(refx1, refy1, false, false);
                    ref1 = cr.FormatAsString();
                    cr = new CellReference(refx2, refy2, false, false);
                    ref2 = cr.FormatAsString();

                    c = r.CreateCell(y);
                    c.CellFormula = ("" + ref1 + operator1 + ref2);



                }
            }

            //make sure we do the maximum value of the Int operator
            if (s.LastRowNum < short.MaxValue)
            {
                r = s.GetRow(0);
                c = r.CreateCell(0);
                c.CellFormula = ("" + "B1" + operator1 + "IV255");
            }

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            OperationalRefVerify(operator1, wb);
        }

        /**
         * Opens the sheet we wrote out by BinomialOperator and makes sure the formulas
         * all Match what we expect (x operator y)
         */
        private static void OperationalRefVerify(String operator1, HSSFWorkbook wb)
        {

            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Row r = null;
            Cell c = null;

            //get our minimum values
            r = s.GetRow(0);
            c = r.GetCell(1);
            //get our minimum values
            Assert.IsTrue(("A2" + operator1 + "A3").Equals(c.CellFormula), "minval Formula is as expected A2" + operator1 + "A3 != " + c.CellFormula);


            for (int x = 1; x < short.MaxValue && x > 0; x = (short)(x * 2))
            {
                r = s.GetRow(x);

                for (int y = 1; y < 256 && y > 0; y++)
                {

                    int refx1;
                    int refy1;
                    int refx2;
                    int refy2;
                    if (x + 50 < short.MaxValue)
                    {
                        refx1 = x + 50;
                        refx2 = x + 46;
                    }
                    else
                    {
                        refx1 = x - 4;
                        refx2 = x - 3;
                    }

                    if (y + 50 < 255)
                    {
                        refy1 = y + 50;
                        refy2 = y + 49;
                    }
                    else
                    {
                        refy1 = y - 4;
                        refy2 = y - 3;
                    }

                    c = r.GetCell(y);
                    CellReference cr = new CellReference(refx1, refy1, false, false);
                    String ref1 = cr.FormatAsString();
                    ref1 = cr.FormatAsString();
                    cr = new CellReference(refx2, refy2, false, false);
                    String ref2 = cr.FormatAsString();


                    Assert.IsTrue((
                    ("" + ref1 + operator1 + ref2).Equals(c.CellFormula)
                                                             ), "loop Formula is as expected " + ref1 + operator1 + ref2 + "!=" + c.CellFormula
                    );
                }
            }

            //Test our maximum values
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.AreEqual("B1" + operator1 + "IV255", c.CellFormula);
        }



        /**
         * Tests Order wrting out == Order writing in for a given formula
         */
        private static void OrderTest(String formula)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;

            //get our minimum values
            r = s.CreateRow(0);
            c = r.CreateCell(1);
            c.CellFormula = (formula);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);

            //get our minimum values
            r = s.GetRow(0);
            c = r.GetCell(1);
            Assert.IsTrue(formula.Equals(c.CellFormula), "minval Formula is as expected"
                      );
        }

        /**
         * All multi-binomial operator Tests use this to Create a worksheet with a
         * huge set of x operator y formulas.  Next we call BinomialVerify and verify
         * that they are all how we expect.
         */
        private static void BinomialOperator(String operator1)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;

            //get our minimum values
            r = s.CreateRow(0);
            c = r.CreateCell(1);
            c.CellFormula = (1 + operator1 + 1);

            for (int x = 1; x < short.MaxValue && x > 0; x = (short)(x * 2))
            {
                r = s.CreateRow(x);

                for (int y = 1; y < 256 && y > 0; y++)
                {

                    c = r.CreateCell(y);
                    c.CellFormula = ("" + x + operator1 + y);

                }
            }

            //make sure we do the maximum value of the Int operator
            if (s.LastRowNum < short.MaxValue)
            {
                r = s.GetRow(0);
                c = r.CreateCell(0);
                c.CellFormula = ("" + short.MaxValue + operator1 + short.MaxValue);
            }
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            BinomialVerify(operator1, wb);
        }

        /**
         * Opens the sheet we wrote out by BinomialOperator and makes sure the formulas
         * all Match what we expect (x operator y)
         */
        private static void BinomialVerify(String operator1, HSSFWorkbook wb)
        {
            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Row r = null;
            Cell c = null;

            //get our minimum values
            r = s.GetRow(0);
            c = r.GetCell(1);
            Assert.IsTrue(("1" + operator1 + "1").Equals(c.CellFormula),
            "minval Formula is as expected 1" + operator1 + "1 != " + c.CellFormula);

            for (int x = 1; x < short.MaxValue && x > 0; x = (short)(x * 2))
            {
                r = s.GetRow(x);

                for (int y = 1; y < 256 && y > 0; y++)
                {

                    c = r.GetCell(y);

                    Assert.IsTrue(("" + x + operator1 + y).Equals(c.CellFormula),
                        "loop Formula is as expected " + x + operator1 + y + "!=" + c.CellFormula
                    );
                }
            }

            //Test our maximum values
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.IsTrue(
            ("" + short.MaxValue + operator1 + short.MaxValue).Equals(c.CellFormula), "maxval Formula is as expected"

            );
        }

        /**
         * Writes a function then Tests to see if its correct
         */
        public static void AreaFunctionTest(String function)
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;


            r = s.CreateRow(0);

            c = r.CreateCell(0);
            c.CellFormula = (function + "(A2:A3)");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.IsTrue((function + "(A2:A3)").Equals((function + "(A2:A3)")), "function =" + function + "(A2:A3)"
                      );
        }

        /**
         * Writes a function then Tests to see if its correct
         */

        public void RefArrayFunctionTest(String function)
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;


            r = s.CreateRow(0);

            c = r.CreateCell(0);
            c.CellFormula = (function + "(A2,A3)");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.IsTrue((function + "(A2,A3)").Equals(c.CellFormula), "function =" + function + "(A2,A3)"
                      );
        }


        /**
         * Writes a function then Tests to see if its correct
         *
         */
        public void RefAreaArrayFunctionTest(String function)
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;


            r = s.CreateRow(0);

            c = r.CreateCell(0);
            c.CellFormula = (function + "(A2:A4,B2:B4)");
            c = r.CreateCell(1);
            c.CellFormula = (function + "($A$2:$A4,B$2:B4)");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.IsTrue(
                        (function + "(A2:A4,B2:B4)").Equals(c.CellFormula), "function =" + function + "(A2:A4,B2:B4)"
                      );

            c = r.GetCell(1);
            Assert.IsTrue((function + "($A$2:$A4,B$2:B4)").Equals(c.CellFormula),
                "function =" + function + "($A$2:$A4,B$2:B4)"
                      );
        }


        [TestMethod]
        public void TestAbsRefs()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r;
            Cell c;

            r = s.CreateRow(0);
            c = r.CreateCell(0);
            c.CellFormula = ("A3+A2");
            c = r.CreateCell(1);
            c.CellFormula = ("$A3+$A2");
            c = r.CreateCell(2);
            c.CellFormula = ("A$3+A$2");
            c = r.CreateCell(3);
            c.CellFormula = ("$A$3+$A$2");
            c = r.CreateCell(4);
            c.CellFormula = ("SUM($A$3,$A$2)");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);
            Assert.IsTrue(("A3+A2").Equals(c.CellFormula), "A3+A2");
            c = r.GetCell(1);
            Assert.IsTrue(("$A3+$A2").Equals(c.CellFormula), "$A3+$A2");
            c = r.GetCell(2);
            Assert.IsTrue(("A$3+A$2").Equals(c.CellFormula), "A$3+A$2");
            c = r.GetCell(3);
            Assert.IsTrue(("$A$3+$A$2").Equals(c.CellFormula), "$A$3+$A$2");
            c = r.GetCell(4);
            Assert.IsTrue(("SUM($A$3,$A$2)").Equals(c.CellFormula), "SUM($A$3,$A$2)");
        }
        [TestMethod]
        public void TestSheetFunctions()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("A");
            Row r = null;
            Cell c = null;
            r = s.CreateRow(0);
            c = r.CreateCell(0); c.SetCellValue(1);
            c = r.CreateCell(1); c.SetCellValue(2);

            s = wb.CreateSheet("B");
            r = s.CreateRow(0);
            c = r.CreateCell(0); c.CellFormula = ("AVERAGE(A!A1:B1)");
            c = r.CreateCell(1); c.CellFormula = ("A!A1+A!B1");
            c = r.CreateCell(2); c.CellFormula = ("A!$A$1+A!$B1");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            s = wb.GetSheet("B");
            r = s.GetRow(0);
            c = r.GetCell(0);
            Assert.IsTrue(("AVERAGE(A!A1:B1)").Equals(c.CellFormula), "expected: AVERAGE(A!A1:B1) got: " + c.CellFormula);
            c = r.GetCell(1);
            Assert.IsTrue(("A!A1+A!B1").Equals(c.CellFormula), "expected: A!A1+A!B1 got: " + c.CellFormula);
        }
        [TestMethod]
        public void TestRVAoperands()
        {
            string tmpfile = TempFile.GetTempFilePath("TestFormulaRVA", ".xls");

            FileStream out1 = new FileStream(tmpfile, FileMode.Create);
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            Row r = null;
            Cell c = null;


            r = s.CreateRow(0);

            c = r.CreateCell(0);
            c.CellFormula = ("A3+A2");
            c = r.CreateCell(1);
            c.CellFormula = ("AVERAGE(A3,A2)");
            c = r.CreateCell(2);
            c.CellFormula = ("ROW(A3)");
            c = r.CreateCell(3);
            c.CellFormula = ("AVERAGE(A2:A3)");
            c = r.CreateCell(4);
            c.CellFormula = ("POWER(A2,A3)");
            c = r.CreateCell(5);
            c.CellFormula = ("SIN(A2)");

            c = r.CreateCell(6);
            c.CellFormula = ("SUM(A2:A3)");

            c = r.CreateCell(7);
            c.CellFormula = ("SUM(A2,A3)");

            r = s.CreateRow(1); c = r.CreateCell(0); c.SetCellValue(2.0);
            r = s.CreateRow(2); c = r.CreateCell(0); c.SetCellValue(3.0);

            wb.Write(out1);
            out1.Close();
            Assert.IsTrue(File.Exists(tmpfile), "file exists");
        }
        [TestMethod]
        public void TestStringFormulas()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("A");
            Row r = null;
            Cell c = null;
            r = s.CreateRow(0);
            c = r.CreateCell(1); c.CellFormula = ("UPPER(\"abc\")");
            c = r.CreateCell(2); c.CellFormula = ("LOWER(\"ABC\")");
            c = r.CreateCell(3); c.CellFormula = ("CONCATENATE(\" my \",\" name \")");

            HSSFTestDataSamples.WriteOutAndReadBack(wb);

            wb = OpenSample("StringFormulas.xls");
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);
            Assert.AreEqual("UPPER(\"xyz\")", c.CellFormula);
        }
        [TestMethod]
        public void TestLogicalFormulas()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("A");
            Row r = null;
            Cell c = null;
            r = s.CreateRow(0);
            c = r.CreateCell(1); c.CellFormula = ("IF(A1<A2,B1,B2)");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(1);
            Assert.AreEqual("IF(A1<A2,B1,B2)", c.CellFormula,"Formula in cell 1 ");
        }
        [TestMethod]
        public void TestDateFormulas()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("TestSheet1");
            Row r = null;
            Cell c = null;

            r = s.CreateRow(0);
            c = r.CreateCell(0);

            NPOI.SS.UserModel.CellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.DataFormat = (HSSFDataFormat.GetBuiltinFormat("m/d/yy h:mm"));
            c.SetCellValue(new DateTime());
            c.CellStyle = (cellStyle);

            // Assert.AreEqual("Checking hour = " + hour, date.GetTime().GetTime(),
            //              NPOI.SS.UserModel.DateUtil.GetJavaDate(excelDate).GetTime());

            for (int k = 1; k < 100; k++)
            {
                r = s.CreateRow(k);
                c = r.CreateCell(0);
                c.CellFormula = ("A" + (k) + "+1");
                c.CellStyle = cellStyle;
            }

            HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }

        [TestMethod]
        public void TestIfFormulas()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("TestSheet1");
            Row r = null;
            Cell c = null;
            r = s.CreateRow(0);
            c = r.CreateCell(1); c.SetCellValue(1);
            c = r.CreateCell(2); c.SetCellValue(2);
            c = r.CreateCell(3); c.CellFormula = ("MAX(A1:B1)");
            c = r.CreateCell(4); c.CellFormula = ("IF(A1=D1,\"A1\",\"B1\")");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(4);

            Assert.IsTrue(("IF(A1=D1,\"A1\",\"B1\")").Equals(c.CellFormula), "expected: IF(A1=D1,\"A1\",\"B1\") got " + c.CellFormula);

            wb = OpenSample("IfFormulaTest.xls");
            s = wb.GetSheetAt(0);
            r = s.GetRow(3);
            c = r.GetCell(0);
            Assert.IsTrue(("IF(A3=A1,\"A1\",\"A2\")").Equals(c.CellFormula), "expected: IF(A3=A1,\"A1\",\"A2\") got " + c.CellFormula);
            //c = r.GetCell((short)1);
            //Assert.IsTrue("expected: A!A1+A!B1 got: "+c.CellFormula, ("A!A1+A!B1").Equals(c.CellFormula));


            wb = new HSSFWorkbook();
            s = wb.CreateSheet("TestSheet1");
            r = null;
            c = null;
            r = s.CreateRow(0);
            c = r.CreateCell(0); c.CellFormula = ("IF(1=1,0,1)");

            HSSFTestDataSamples.WriteOutAndReadBack(wb);

            wb = new HSSFWorkbook();
            s = wb.CreateSheet("TestSheet1");
            r = null;
            c = null;
            r = s.CreateRow(0);
            c = r.CreateCell(0);
            c.SetCellValue(1);

            c = r.CreateCell(1);
            c.SetCellValue(3);


            Cell formulaCell = r.CreateCell(3);

            r = s.CreateRow(1);
            c = r.CreateCell(0);
            c.SetCellValue(3);

            c = r.CreateCell(1);
            c.SetCellValue(7);

            formulaCell.CellFormula = ("IF(A1=B1,AVERAGE(A1:B1),AVERAGE(A2:B2))");

            HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }
        [TestMethod]
        public void TestSumIf()
        {
            String function = "SUMIF(A1:A5,\">4000\",B1:B5)";

            HSSFWorkbook wb = OpenSample("sumifformula.xls");

            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Row r = s.GetRow(0);
            Cell c = r.GetCell(2);
            Assert.AreEqual(function, c.CellFormula);


            wb = new HSSFWorkbook();
            s = wb.CreateSheet();

            r = s.CreateRow(0);
            c = r.CreateCell(0); c.SetCellValue(1000);
            c = r.CreateCell(1); c.SetCellValue(1);


            r = s.CreateRow(1);
            c = r.CreateCell(0); c.SetCellValue(2000);
            c = r.CreateCell(1); c.SetCellValue(2);

            r = s.CreateRow(2);
            c = r.CreateCell(0); c.SetCellValue(3000);
            c = r.CreateCell(1); c.SetCellValue(3);

            r = s.CreateRow(3);
            c = r.CreateCell(0); c.SetCellValue(4000);
            c = r.CreateCell(1); c.SetCellValue(4);

            r = s.CreateRow(4);
            c = r.CreateCell(0); c.SetCellValue(5000);
            c = r.CreateCell(1); c.SetCellValue(5);

            r = s.GetRow(0);
            c = r.CreateCell(2); c.CellFormula = (function);

            HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }
        [TestMethod]
        public void TestSquareMacro()
        {
            HSSFWorkbook w = OpenSample("SquareMacro.xls");

            NPOI.SS.UserModel.Sheet s0 = w.GetSheetAt(0);
            Row[] r = { s0.GetRow(0), s0.GetRow(1) };

            Cell a1 = r[0].GetCell(0);
            Assert.AreEqual("square(1)", a1.CellFormula);
            Assert.AreEqual(1d, a1.NumericCellValue, 1e-9);

            Cell a2 = r[1].GetCell(0);
            Assert.AreEqual("square(2)", a2.CellFormula);
            Assert.AreEqual(4d, a2.NumericCellValue, 1e-9);

            Cell b1 = r[0].GetCell(1);
            Assert.AreEqual("IF(TRUE,square(1))", b1.CellFormula);
            Assert.AreEqual(1d, b1.NumericCellValue, 1e-9);

            Cell b2 = r[1].GetCell(1);
            Assert.AreEqual("IF(TRUE,square(2))", b2.CellFormula);
            Assert.AreEqual(4d, b2.NumericCellValue, 1e-9);

            Cell c1 = r[0].GetCell(2);
            Assert.AreEqual("square(square(1))", c1.CellFormula);
            Assert.AreEqual(1d, c1.NumericCellValue, 1e-9);

            Cell c2 = r[1].GetCell(2);
            Assert.AreEqual("square(square(2))", c2.CellFormula);
            Assert.AreEqual(16d, c2.NumericCellValue, 1e-9);

            Cell d1 = r[0].GetCell(3);
            Assert.AreEqual("square(one())", d1.CellFormula);
            Assert.AreEqual(1d, d1.NumericCellValue, 1e-9);

            Cell d2 = r[1].GetCell(3);
            Assert.AreEqual("square(two())", d2.CellFormula);
            Assert.AreEqual(4d, d2.NumericCellValue, 1e-9);
        }
        [TestMethod]
        public void TestStringFormulaRead()
        {
            HSSFWorkbook w = OpenSample("StringFormulas.xls");
            Cell c = w.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.AreEqual("XYZ", c.RichStringCellValue.String, "String Cell value");
        }

        /** Test for bug 34021*/
        [TestMethod]
        public void TestComplexSheetRefs()
        {
            HSSFWorkbook sb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s1 = sb.CreateSheet("Sheet a.1");
            NPOI.SS.UserModel.Sheet s2 = sb.CreateSheet("Sheet.A");
            s2.CreateRow(1).CreateCell(2).CellFormula = ("'Sheet a.1'!A1");
            s1.CreateRow(1).CreateCell(2).CellFormula = ("'Sheet.A'!A1");
            string tmpfile = TempFile.GetTempFilePath("TestComplexSheetRefs", ".xls");
            FileStream fs = new FileStream(tmpfile,FileMode.Create);
            sb.Write(fs);
            fs.Close();
        }

        /** Unknown Ptg 3C*/
        [TestMethod]
        public void Test27272_1()  {
            HSSFWorkbook wb = OpenSample("27272_1.xls");
            wb.GetSheetAt(0);
            Assert.AreEqual("Compliance!#REF!", wb.GetNameAt(0).RefersToFormula, "Reference for named range ");
            string tmpfile=TempFile.GetTempFilePath("bug27272_1",".xls");
            FileStream fs = new FileStream(tmpfile, FileMode.OpenOrCreate);
            wb.Write(fs);
            fs.Close();
            Console.WriteLine("Open "+Path.GetFullPath(tmpfile)+" in Excel");
        
        }
        /** Unknown Ptg 3D*/
        [TestMethod]
        public void Test27272_2()  {
        HSSFWorkbook wb = OpenSample("27272_2.xls");
        Assert.AreEqual("LOAD.POD_HISTORIES!#REF!", wb.GetNameAt(0).RefersToFormula, "Reference for named range ");
        string tmpfile = TempFile.GetTempFilePath("bug27272_2", ".xls");
        FileStream fs = new FileStream(tmpfile, FileMode.OpenOrCreate);
        wb.Write(fs);
        Console.WriteLine("Open " + Path.GetFullPath(tmpfile) + " in Excel");
    }

        /** MissingArgPtg */
        [TestMethod]
        public void TestMissingArgPtg()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Cell cell = wb.CreateSheet("Sheet1").CreateRow(4).CreateCell(0);
            cell.CellFormula = ("IF(A1=\"A\",1,)");
        }
        [TestMethod]
        public void TestSharedFormula()
        {
            HSSFWorkbook wb = OpenSample("SharedFormulaTest.xls");

            Assert.AreEqual("A$1*2", wb.GetSheetAt(0).GetRow(1).GetCell(1).ToString());
            Assert.AreEqual("$A11*2", wb.GetSheetAt(0).GetRow(11).GetCell(1).ToString());
            Assert.AreEqual("DZ2*2", wb.GetSheetAt(0).GetRow(1).GetCell(128).ToString());
            Assert.AreEqual("B32770*2", wb.GetSheetAt(0).GetRow(32768).GetCell(1).ToString());
        }
    }

}