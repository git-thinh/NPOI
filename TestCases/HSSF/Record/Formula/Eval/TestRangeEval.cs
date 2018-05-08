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

namespace TestCases.HSSF.Record.Formula.Eval
{

    using NPOI.HSSF.Record.Formula.Eval;
    using TestCases.HSSF.Record.Formula.Functions;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using HSSFFunctions = NPOI.HSSF.Record.Formula.Functions;
    using System;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Test for unary plus operator Evaluator.
     *
     * @author Josh Micich
     */
    [TestClass]
    public class TestRangeEval
    {
        [TestMethod]
        public void TestPermutations()
        {

            Confirm("B3", "D7", "B3:D7");
            Confirm("B1", "B1", "B1:B1");

            Confirm("B7", "D3", "B3:D7");
            Confirm("D3", "B7", "B3:D7");
            Confirm("D7", "B3", "B3:D7");
        }

        private static void Confirm(String refA, String refB, String expectedAreaRef)
        {

            ValueEval[] args = {
			CreateRefEval(refA),
			CreateRefEval(refB),
		};
            AreaReference ar = new AreaReference(expectedAreaRef);
            ValueEval result = EvalInstances.Range.Evaluate(args, 0, (short)0);
            Assert.IsTrue(result is AreaEval);
            AreaEval ae = (AreaEval)result;
            Assert.AreEqual(ar.FirstCell.Row, ae.FirstRow);
            Assert.AreEqual(ar.LastCell.Row, ae.LastRow);
            Assert.AreEqual(ar.FirstCell.Col, ae.FirstColumn);
            Assert.AreEqual(ar.LastCell.Col, ae.LastColumn);
        }

        private static ValueEval CreateRefEval(String refStr)
        {
            CellReference cr = new CellReference(refStr);
            return new MockRefEval(cr.Row, cr.Col);

        }

        private class MockRefEval : RefEvalBase
        {

            public MockRefEval(int rowIndex, int columnIndex)
                : base(rowIndex, columnIndex)
            {

            }
            public override ValueEval InnerValueEval
            {
                get
                {
                    throw new Exception("not expected to be called during this test");
                }
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx,
                    int relLastColIx)
            {
                AreaI area = new OffsetArea(Row, Column,
                        relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);
                return new MockAreaEval(area);
            }
        }

        private class MockAreaEval : AreaEvalBase
        {

            public MockAreaEval(AreaI ptg)
                : base(ptg)
            {

            }
            public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
            {
                throw new Exception("not expected to be called during this test");
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx,
                    int relLastColIx)
            {
                AreaI area = new OffsetArea(FirstRow, FirstColumn,
                        relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);

                return new MockAreaEval(area);
            }
        }
        [TestMethod]
        public void TestRangeUsingOffSetFunc_bug46948()
        {
            Workbook wb = new HSSFWorkbook();
            Row row = wb.CreateSheet("Sheet1").CreateRow(0);
            Cell cellA1 = row.CreateCell(0);
            Cell cellB1 = row.CreateCell(1);
            row.CreateCell(2).SetCellValue(5.0); // C1
            row.CreateCell(3).SetCellValue(7.0); // D1
            row.CreateCell(4).SetCellValue(9.0); // E1


            try
            {
                cellA1.CellFormula = ("SUM(C1:OFFSET(C1,0,B1))");
            }
            catch (Exception e)
            {
                // TODO fix formula Parser to handle ':' as a proper operator
                if (!e.GetType().Name.StartsWith(typeof(FormulaParser).Name))
                {
                    throw e;
                }
                // FormulaParseException is expected until the Parser is fixed up
                // Poke the formula in directly:
                //pokeInOffSetFormula(cellA1);
            }


            cellB1.SetCellValue(1.0); // range will be C1:D1

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv;
            try
            {
                cv = fe.Evaluate(cellA1);
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Unexpected ref arg class (NPOI.SS.formula.LazyAreaEval)"))
                {
                    throw new AssertFailedException("Identified bug 46948");
                }
                throw e;
            }

            Assert.AreEqual(12.0, cv.NumberValue, 0.0);

            cellB1.SetCellValue(2.0); // range will be C1:E1
            fe.NotifyUpdateCell(cellB1);
            cv = fe.Evaluate(cellA1);
            Assert.AreEqual(21.0, cv.NumberValue, 0.0);

            cellB1.SetCellValue(0.0); // range will be C1:C1
            fe.NotifyUpdateCell(cellB1);
            cv = fe.Evaluate(cellA1);
            Assert.AreEqual(5.0, cv.NumberValue, 0.0);
        }

        /**
         * Directly Sets the formula "SUM(C1:OFFSET(C1,0,B1))" in the specified cell.
         * This hack can be Removed when the formula Parser can handle functions as
         * operands to the range (:) operator.
         *
         */
        //private static void pokeInOffSetFormula(HSSFCell cell) {
        //    cell.SetCellFormula("1");
        //    FormulaRecordAggregate fr;
        //    try {
        //        Field field = HSSFCell.class.GetDeclaredField("_record");
        //        field.SetAccessible(true);
        //        fr = (FormulaRecordAggregate) field.Get(cell);
        //    } catch (ArgumentException e) {
        //        throw new Exception(e);
        //    } catch (IllegalAccessException e) {
        //        throw new Exception(e);
        //    } catch (NoSuchFieldException e) {
        //        throw new Exception(e);
        //    }
        //    Ptg[] ptgs = {
        //            new RefPtg("C1"),
        //            new RefPtg("C1"),
        //            new IntPtg(0),
        //            new RefPtg("B1"),
        //            FuncVarPtg.Create("OFFSET", (byte)3),
        //            RangePtg.instance,
        //            AttrPtg.SUM,
        //        };
        //    fr.SetParsedExpression(ptgs);
        //}
    }
}










