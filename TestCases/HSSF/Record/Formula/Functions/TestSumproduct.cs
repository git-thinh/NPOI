/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record.Formula.Functions
{
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Formula.Functions;
    using NPOI.HSSF.Record.Formula.Eval;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Test cases for SUMPRODUCT()
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestSumproduct
    {

        private static ValueEval InvokeSumproduct(ValueEval[] args)
        {
            // srcCellRow and srcCellColumn are ignored by SUMPRODUCT
            return new Sumproduct().Evaluate(args, -1, (short)-1);
        }
        private static void ConfirmDouble(double expected, ValueEval actualEval)
        {
            if (!(actualEval is NumericValueEval))
            {
                Assert.Fail("Expected numeric result");
            }
            NumericValueEval nve = (NumericValueEval)actualEval;
            Assert.AreEqual(expected, nve.NumberValue, 0);
        }
        [TestMethod]
        public void TestScalarSimple()
        {

            RefEval refEval = EvalFactory.CreateRefEval("A1", new NumberEval(3));
            ValueEval[] args = {
			    refEval, 
			    new NumberEval(2),
		    };
            ValueEval result = InvokeSumproduct(args);
            ConfirmDouble(6D, result);
        }

        [TestMethod]
        public void TestAreaSimple()
        {

            ValueEval[] aValues = {
			    new NumberEval(2),
			    new NumberEval(4),
			    new NumberEval(5),
		    };
            ValueEval[] bValues = {
			    new NumberEval(3),
			    new NumberEval(6),
			    new NumberEval(7),
		    };

            AreaEval aeA = EvalFactory.CreateAreaEval("A1:A3", aValues);
            AreaEval aeB = EvalFactory.CreateAreaEval("B1:B3", bValues);


            ValueEval[] args = { aeA, aeB, };
            ValueEval result = InvokeSumproduct(args);
            ConfirmDouble(65D, result);
        }

        /**
         * For scalar products, the terms may be 1x1 area refs
         */
        [TestMethod]
        public void TestOneByOneArea()
        {

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A1", new ValueEval[] { new NumberEval(7), });

            ValueEval[] args = {
				ae, 
				new NumberEval(2),
			};
            ValueEval result = InvokeSumproduct(args);
            ConfirmDouble(14D, result);
        }

        [TestMethod]
        public void TestMismatchAreaDimensions()
        {

            AreaEval aeA = EvalFactory.CreateAreaEval("A1:A3", new ValueEval[3]);
            AreaEval aeB = EvalFactory.CreateAreaEval("B1:D1", new ValueEval[3]);

            ValueEval[] args;
            args = new ValueEval[] { aeA, aeB, };
            Assert.AreEqual(ErrorEval.VALUE_INVALID, InvokeSumproduct(args));

            args = new ValueEval[] { aeA, new NumberEval(5), };
            Assert.AreEqual(ErrorEval.VALUE_INVALID, InvokeSumproduct(args));
        }
        [TestMethod]
        public void TestAreaWithErrorCell()
        {
            ValueEval[] aValues = {
			    ErrorEval.REF_INVALID,
			    null,
		    };
            AreaEval aeA = EvalFactory.CreateAreaEval("A1:A2", aValues);
            AreaEval aeB = EvalFactory.CreateAreaEval("B1:B2", new ValueEval[2]);

            ValueEval[] args = { aeA, aeB, };
            Assert.AreEqual(ErrorEval.REF_INVALID, InvokeSumproduct(args));
        }

    }
}