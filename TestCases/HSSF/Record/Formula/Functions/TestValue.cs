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

namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.HSSF.Record.Formula.Functions;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * Tests for {@link Value}
     *
     * @author Josh Micich
     */
    [TestClass]
    public class TestValue
    {

        private static ValueEval InvokeValue(String strText)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(strText), };
            return new Value().Evaluate(args, -1, (short)-1);
        }

        private static void ConfirmValue(String strText, double expected)
        {
            ValueEval result = InvokeValue(strText);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.0);
        }

        private static void ConfirmValueError(String strText)
        {
            ValueEval result = InvokeValue(strText);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }
        [TestMethod]
        public void TestBasic()
        {

            ConfirmValue("100", 100);
            ConfirmValue("-2.3", -2.3);
            ConfirmValue(".5", 0.5);
            ConfirmValue(".5e2", 50);
            ConfirmValue(".5e-2", 0.005);
            ConfirmValue(".5e+2", 50);
            ConfirmValue("+5", 5);
            ConfirmValue("$1,000", 1000);
            ConfirmValue("100.5e1", 1005);
            ConfirmValue("1,0000", 10000);
            ConfirmValue("1,000,0000", 10000000);
            ConfirmValue("1,000,0000,00000", 1000000000000.0);
            ConfirmValue(" 100 ", 100);
            ConfirmValue(" + 100", 100);
            ConfirmValue("10000", 10000);
            ConfirmValue("$-5", -5);
            ConfirmValue("$.5", 0.5);
            ConfirmValue("123e+5", 12300000);
            ConfirmValue("1,000e2", 100000);
            ConfirmValue("$10e2", 1000);
            ConfirmValue("$1,000e2", 100000);
        }

        [TestMethod]
        public void TestErrors()
        {
            ConfirmValueError("1+1");
            ConfirmValueError("1 1");
            ConfirmValueError("1,00.0");
            ConfirmValueError("1,00");
            ConfirmValueError("$1,00.5e1");
            ConfirmValueError("1,00.5e1");
            ConfirmValueError("1,0,000");
            ConfirmValueError("1,00,000");
            ConfirmValueError("++100");
            ConfirmValueError("$$5");
            ConfirmValueError("-");
            ConfirmValueError("+");
            ConfirmValueError("$");
            ConfirmValueError(",300");
            ConfirmValueError("0.233,4");
            ConfirmValueError("1e2.5");
        }
    }
}