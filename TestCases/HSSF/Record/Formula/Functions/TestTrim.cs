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
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record.Formula.Functions;
    using NPOI.HSSF.Record.Formula.Eval;
    /**
     * Tests for Excel function TRIM()
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestTrim
    {


        private static ValueEval InvokeTrim(ValueEval text)
        {
            ValueEval[] args = new ValueEval[] { text, };
            return new Trim().Evaluate(args, -1, (short)-1);
        }

        private void ConfirmTrim(ValueEval text, String expected)
        {
            ValueEval result = InvokeTrim(text);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue);
        }

        private void ConfirmTrim(ValueEval text, ErrorEval expectedError)
        {
            ValueEval result = InvokeTrim(text);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(expectedError.ErrorCode, ((ErrorEval)result).ErrorCode);
        }
        [TestMethod]
        public void TestBasic()
        {

            ConfirmTrim(new StringEval(" hi "), "hi");
            ConfirmTrim(new StringEval("hi "), "hi");
            ConfirmTrim(new StringEval("  hi"), "hi");
            ConfirmTrim(new StringEval(" hi there  "), "hi there");
            ConfirmTrim(new StringEval(""), "");
            ConfirmTrim(new StringEval("   "), "");
        }

        /**
         * Valid cases where text arg is1 not exactly a string
         */
        [TestMethod]
        public void TestUnusualArgs()
        {

            // text (first) arg type is1 number, other args are strings with fractional digits 
            ConfirmTrim(new NumberEval(123456), "123456");
            ConfirmTrim(BoolEval.FALSE, "FALSE");
            ConfirmTrim(BoolEval.TRUE, "TRUE");
            ConfirmTrim(BlankEval.instance, "");
        }
        [TestMethod]
        public void TestErrors()
        {
            ConfirmTrim(ErrorEval.NAME_INVALID, ErrorEval.NAME_INVALID);
        }
    }

}