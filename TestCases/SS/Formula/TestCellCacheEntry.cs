/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula
{

    using System;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.SS.Formula;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Tests {@link CellCacheEntry}.
     *
     * @author Josh Micich
     */
    [TestClass]
    public class TestCellCacheEntry
    {
        [TestMethod]
        public void TestBasic()
        {
            CellCacheEntry pcce = new PlainValueCellCacheEntry(new NumberEval(42.0));
            ValueEval ve = pcce.GetValue();
            Assert.AreEqual(42, ((NumberEval)ve).NumberValue, 0.0);

            FormulaCellCacheEntry fcce = new FormulaCellCacheEntry();
            fcce.UpdateFormulaResult(new NumberEval(10.0), CellCacheEntry.EMPTY_ARRAY, null);

            ve = fcce.GetValue();
            Assert.AreEqual(10, ((NumberEval)ve).NumberValue, 0.0);
        }
    }
}