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

namespace TestCases.HSSF.Record.Formula.Eval
{
    using System;
    using System.Collections;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Formula.Eval;
    using TestCases.HSSF.Record.Formula.Functions;
    /**
     * Tests for <tt>AreaEval</tt>
     *  
     * @author Josh Micich
     */
    [TestClass]
    public class TestAreaEval
    {
        [TestMethod]
        public void TestGetValue_bug44950()
        {

            AreaPtg ptg = new AreaPtg("B2:D3");
            NumberEval one = new NumberEval(1);
            ValueEval[] values = {
				one,	
				new NumberEval(2),	
				new NumberEval(3),	
				new NumberEval(4),	
				new NumberEval(5),	
				new NumberEval(6),	
		};
            AreaEval ae = EvalFactory.CreateAreaEval(ptg, values);
            if (one == ae.GetValueAt(1, 2))
            {
                throw new AssertFailedException("Identified bug 44950 a");
            }
            Confirm(1, ae, 1, 1);
            Confirm(2, ae, 1, 2);
            Confirm(3, ae, 1, 3);
            Confirm(4, ae, 2, 1);
            Confirm(5, ae, 2, 2);
            Confirm(6, ae, 2, 3);

        }

        private static void Confirm(int expectedValue, AreaEval ae, int row, int col)
        {
            NumberEval v = (NumberEval)ae.GetValueAt(row, col);
            Assert.AreEqual(expectedValue, v.NumberValue, 0.0);
        }

    }
}