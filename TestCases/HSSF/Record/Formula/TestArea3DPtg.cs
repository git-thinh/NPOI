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


namespace TestCases.HSSF.Record.Formula
{
    using System;
    using NPOI.HSSF.Record.Formula;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.UserModel;

    /**
     * Tests for Area3DPtg
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestArea3DPtg : AbstractPtgTestCase
    {

        /**
         * Confirms that sheet names Get properly escaped
         */
        [TestMethod]
        public void TestToFormulaString()
        {

            Area3DPtg tarGet = new Area3DPtg("A1:B1", (short)0);

            String sheetName = "my sheet";

            HSSFWorkbook wb = CreateWorkbookWithSheet(sheetName);
            HSSFEvaluationWorkbook book = HSSFEvaluationWorkbook.Create(wb);
            Assert.AreEqual("'my sheet'!A1:B1", tarGet.ToFormulaString(book));

            wb.SetSheetName(0, "Sheet1");
            Assert.AreEqual("Sheet1!A1:B1", tarGet.ToFormulaString(book));

            wb.SetSheetName(0, "C64");
            Assert.AreEqual("'C64'!A1:B1", tarGet.ToFormulaString(book));
        }
    }
}
