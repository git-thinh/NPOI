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

/*
 * TestFormulaRecordAggregate.java
 *
 * Created on March 21, 2003, 12:32 AM
 */

namespace TestCases.HSSF.Record.Aggregates
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     *
     * @author  avik
     */
    [TestClass]
    public class TestFormulaRecordAggregate
    {
        [TestMethod]
        public void TestBasic()
        {
            FormulaRecord f = new FormulaRecord();
            f.SetCachedResultTypeString();
            StringRecord s = new StringRecord();
            s.String = ("abc");
            FormulaRecordAggregate fagg = new FormulaRecordAggregate(f, s, SharedValueManager.EMPTY);
            Assert.AreEqual("abc", fagg.StringValue);
        }
        /**
 * Sometimes a {@link StringRecord} appears after a {@link FormulaRecord} even though the
 * formula has evaluated to a text value.  This might be more likely to occur when the formula
 * <i>can</i> evaluate to a text value.<br/>
 * Bug 46213 attachment 22874 has such an extra {@link StringRecord} at stream offset 0x5765.
 * This file seems to open in Excel (2007) with no trouble.  When it is re-saved, Excel omits
 * the extra record.  POI should do the same.
 */
        [TestMethod]
        public void TestExtraStringRecord_bug46213()
        {
            FormulaRecord fr = new FormulaRecord();
            fr.Value = (2.0);
            StringRecord sr = new StringRecord();
            sr.String = ("NA");
            SharedValueManager svm = SharedValueManager.EMPTY;
            FormulaRecordAggregate fra;

            try
            {
                fra = new FormulaRecordAggregate(fr, sr, svm);
            }
            catch (RecordFormatException e)
            {
                if ("String record was  supplied but formula record flag is not  set".Equals(e.Message))
                {
                    throw new AssertFailedException("Identified bug 46213");
                }
                throw e;
            }
            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rc = new TestCases.HSSF.UserModel.RecordInspector.RecordCollector();
            fra.VisitContainedRecords(rc);
            Record[] vraRecs = rc.Records;
            Assert.AreEqual(1, vraRecs.Length);
            Assert.AreEqual(fr, vraRecs[0]);
        }
    }
}