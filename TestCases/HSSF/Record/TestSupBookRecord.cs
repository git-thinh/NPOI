
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



namespace TestCases.HSSF.Record
{

    using System;
    using NPOI.HSSF.Record;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Tests the serialization and deserialization of the SupBook record
     * class works correctly.  
     *
     * @author Andrew C. Oliver (acoliver at apache dot org)
     */
    [TestClass]
    public class TestSupBookRecord
    {
        /**
         * This contains a fake data section of a SubBookRecord
         */
        byte[] dataIR = new byte[] {
        (byte)0x04,(byte)0x00,(byte)0x01,(byte)0x04,
    };
        byte[] dataAIF = new byte[] {
        (byte)0x01,(byte)0x00,(byte)0x01,(byte)0x3A,
    };
        byte[] dataER = new byte[] {
        (byte)0x02,(byte)0x00,
        (byte)0x07,(byte)0x00,   (byte)0x00,   
                (byte)'t', (byte)'e', (byte)'s', (byte)'t', (byte)'U', (byte)'R', (byte)'L',  
        (byte)0x06,(byte)0x00,   (byte)0x00,   
                (byte)'S', (byte)'h', (byte)'e', (byte)'e', (byte)'t', (byte)'1', 
        (byte)0x06,(byte)0x00,   (byte)0x00,   
                (byte)'S', (byte)'h', (byte)'e', (byte)'e', (byte)'t', (byte)'2', 
   };

        public TestSupBookRecord()
        {

        }

        /**
         * tests that we can load the record
         */
        [TestMethod]
        public void TestLoadIR()
        {

            SupBookRecord record = new SupBookRecord(TestcaseRecordInputStream.Create(0x01AE, dataIR));
            Assert.IsTrue(record.IsInternalReferences);             //expected flag
            Assert.AreEqual(0x4, record.NumberOfSheets);    //expected # of sheets

            Assert.AreEqual(8, record.RecordSize);  //sid+size+data
        }
        /**
         * tests that we can load the record
         */
        [TestMethod]
        public void TestLoadER()
        {

            SupBookRecord record = new SupBookRecord(TestcaseRecordInputStream.Create(0x01AE, dataER));
            Assert.IsTrue(record.IsExternalReferences);             //expected flag
            Assert.AreEqual(0x2, record.NumberOfSheets);    //expected # of sheets

            Assert.AreEqual(34, record.RecordSize);  //sid+size+data

            Assert.AreEqual("testURL", record.URL);
            String[] sheetNames = record.SheetNames;
            Assert.AreEqual(2, sheetNames.Length);
            Assert.AreEqual("Sheet1", sheetNames[0]);
            Assert.AreEqual("Sheet2", sheetNames[1]);
        }

        /**
         * tests that we can load the record
         */
        [TestMethod]
        public void TestLoadAIF()
        {

            SupBookRecord record = new SupBookRecord(TestcaseRecordInputStream.Create(0x01AE, dataAIF));
            Assert.IsTrue(record.IsAddInFunctions);             //expected flag
            Assert.AreEqual(0x1, record.NumberOfSheets);    //expected # of sheets
            Assert.AreEqual(8, record.RecordSize);  //sid+size+data
        }

        /**
         * Tests that we can store the record
         *
         */
        [TestMethod]
        public void TestStoreIR()
        {
            SupBookRecord record = SupBookRecord.CreateInternalReferences((short)4);

            TestcaseRecordInputStream.ConfirmRecordEncoding(0x01AE, dataIR, record.Serialize());
        }
        [TestMethod]
        public void TestStoreER()
        {
            String url = "testURL";
            String[] sheetNames = { "Sheet1", "Sheet2", };
            SupBookRecord record = SupBookRecord.CreateExternalReferences(url, sheetNames);

            TestcaseRecordInputStream.ConfirmRecordEncoding(0x01AE, dataER, record.Serialize());
        }
    }
}