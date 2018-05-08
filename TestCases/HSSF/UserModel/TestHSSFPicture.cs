/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.SS.UserModel;
    using TestCases.HSSF;

    /**
     * Test <c>HSSFPicture</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestClass]
    public class TestHSSFPicture
    {
        [TestMethod]
        public void TestResize()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sh1 = wb.CreateSheet();
            Drawing p1 = sh1.CreateDrawingPatriarch();

            byte[] pictureData = HSSFTestDataSamples.GetTestDataFileContent("logoKarmokar4.png");
            int idx1 = wb.AddPicture(pictureData, PictureType.PNG);
            Picture picture1 = p1.CreatePicture(new HSSFClientAnchor(), idx1);
            ClientAnchor anchor1 = picture1.GetPreferredSize();

            //assert against what would BiffViewer print if we insert the image in xls and dump the file
            Assert.AreEqual(0, anchor1.Col1);
            Assert.AreEqual(0, anchor1.Row1);
            Assert.AreEqual(1, anchor1.Col2);
            Assert.AreEqual(9, anchor1.Row2);
            Assert.AreEqual(0, anchor1.Dx1);
            Assert.AreEqual(0, anchor1.Dy1);
            Assert.AreEqual(848, anchor1.Dx2);
            Assert.AreEqual(240, anchor1.Dy2);
        }

        /**
         * Bug # 45829 reported ArithmeticException (/ by zero) when resizing png with zero DPI.
         */
        [TestMethod]
        public void Test45829()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sh1 = wb.CreateSheet();
            Drawing p1 = sh1.CreateDrawingPatriarch();

            byte[] pictureData = HSSFTestDataSamples.GetTestDataFileContent("45829.png");
            int idx1 = wb.AddPicture(pictureData, PictureType.PNG);
            Picture pic = p1.CreatePicture(new HSSFClientAnchor(), idx1);
            pic.Resize();
        }
    }
}