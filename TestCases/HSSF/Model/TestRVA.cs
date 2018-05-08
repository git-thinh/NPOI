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

namespace TestCases.HSSF.Model
{
    using System;
    using System.Text;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using TestCases.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Tests 'operand class' transformation performed by
     * <tt>OperandClassTransformer</tt> by comparing its results with those
     * directly produced by Excel (in a sample spReadsheet).
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestRVA
    {

        [TestMethod]
        public void TestFormulas()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("testRVA.xls");
            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);

            int countFailures = 0;
            int countErrors = 0;

            int rowIx = 0;
            while (rowIx < 65535)
            {
                Row row = sheet.GetRow(rowIx);
                if (row == null)
                {
                    break;
                }
                Cell cell = row.GetCell(0);
                if (cell == null || cell.CellType == NPOI.SS.UserModel.CellType.BLANK)
                {
                    break;
                }
                String formula = cell.CellFormula;
                //try
                //{
                    ConfirmCell(cell, formula, wb);
                //}
                //catch (AssertFailedException e)
                //{
                //    Console.Error.WriteLine("Problem with row[" + rowIx + "] formula '" + formula + "'");
                //    Console.Error.WriteLine(e.Message);
                //    countFailures++;
                //}
                //catch (Exception e)
                //{
                //    Console.Error.WriteLine("Problem with row[" + rowIx + "] formula '" + formula + "'");
                //    countErrors++;
                //}
                rowIx++;
            }
            if (countErrors + countFailures > 0)
            {
                String msg = "One or more RVA tests failed: countFailures=" + countFailures
                        + " countFailures=" + countErrors + ". See stderr for details.";
                throw new AssertFailedException(msg);
            }
        }

        private void ConfirmCell(Cell formulaCell, String formula, HSSFWorkbook wb)
        {
            Ptg[] excelPtgs = FormulaExtractor.GetPtgs(formulaCell);
            Ptg[] poiPtgs = HSSFFormulaParser.Parse(formula, wb);
            int nExcelTokens = excelPtgs.Length;
            int nPoiTokens = poiPtgs.Length;
            if (nExcelTokens != nPoiTokens)
            {
                if (nExcelTokens == nPoiTokens + 1 && excelPtgs[0].GetType() == typeof(AttrPtg))
                {
                    // compensate for missing tAttrVolatile, which belongs in any formula 
                    // involving OFFSET() et al. POI currently does not insert where required
                    Ptg[] temp = new Ptg[nExcelTokens];
                    temp[0] = excelPtgs[0];
                    Array.Copy(poiPtgs, 0, temp, 1, nPoiTokens);
                    poiPtgs = temp;
                }
                else
                {
                    throw new Exception("Expected " + nExcelTokens + " tokens but got "
                            + nPoiTokens);
                }
            }
            bool hasMismatch = false;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < nExcelTokens; i++)
            {
                Ptg poiPtg = poiPtgs[i];
                Ptg excelPtg = excelPtgs[i];
                if (excelPtg.GetType() != poiPtg.GetType())
                {
                    hasMismatch = true;
                    sb.Append("  mismatch token type[" + i + "] " + GetShortClassName(excelPtg) + " "
                            + GetOperandClassName(excelPtg) + " - " + GetShortClassName(poiPtg) + " "
                            + GetOperandClassName(poiPtg));
                    sb.Append(Environment.NewLine);
                    continue;
                }
                if (poiPtg.IsBaseToken)
                {
                    continue;
                }
                sb.Append("  token[" + i + "] " + excelPtg.ToString() + " "
                        + GetOperandClassName(excelPtg));

                if (excelPtg.PtgClass != poiPtg.PtgClass)
                {
                    hasMismatch = true;
                    sb.Append(" - was " + GetOperandClassName(poiPtg));
                }
                sb.Append(Environment.NewLine);
            }
            //if (false)
            //{ // Set 'true' to see trace of RVA values
            //    Console.WriteLine(formula);
            //    Console.WriteLine(sb.ToString());
            //}
            if (hasMismatch)
            {
                throw new AssertFailedException(sb.ToString());
            }
        }

        private String GetShortClassName(Object o)
        {
            String cn = o.GetType().Name;
            int pos = cn.LastIndexOf('.');
            return cn.Substring(pos + 1);
        }

        private static String GetOperandClassName(Ptg ptg)
        {
            byte ptgClass = ptg.PtgClass;
            switch (ptgClass)
            {
                case Ptg.CLASS_REF: return "R";
                case Ptg.CLASS_VALUE: return "V";
                case Ptg.CLASS_ARRAY: return "A";
            }
            throw new Exception("Unknown operand class (" + ptgClass + ")");
        }
    }
}