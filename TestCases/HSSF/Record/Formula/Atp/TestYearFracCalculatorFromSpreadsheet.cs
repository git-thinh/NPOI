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

namespace TestCases.HSSF.Record.Formula.ATP
{
    using System;
    using System.IO;
    using System.Collections;
    using TestCases.HSSF;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record.Formula.Atp;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.UserModel;
    /**
     * Tests YearFracCalculator using test-cases listed in a sample spreadsheet
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestYearFracCalculatorFromSpreadsheet
    {

        private class SS
        {

            public static int BASIS_COLUMN = 1; // "B"
            public static int START_YEAR_COLUMN = 2; // "C"
            public static int END_YEAR_COLUMN = 5; // "F"
            public static int YEARFRAC_FORMULA_COLUMN = 11; // "L"
            public static int EXPECTED_RESULT_COLUMN = 13; // "N"
        }
        [TestMethod]
        public void TestAll()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("yearfracExamples.xls");
            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);
            HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(wb);
            int nSuccess = 0;
            int nFailures = 0;
            int nUnexpectedErrors = 0;
            IEnumerator rowIterator = sheet.GetRowEnumerator();
            while (rowIterator.MoveNext())
            {
                Row row = (Row)rowIterator.Current;

                Cell cell = row.GetCell(SS.YEARFRAC_FORMULA_COLUMN);
                if (cell == null || cell.CellType != NPOI.SS.UserModel.CellType.FORMULA)
                {
                    continue;
                }
                try
                {
                    ProcessRow(row, cell, formulaEvaluator);
                    nSuccess++;
                }
                catch (AssertFailedException e)
                {
                    nFailures++;
                }
                catch (Exception e)
                {
                    nUnexpectedErrors++;
                }
            }
            if (nUnexpectedErrors + nFailures > 0)
            {
                String msg = nFailures + " failures(s) and " + nUnexpectedErrors
                    + " unexpected errors(s) occurred. See stderr for details";
                throw new AssertFailedException(msg);
            }
            if (nSuccess < 1)
            {
                throw new Exception("No test sample cases found");
            }
        }

        private static void ProcessRow(Row row, Cell cell, HSSFFormulaEvaluator formulaEvaluator)
        {

            double startDate = MakeDate(row, SS.START_YEAR_COLUMN);
            double endDate = MakeDate(row, SS.END_YEAR_COLUMN);

            int basis = GetIntCell(row, SS.BASIS_COLUMN);

            double expectedValue = GetDoubleCell(row, SS.EXPECTED_RESULT_COLUMN);

            double actualValue;
            try
            {
                actualValue = YearFracCalculator.Calculate(startDate, endDate, basis);
            }
            catch (EvaluationException)
            {
                throw;
            }
            if (expectedValue != actualValue)
            {
                throw new AssertFailedException("Direct calculate failed - row " + (row.RowNum + 1) +
                        ", expected:" + expectedValue.ToString() + ", actual" + actualValue.ToString());
            }
            actualValue = formulaEvaluator.Evaluate(cell).NumberValue;
            if (expectedValue != actualValue)
            {
                throw new AssertFailedException("Formula evaluate failed - row " + (row.RowNum + 1) +
                    ", expected:" + expectedValue.ToString()+", actual"+actualValue.ToString());
            }
        }

        private static double MakeDate(Row row, int yearColumn)
        {
            int year = GetIntCell(row, yearColumn + 0);
            int month = GetIntCell(row, yearColumn + 1);
            int day = GetIntCell(row, yearColumn + 2);

            DateTime dt = new DateTime(year, month, day, 0, 0, 0);
            return NPOI.SS.UserModel.DateUtil.GetExcelDate(dt);
        }

        private static int GetIntCell(Row row, int colIx)
        {
            double dVal = GetDoubleCell(row, colIx);
            if (Math.Floor(dVal) != dVal)
            {
                throw new Exception("Non integer value (" + dVal
                        + ") cell found at column " + (char)('A' + colIx));
            }
            return (int)dVal;
        }

        private static double GetDoubleCell(Row row, int colIx)
        {
            Cell cell = row.GetCell(colIx);
            if (cell == null)
            {
                throw new Exception("No cell found at column " + colIx);
            }
            double dVal = cell.NumericCellValue;
            return dVal;
        }
    }
}