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

namespace NPOI.HSSF.Record.Formula.Functions
{
    using System;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Tests for the INDIRECT() function.</p>
     *
     * @author Josh Micich
     */
    [TestClass]
    public class TestIndirect
    {
        private void CreateDataRow(Sheet sheet, int rowIndex, params double[] vals)
        {
            HSSFRow row = (NPOI.HSSF.UserModel.HSSFRow)sheet.CreateRow(rowIndex);
            for (int i = 0; i < vals.Length; i++)
            {
                row.CreateCell(i).SetCellValue(vals[i]);
            }
        }

        private HSSFWorkbook CreateWBA()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet1 = wb.CreateSheet("Sheet1");
            Sheet sheet2 = wb.CreateSheet("Sheet2");
            Sheet sheet3 = wb.CreateSheet("John's sales");

            CreateDataRow(sheet1, 0, 11, 12, 13, 14);
            CreateDataRow(sheet1, 1, 21, 22, 23, 24);
            CreateDataRow(sheet1, 2, 31, 32, 33, 34);

            CreateDataRow(sheet2, 0, 50, 55, 60, 65);
            CreateDataRow(sheet2, 1, 51, 56, 61, 66);
            CreateDataRow(sheet2, 2, 52, 57, 62, 67);

            CreateDataRow(sheet3, 0, 30, 31, 32);
            CreateDataRow(sheet3, 1, 33, 34, 35);
            return wb;
        }

        private HSSFWorkbook CreateWBB()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Sheet sheet1 = wb.CreateSheet("Sheet1");
            Sheet sheet2 = wb.CreateSheet("Sheet2");
            Sheet sheet3 = wb.CreateSheet("## Look here!");

            CreateDataRow(sheet1, 0, 400, 440, 480, 520);
            CreateDataRow(sheet1, 1, 420, 460, 500, 540);

            CreateDataRow(sheet2, 0, 50, 55, 60, 65);
            CreateDataRow(sheet2, 1, 51, 56, 61, 66);

            CreateDataRow(sheet3, 0, 42);

            return wb;
        }

        [TestMethod]
        public void TestBasic()
        {

            HSSFWorkbook wbA = CreateWBA();
            Cell c = wbA.GetSheetAt(0).CreateRow(5).CreateCell(2);
            HSSFFormulaEvaluator feA = new HSSFFormulaEvaluator(wbA);

            // non-error cases
            Confirm(feA, c, "INDIRECT(\"C2\")", 23);
            Confirm(feA, c, "INDIRECT(\"$C2\")", 23);
            Confirm(feA, c, "INDIRECT(\"C$2\")", 23);
            Confirm(feA, c, "SUM(INDIRECT(\"Sheet2!B1:C3\"))", 351); // area ref
            Confirm(feA, c, "SUM(INDIRECT(\"Sheet2! B1 : C3 \"))", 351); // spaces in area ref
            Confirm(feA, c, "SUM(INDIRECT(\"'John''s sales'!A1:C1\"))", 93); // special chars in sheet name
            Confirm(feA, c, "INDIRECT(\"'Sheet1'!B3\")", 32); // redundant sheet name quotes
            Confirm(feA, c, "INDIRECT(\"sHeet1!B3\")", 32); // case-insensitive sheet name
            Confirm(feA, c, "INDIRECT(\" D3 \")", 34); // spaces around cell ref
            Confirm(feA, c, "INDIRECT(\"Sheet1! D3 \")", 34); // spaces around cell ref
            Confirm(feA, c, "INDIRECT(\"A1\", TRUE)", 11); // explicit arg1. only TRUE supported so far

            Confirm(feA, c, "INDIRECT(\"A1:G1\")", 13); // de-reference area ref (note formula is in C4)


            // simple error propagation:

            // arg0 is Evaluated to text first
            Confirm(feA, c, "INDIRECT(#DIV/0!)", ErrorEval.DIV_ZERO);
            Confirm(feA, c, "INDIRECT(#DIV/0!)", ErrorEval.DIV_ZERO);
            Confirm(feA, c, "INDIRECT(#NAME?, \"x\")", ErrorEval.NAME_INVALID);
            Confirm(feA, c, "INDIRECT(#NUM!, #N/A)", ErrorEval.NUM_ERROR);

            // arg1 is Evaluated to bool before arg0 is decoded
            Confirm(feA, c, "INDIRECT(\"garbage\", #N/A)", ErrorEval.NA);
            Confirm(feA, c, "INDIRECT(\"garbage\", \"\")", ErrorEval.VALUE_INVALID); // empty string is not valid bool
            Confirm(feA, c, "INDIRECT(\"garbage\", \"flase\")", ErrorEval.VALUE_INVALID); // must be "TRUE" or "FALSE"


            // spaces around sheet name (with or without quotes Makes no difference)
            Confirm(feA, c, "INDIRECT(\"'Sheet1 '!D3\")", ErrorEval.REF_INVALID);
            Confirm(feA, c, "INDIRECT(\" Sheet1!D3\")", ErrorEval.REF_INVALID);
            Confirm(feA, c, "INDIRECT(\"'Sheet1' !D3\")", ErrorEval.REF_INVALID);


            Confirm(feA, c, "SUM(INDIRECT(\"'John's sales'!A1:C1\"))", ErrorEval.REF_INVALID); // bad quote escaping
            Confirm(feA, c, "INDIRECT(\"[Book1]Sheet1!A1\")", ErrorEval.REF_INVALID); // unknown external workbook
            Confirm(feA, c, "INDIRECT(\"Sheet3!A1\")", ErrorEval.REF_INVALID); // unknown sheet
            if (false)
            { // TODO - support Evaluation of defined names
                Confirm(feA, c, "INDIRECT(\"Sheet1!IW1\")", ErrorEval.REF_INVALID); // bad column
                Confirm(feA, c, "INDIRECT(\"Sheet1!A65537\")", ErrorEval.REF_INVALID); // bad row
            }
            Confirm(feA, c, "INDIRECT(\"Sheet1!A 1\")", ErrorEval.REF_INVALID); // space in cell ref
        }
        [TestMethod]
        public void TestMultipleWorkbooks()
        {
            HSSFWorkbook wbA = CreateWBA();
            Cell cellA = wbA.GetSheetAt(0).CreateRow(10).CreateCell(0);
            HSSFFormulaEvaluator feA = new HSSFFormulaEvaluator(wbA);

            HSSFWorkbook wbB = CreateWBB();
            Cell cellB = wbB.GetSheetAt(0).CreateRow(10).CreateCell(0);
            HSSFFormulaEvaluator feB = new HSSFFormulaEvaluator(wbB);

            String[] workbookNames = { "MyBook", "Figures for January", };
            HSSFFormulaEvaluator[] Evaluators = { feA, feB, };
            HSSFFormulaEvaluator.SetupEnvironment(workbookNames, Evaluators);

            Confirm(feB, cellB, "INDIRECT(\"'[Figures for January]## Look here!'!A1\")", 42); // same wb
            Confirm(feA, cellA, "INDIRECT(\"'[Figures for January]## Look here!'!A1\")", 42); // across workbooks

            // 2 level recursion
            Confirm(feB, cellB, "INDIRECT(\"[MyBook]Sheet2!A1\")", 50); // set up (and Check) first level
            Confirm(feA, cellA, "INDIRECT(\"'[Figures for January]Sheet1'!A11\")", 50); // points to cellB
        }

        private void Confirm(FormulaEvaluator fe, Cell cell, String formula,
                double expectedResult)
        {
            fe.ClearAllCachedResultValues();
            cell.CellFormula = formula;
            CellValue cv = fe.Evaluate(cell);
            if (cv.CellType != CellType.NUMERIC)
            {
                throw new AssertFailedException("expected numeric cell type but got " + cv.FormatAsString());
            }
            Assert.AreEqual(expectedResult, cv.NumberValue, 0.0);
        }
        private void Confirm(FormulaEvaluator fe, Cell cell, String formula,
                ErrorEval expectedResult)
        {
            fe.ClearAllCachedResultValues();
            cell.CellFormula = formula;
            CellValue cv = fe.Evaluate(cell);
            if (cv.CellType != CellType.ERROR)
            {
                throw new AssertFailedException("expected error cell type but got " + cv.FormatAsString());
            }
            int expCode = expectedResult.ErrorCode;
            if (cv.ErrorValue != expCode)
            {
                throw new AssertFailedException("Expected error '" + ErrorEval.GetText(expCode)
                        + "' but got '" + cv.FormatAsString() + "'.");
            }
        }
    }
}



