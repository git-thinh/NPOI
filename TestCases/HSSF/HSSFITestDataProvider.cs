﻿using System;
using System.Collections.Generic;
using System.Text;
using TestCases.SS;
using NPOI.HSSF.UserModel;
using NPOI.SS;
using NPOI.SS.UserModel;

namespace TestCases.HSSF
{
    public class HSSFITestDataProvider : ITestDataProvider
    {
        public Workbook OpenSampleWorkbook(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        public Workbook WriteOutAndReadBack(Workbook original)
        {
            if (!(original is HSSFWorkbook))
            {
                throw new ArgumentException("Expected an instance of HSSFWorkbook");
            }

            return HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)original);
        }

        public Workbook CreateWorkbook()
        {
            return new HSSFWorkbook();
        }

        public byte[] GetTestDataFileContent(String fileName)
        {
            return POIDataSamples.GetSpreadSheetInstance().ReadFile(fileName);
        }

        public SpreadsheetVersion GetSpreadsheetVersion()
        {
            return SpreadsheetVersion.EXCEL97;
        }

        private HSSFITestDataProvider() { }
        private static HSSFITestDataProvider inst = new HSSFITestDataProvider();
        public static HSSFITestDataProvider getInstance()
        {
            return inst;
        }
    }
}
