using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.HPSF;
using System.IO;

namespace EzXlsFinancial.Objects
{
    public class IRR
    {
        public IRR()
        {
            SetupWorksheet();
        }
        private HSSFWorkbook workbook = new HSSFWorkbook();
        private ISheet sheet;
        private string rateFormula = "B1";
        public string RateFormula
        {
            get
            {
                return rateFormula;
            }
            set
            {
                rateFormula = value;
            }
        }

        private int rowIndex = 0;
        private int maxRows = 5000;
        public static readonly string RATE_VAR = "{RATE}";
        private void SetupWorksheet()
        {
            this.workbook = new HSSFWorkbook();
            this.sheet = workbook.CreateSheet("Financial");
            rowIndex = 0;
            var row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("Rate");
            row.CreateCell(1).SetCellValue(0);
            row.CreateCell(2).SetCellValue("IRR");
            row.CreateCell(3).SetCellFormula(string.Format("IRR({0},B4:B51)", RateFormula.Replace(IRR.RATE_VAR, "B1")));
            rowIndex++; rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("Period");
            row.CreateCell(1).SetCellValue("Total");
            for (int i = 1; i < maxRows; i++)
            {
                rowIndex++;
                row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue(i);
                row.CreateCell(1).SetCellValue(0);
            }
        }
        public bool SaveToFile(string FileName)
        {
            using (var file2 = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
            {
                workbook.Write(file2);
                file2.Close();
            }
            return true;
        }
        public void Clear()
        {
            sheet.SetCellValue(0, 1, 0.0);
            for (int i = 3; i < maxRows+3; i++)
            {
                sheet.SetCellValue(i, 1, 0.0);
            }
        }
        public double Calculate(double rate, List<double> values)
        {
            if (values.Count > maxRows) throw new Exception(string.Format("Cannot handle values list over {0}!", values));
            this.Clear();
            sheet.SetCellValue(0, 1, rate);
            var currRow = 3;
            foreach (var value in values)
            {
                sheet.SetCellValue(currRow, 1, value);
                currRow++;
            }
            this.sheet.SetCellFormula(0, 3, string.Format("IRR(B4:B{0}, {1})*12", currRow, rateFormula.Replace(IRR.RATE_VAR, "B1")));
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            var irrValue = sheet.GetCellValue(0, 3, 0d);
            return irrValue;
        }
    }
}
