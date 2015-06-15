using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Converter.Template_workbooks;
using OfficeOpenXml;

namespace Formater
{
    class ExcelRowLocationDecoder
    {
        private readonly int lastRow;
        private readonly ExcelWorksheet worksheet;
        private readonly XlTemplateWorkbookType wbType;

        private ExcelLocationRow currentRow;
        private int rowIndex;

        public bool DecodeDescription { get; set; }

        public ExcelRowLocationDecoder(ExcelWorksheet worksheet, XlTemplateWorkbookType wbType)
        {
            DecodeDescription = true;

            this.worksheet = worksheet;
            this.wbType = wbType;
            lastRow = worksheet.Dimension.End.Row;
        }

//        public bool ReadNext()
//        {
//            if (rowIndex + 1 > lastRow) return false;
//
//            currentRow = new ExcelLocationRow(worksheet,++rowIndex,wbType,);
//            return true;
//        }

        public void DecodeRow()
        {
            DecodeSubjectCell();
            DecodeRegionCell();
            DecodeNearCityCell();

        }

        private void DecodeNearCityCell()
        {
            throw new NotImplementedException();
        }

        private void DecodeRegionCell()
        {
            throw new NotImplementedException();
        }

        public void DecodeSubjectCell()
        {
            throw new NotImplementedException();
        }
    }
}
