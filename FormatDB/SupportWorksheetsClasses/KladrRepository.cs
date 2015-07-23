using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Entity;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using AutoMapper.QueryableExtensions;
using ExcelRLibrary;
using OfficeOpenXml;
using REntities.Kladr;
using REntities.Kladr.DTO;
using REntities.Oktmo;

namespace Formater.SupportWorksheetsClasses
{
    public class KladrRepository:IDisposable
    {
        private readonly Dictionary<string,string> cachDictionary = new Dictionary<string, string>();
        private readonly KladrContext db;

        public KladrRepository(DataTable table)
        {
            const int streetColIndex = 1;
            const int strTypeColIndex = 2;
            const int codeColIndex = 4;
            const int subjColIndex = 5;
            const int regColIndex = 6;
            const int nearCityColIndex = 8;
            const int cityTypeShortColIndex = 8;
            const int nearCityTypeColIndex = 10;

            Rows = table.Select().Select(r => new KladrLineDTO()
            {
                Code = (r[codeColIndex - 1] ?? "").ToString(),
                Subject = (r[subjColIndex - 1] ?? "").ToString(),
                Region = (r[regColIndex - 1] ?? "").ToString(),
                CityName = (r[nearCityColIndex - 1] ?? "").ToString(),
                CityType = (r[nearCityTypeColIndex - 1] ?? "").ToString(),
                CityTypeShort = (r[cityTypeShortColIndex - 1] ?? "").ToString(),
                Street = (r[streetColIndex - 1] ?? "").ToString(),
                StreetType = (r[strTypeColIndex - 1] ?? "").ToString(),
            }).ToList().AsReadOnly();
        }

        public KladrRepository()
        {
            db = new KladrContext();
            Mapper.AddProfile<KladrToDtoProfile>();

            Rows = db.KladrLines.AsNoTracking().Project().To<KladrLineDTO>().ToList().AsReadOnly();
        }

        public  ReadOnlyCollection<KladrLineDTO> Rows { get; private set; }

        private IEnumerable<OktmoRowDTO> locations =  null; 
        public IEnumerable<OktmoRowDTO> Locations
        {
            get
            {
                if (locations == null)
                    LoadLocationsFromDb();
                return locations;
            }
        }

        private void LoadLocationsFromDb()
        {
            locations = db.MacroLocations.Select(location => new OktmoRowDTO
            {
                Subject = location.Subject.Text,
                Region = location.Region.Text,
                Settlement = new SettlementDTO(),
                City = new CityDTO { Name = location.City.CityName.Text, Type = location.City.CityType.Text }
            }).ToList();
        }

        private void LoadLocationsFromExcel()
        {
            const string wbPath = @"D:\KLADR.xlsx";

            const int cityCol = 3;
            const int typeCol = 5;
            const int subjCol = 7;
            const int regCol = 13;


            using (var pckg = new ExcelPackage(new FileInfo(wbPath)))
            {
                var wb = pckg.Workbook;
                var ws = wb.Worksheets.First();
                const int startRow = 2;
                var rowsCount = ws.Dimension.End.Row;
                var rowsEnum = Enumerable.Range(startRow, rowsCount - 1);

                locations = rowsEnum.Where(i =>
                {
                    var regVal = GetCellValue(ws.Cells[i, regCol]);
                    return regVal != "" && !regVal.Contains("муниц");
                } ).Select(i => new OktmoRowDTO
                {
                    Subject = GetCellValue(ws.Cells[i, subjCol]),
                    Region = GetCellValue(ws.Cells[i, regCol]),
                    City =
                        new CityDTO
                        {
                            Name = GetCellValue(ws.Cells[i, cityCol]),
                            Type = GetCellValue(ws.Cells[i, typeCol])
                        }
                }).ToList();


                ws.Dispose();
                wb.Dispose();
            }
        }

        private string GetCellValue(ExcelRange cell)
        {
            return (cell.Value ?? "").ToString();
        }



        public string GetStreetType(string street)
        {
            string type;
            if (cachDictionary.TryGetValue(street, out type))
                return type;

            var rows = Rows.Where(r => r.Street.EqualNoCase(street));
            if (rows.Count() != 1) return "";
            var city = Rows.First();

            type = city.CityType;
            cachDictionary.Add(street, type);

            return type;
        }

        public string GetStreetTypeFromCity(string street, string city)
        {
            string type;
            if (cachDictionary.TryGetValue(city+street, out type))
                return type;

            var rows = Rows.Where(r => r.Street.EqualNoCase(street) && r.CityName.EqualNoCase(city));
            if (rows.Count() != 1) return "";
            var row = Rows.First();

            type = row.StreetType;
            cachDictionary.Add(city+street, type);

            return type;
        }

        public string GetStreetTypeFromRegion(string street, string region)
        {
            string type;
            if (cachDictionary.TryGetValue(region + street, out type))
                return type;

            var rows = Rows.Where(r => r.Street.EqualNoCase(street) && r.Region.EqualNoCase(region)).ToArray();
            if (rows.Count() != 1) return "";
            var row = rows.First();

            type = row.StreetType;
            cachDictionary.Add(region + street, type);

            return type;
        }

        public bool IsStreetFromNearCity(string street, string city)
        {
            var result = Rows.Any(r => r.CityName.EqualNoCase(city) && r.Street.EqualNoCase(street));
            if (result) return true;

            return Rows.Any(r => r.Region.EndsWith(city) && r.Street.EqualNoCase(street));
        }

        public bool IsCityFromRegion(string city, string region)
        {
            return Rows.Any(r => r.Region.EqualNoCase(region) && r.CityName.EqualNoCase(city));
        }

        public bool IsStreetFromKladr(string street)
        {
            return Rows.Any(r => r.Street.EqualNoCase(street));
        }

        public string GetCityFullTypeFromShort(string shortType)
        {
            var row = Rows.SingleOrDefault(r => r.CityTypeShort.EqualNoCase(shortType));
            return row != null ? row.CityType : "";
        }

        private bool disposed = false;
        public void Dispose()
        {
            if (disposed) return;
            db.Dispose();
            disposed = true;
        }

        public bool IsStreetFromRegion(string street, string region)
        {
            return Rows.Any(r => r.Region.EqualNoCase(region) && r.Street.EqualNoCase(street));
        }

        public bool TypeFromBase(string value)
        {
            return Rows.Select(r => r.StreetType).Any(s => s.EqualNoCase(value));
        }
    }
}
