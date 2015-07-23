using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using ExcelRLibrary.Annotations;
using Formater.SupportWorksheetsClasses;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PatternsLib;
using REntities.Kladr;
using REntities.Kladr.DTO;
using REntities.Oktmo;
using KladrRepository = Formater.SupportWorksheetsClasses.KladrRepository;

namespace Formater
{
    public class ExcelLocationRow : IDisposable
    {
        private const string DashPattern = @"\s*\-\s*";
        private const string SpacePattern = @"\s+";
        private readonly ExcelWorksheet worksheet;
        private readonly SupportWorksheets supportWorksheets;
        private readonly int row;
        private readonly int descriptionColumn;

//        private DataTable customTable;
//        private DataTable subjectTable;
        private string regCenter;
        private string regName;

        private bool cellsFilled;
        private bool breakFromRow = false;
        private readonly List<int> rowsToDelete;
        private readonly int typeOfNearCityColumn;

        public ExcelLocationRow(ExcelWorksheet worksheet, int row, XlTemplateWorkbookType wbType,
            SupportWorksheets supportWorksheets)
            : this(wbType, supportWorksheets)
        {
            this.worksheet = worksheet;
            this.row = row;

            using (var db = new TemplateWbsContext())
            {
                oktmoComposition = new OktmoHelper();
                var columns = db.TemplateWorkbooks.First(w => w.WorkbookType == wbType).Columns.ToList();

                var subjColumn = (byte) columns.First(c => c.CodeName.Equals("SUBJECT")).ColumnIndex;
                var regionColumn = (byte) columns.First(c => c.CodeName.Equals("REGION")).ColumnIndex;
                var settlementColumn = (byte) columns.First(c => c.CodeName.Equals("SETTLEMENT")).ColumnIndex;
                var nearCityColumn = (byte) columns.First(c => c.CodeName.Equals("NEAR_CITY")).ColumnIndex;
                typeOfNearCityColumn = (byte) columns.First(c => c.CodeName.Equals("TERRITORY_TYPE")).ColumnIndex;
                var vgtColumn = (byte) columns.First(c => c.CodeName.Equals("VGT")).ColumnIndex;
                var streetColumn = (byte) columns.First(c => c.CodeName.Equals("STREET")).ColumnIndex;
                var typeOfStreetColumn = (byte) columns.First(c => c.CodeName.Equals("STREET_TYPE")).ColumnIndex;
                var sourceLinkColumn = (byte) columns.First(c => c.CodeName.Equals("URL_SALE")).ColumnIndex;
                var distToRegCenterColumn = (byte) columns.First(c => c.CodeName.Equals("DIST_REG_CENTER")).ColumnIndex;
                var distToNearCityColumn = (byte) columns.First(c => c.CodeName.Equals("DIST_NEAR_CITY")).ColumnIndex;
                var inCityColumn = (byte) columns.First(c => c.CodeName.Equals("IN_CITY")).ColumnIndex;
                var houseNumColumn = (byte) columns.First(c => c.CodeName.Equals("HOUSE_NUM")).ColumnIndex;
                var letterColumn = (byte) columns.First(c => c.CodeName.Equals("LETTER")).ColumnIndex;
                var sntKpDnpColumn = (byte) columns.First(c => c.CodeName.Equals("ASSOCIATIONS")).ColumnIndex;
                var additionalInfoColumn = (byte) columns.First(c => c.CodeName.Equals("ADDITIONAL")).ColumnIndex;
                var buildColumn = (byte) columns.First(c => c.CodeName.Equals("HOUSE_NUM")).ColumnIndex;
                descriptionColumn = (byte) columns.First(c => c.CodeName.Equals("DESCRIPTION")).ColumnIndex;


                SubjectCell = new DataCell(worksheet.Cells[row, subjColumn]);
                RegionCell = new DataCell(worksheet.Cells[row, regionColumn]);
                SettlementCell = new DataCell(worksheet.Cells[row, settlementColumn]);
                NearCityCell = new DataCell(worksheet.Cells[row, nearCityColumn]);
                VgtCell = new DataCell(worksheet.Cells[row, vgtColumn]);
                StreetCell = new DataCell(worksheet.Cells[row, streetColumn]);
                TypeOfNearCityCell = new DataCell(worksheet.Cells[row, typeOfNearCityColumn]);
                LandMarkCell = new DataCell(worksheet.Cells[row, additionalInfoColumn]);
                TypeOfStreetCell = new DataCell(worksheet.Cells[row, typeOfStreetColumn]);
                DistToRegCenterCell = new DataCell(worksheet.Cells[row, distToRegCenterColumn]);
                DictToNearCityCell = new DataCell(worksheet.Cells[row, distToNearCityColumn]);
                SntKpsCell = new DataCell(worksheet.Cells[row, sntKpDnpColumn]);
                InCityCell = new DataCell(worksheet.Cells[row, inCityColumn]);
                HouseNumCell = new DataCell(worksheet.Cells[row, houseNumColumn]);
                LetterCell = new DataCell(worksheet.Cells[row, letterColumn]);
                SourceLinkCell = new DataCell(worksheet.Cells[row, sourceLinkColumn]);
                BuildsCell = new DataCell(worksheet.Cells[row, buildColumn]);
            }

            SplitConcatenatedCellS();
        }

        private ExcelLocationRow(XlTemplateWorkbookType wbType, SupportWorksheets supportWorksheets)
        {
            oktmoComposition = new OktmoHelper();
            this.supportWorksheets = supportWorksheets;
            rowsToDelete = new List<int>();
            DoDescription = true;
        }

        public bool DoDescription { get; set; }

        #region DataCells

        public DataCell SubjectCell { get; set; }
        public DataCell RegionCell { get; set; }
        public DataCell SettlementCell { get; set; }
        public DataCell NearCityCell { get; set; }
        public DataCell TypeOfNearCityCell { get; set; }
        public DataCell VgtCell { get; set; }
        public DataCell StreetCell { get; set; }
        public DataCell TypeOfStreetCell { get; set; }
        public DataCell LandMarkCell { get; set; }
        public DataCell DistToRegCenterCell { get; set; }
        public DataCell DictToNearCityCell { get; set; }
        public DataCell SntKpsCell { get; set; }
        public DataCell InCityCell { get; set; }
        public DataCell HouseNumCell { get; set; }
        public DataCell LetterCell { get; set; }
        public DataCell SourceLinkCell { get; set; }
        public DataCell BuildsCell { get; set; }

        #endregion

        public void SaveCells()
        {
            SubjectCell.Save();
            RegionCell.Save();
            SettlementCell.Save();
            NearCityCell.Save();
            TypeOfNearCityCell.Save();
            VgtCell.Save();
            StreetCell.Save();
            TypeOfStreetCell.Save();
            LandMarkCell.Save();
            DictToNearCityCell.Save();
            DistToRegCenterCell.Save();
            SntKpsCell.Save();
            InCityCell.Save();

            HouseNumCell.Save();
            LetterCell.Save();
            SourceLinkCell.Save();
            BuildsCell.Save();
        }


        private readonly LocatonRegexpHandler regexpHandler = LocatonRegexpHandler.Init();
        private readonly OktmoHelper oktmoComposition;

        public void CheckRowForLocations()
        {
            CheckSubejctCell();
            if (breakFromRow) return;

            CheckRegionCell();
            if (breakFromRow) return;

            TryFillClassificator();
            CheckNearCityCell();
            if (breakFromRow) return;
            TryFillClassificator();

            CheckVgtCell();
            if (breakFromRow) return;

            if (!CheckStreetCell()) return;

            CheckLandmarkdsCell();
            if (breakFromRow) return;

            CheckDescriptionCell();
            if (breakFromRow) return;

            TryFillByStreet();
            FinalCheckLocationComposition();
            if (!cellsFilled)
            {
                FillDefaultValues();
                if (!cellsFilled)
                    FillOldvalues();
                FinalCheckLocationComposition();
            }

            SaveCells();
        }

        private bool CheckStreetCell()
        {
            var val = StreetCell.InitValue;
            TryFillStreet(ref val);
            StreetCell.InitValue = val;
            return true;
        }


        /// <summary>
        /// Метод проверяет все текущие ячейки на предмет связи между собой
        /// </summary>
        private void FinalCheckLocationComposition()
        {
            if (!oktmoComposition.HasEqualSubject(SubjectCell.Value)) return;
            CheckCell(RegionCell, oktmoComposition.HasEqualRegion);
            CheckCell(SettlementCell, oktmoComposition.HasEqualSettlement);
            CheckCell(NearCityCell, oktmoComposition.HasEqualNearCity);
            CheckCell(TypeOfNearCityCell, oktmoComposition.HasEqualCityType);

            Func<string, bool> streetFromComposition = (s) =>
            {
                if (s == "") return true;
                if (NearCityCell.Valid)
                    return supportWorksheets.Kladr.IsStreetFromNearCity(s, NearCityCell.Value);
                return RegionCell.Valid && supportWorksheets.Kladr.IsStreetFromRegion(s, RegionCell.Value);
            };
            CheckCell(StreetCell, streetFromComposition);

            Func<string, bool> vgtFromComposition = (s) =>
            {
                if (s == "" || RegionCell.Value == "") return true;
                if (RegionCell.Valid)
                    return supportWorksheets.VgtWorksheet.CombinationExists(RegionCell.Value, s);
                return supportWorksheets.VgtWorksheet.TerritotyExists(s);
            };
            CheckCell(VgtCell, vgtFromComposition);
        }

        private void CheckCell(DataCell cell, Func<string, bool> validation)
        {
            if (cell.Value == "") return;
            var status = validation(cell.Value)
                ? DataCell.DataCellStatus.Valid
                : DataCell.DataCellStatus.InValid;

            cell.SetStatus(status);
        }

        private void FillOldvalues()
        {
            var cells = new[]
            {
                SubjectCell, RegionCell, SettlementCell, NearCityCell
            };

            cells.ForEach(cell =>
            {
                if (cell.Valid) return;
                if (cell.InitValue == "") return;
                if (cell.Value != "") return;
                cell.SetDefaultValue();
                cell.SetStatus(DataCell.DataCellStatus.InValid);
            });
        }

        /// <summary>
        /// При наличии улицы метод будет пытаться заполнить недостающие значения
        /// </summary>
        private void TryFillByStreet()
        {
            if (cellsFilled) return;
            if (!StreetCell.Valid) return;

            IEnumerable<KladrLineDTO> rows = null;
            if (SubjectCell.Valid)
                rows = supportWorksheets.Kladr.Rows.Where(r => r.Subject.Equals(SubjectCell.Value));

            var srachEnum = rows ?? supportWorksheets.Kladr.Rows;
            if (RegionCell.Valid)
                rows = srachEnum.Where(r => r.Region.Equals(RegionCell.Value));

            srachEnum = rows ?? supportWorksheets.Kladr.Rows;
            if (NearCityCell.Valid)
                rows = srachEnum.Where(r => r.CityName.Equals(NearCityCell.Value));

            if (rows == null) return;

            rows = rows.Where(r => r.Street.Equals(StreetCell.Value));
            var kladrLineDtos = rows.ToArray();
            if (kladrLineDtos.Count() != 1) return;

            var kladrRow = kladrLineDtos.Single();

            if (!NearCityCell.Valid)
            {
                if (NearCityCell.Value == "")
                {
                    NearCityCell.Value = kladrRow.CityName;
                }
            }

            if (!RegionCell.Valid)
            {
                if (RegionCell.Value == "")
                {
                    RegionCell.Value = kladrRow.Region;
                }
            }
        }


        private void FillDefaultValues()
        {
            //Вписываем дефолтные значения Если населенный пункт так и не заполнен
            if (NearCityCell.Value == "")
            {
                //Находим дефолтный населенный пункт по ссылке на объявление
                var newCity =
                    supportWorksheets.SoubjectSourceWorksheet.GetDefaultNearCityByLink(SourceLinkCell.InitValue);

                if (!string.IsNullOrEmpty(newCity))
                {
                    //Мы пишем насел пункт только если он подходит к нашей выборке
                    //Т.е. подходит и к субъекту и к муниципальному образованию, есть таковой есть
                    if (oktmoComposition.HasEqualNearCity(newCity))
                    {
                        NearCityCell.Value = newCity;
                        TypeOfNearCityCell.Value = "город";

                        var spec = new NearCitySpecification(newCity);
                        oktmoComposition.SetSpecifications(spec);

                        TryFillClassificator();
                    }
                }
                //или ставим муниципальное образование как город
                //При условии что это не региональный центр
                else if (RegionCell.Value != "" &&
                         RegionCell.Valid
                         && RegionCell.Value.Contains("город", StringComparison.OrdinalIgnoreCase))
                {
                    if (RegionCell.Value == regName)
                    {
                        NearCityCell.Value = regCenter;
                        var spec = new NearCitySpecification(regName);
                        oktmoComposition.SetSpecifications(spec);
                        TryFillClassificator();
                    }
                    else
                    {
                        var name = Regex.Replace(RegionCell.Value, @"(^|\b)город(\b|$)", "").Trim();
                        name = name.Replace("(ЗАТО)", "");
                        if (oktmoComposition.HasEqualNearCity(name))
                        {
                            cellsFilled = false;
                            NearCityCell.Value = name;
                            TypeOfNearCityCell.Value = "город";

                            var spec = new NearCitySpecification(name);
                            oktmoComposition.SetSpecifications(spec);
                            TryFillClassificator();
                        }
                    }
                }
            }
            //Ставим дефолтное значение для муниципального образования, если оно пустое, а текущий насленный пункт у нас является региональным центро
            else if (RegionCell.Value == "" &&
                     string.Equals(NearCityCell.Value, regName, StringComparison.OrdinalIgnoreCase))
            {
                var spec = new NearCitySpecification(regName);
                var regSpec =
                    new ExpressionSpecification<OktmoRowDTO>(
                        oktmoRow =>
                            oktmoRow.Region.ToLower()
                                .Replace("город", "")
                                .Trim()
                                .Equals(regCenter.ToLower().Replace("город", "").Trim()));
                var totalSpec = spec.And(regSpec);
                oktmoComposition.SetSpecifications(totalSpec);

                TryFillClassificator();
            }
            //Дефолное значение для типа населенного пункта, если найденный насел пункт совпадает по названию с региональным центром
            else if (TypeOfNearCityCell.Value == "" &&
                     string.Equals(NearCityCell.Value, regName, StringComparison.OrdinalIgnoreCase))
            {
                TypeOfNearCityCell.Value = "город";
            }

            if (!RegionCell.Valid && NearCityCell.Value == regName)
            {
                oktmoComposition.ResetToSubject();
                var regnSpec = new NearCitySpecification(regName);
                oktmoComposition.SetSpecifications(regnSpec);

                var name = RegionCell.Value.Replace(" муниципальный район", "");
                if (oktmoComposition.CustomOktmoRows.Count == 1)
                {
                    if (supportWorksheets.VgtWorksheet.CombinationExists(
                        oktmoComposition.CustomOktmoRows.First().Region, name))
                    {
                        VgtCell.Value = name;
                        VgtCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                    else
                    {
                        AppendToLandMarkCell(RegionCell.Value);
                    }

                    RegionCell.Value = "";
                    TryFillClassificator();
                }
            }

            if (RegionCell.Value.EqualNoCase(regCenter))
            {
                if (oktmoComposition.CustomOktmoRows.Any(r => !r.Region.Equals(regName)))
                {
                    var spec = new RegionSpecification(regName);
                    oktmoComposition.SetSpecifications(spec);
                    TryFillClassificator();
                }
            }
            else if (NearCityCell.Value == regName)
            {
                if (RegionCell.Value != "")
                    AppendToLandMarkCell(RegionCell.Value);
                oktmoComposition.ResetToSubject();

                var regCenterSpecs = new NearCitySpecification(regName);
                var regNameSpecs = new RegionSpecification(regCenter);
                var totalSpecs = regCenterSpecs.And(regNameSpecs);
                oktmoComposition.SetSpecifications(totalSpecs);
                TryFillClassificator();
            }

            if (RegionCell.Value != "" || NearCityCell.Value != "") return;
            var superTips = new[] { "Москва", "Санкт-Петербург" };
            var eqName = superTips.FirstOrDefault(s => s.EqualNoCase(SubjectCell.Value));
            if (eqName == null) return;
            NearCityCell.Value = eqName;
            TypeOfNearCityCell.Value = "город";
        }

        private void CheckDescriptionCell()
        {
            if (!DoDescription) return;
            //Вначале мы ищем наименования по типу
            //После мы пытаемся отнести найдненные в описании Именования без типов
            var cell = worksheet.Cells[row, descriptionColumn];
            if ((string) cell.Value == "") return;

            var descrtContent = ReplaceYo((cell.Value ?? "").ToString()).Trim().Trim(',').Trim();


            //
            //----Товарищества
            //

            var match = regexpHandler.SntToLeftRegex.Match(descrtContent);
            while (match.Success)
            {
                //Берём только первое совпадение!
                var name = TemplateName(match.Groups["name"].Value);

                SntKpsCell.Value = SntKpsCell.Value == "" ||
                                   SntKpsCell.Value.Length < 3
                    ? name
                    : ", " + name;
                descrtContent = regexpHandler.SntToLeftRegex.Replace(descrtContent, ", ").Trim().Trim(',').Trim();
                match = match.NextMatch();
            }

            TryFillStreet(ref descrtContent);

            //
            //---Субъект для сравнение с проставленным
            //
            var tmpRegex = regexpHandler.SubjRegEx;
            match = tmpRegex.Match(descrtContent);
            if (!match.Success)
            {
                tmpRegex = regexpHandler.SubjToLeftRegex;
                match = tmpRegex.Match(descrtContent);
            }
            if (match.Success)
            {
                //Собственно это главное, зачем мы входили в это условие. Исключаем Субъект для дальнейшего облегчения поиска других типов
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ");
                var fullName = supportWorksheets.OKTMOWs.GetFullName(
                    TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumn.Subject);

                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value != "" &&
                    !string.Equals(SubjectCell.Value.Trim(), fullName.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    rowsToDelete.Add(row);
                    SubjectCell.Value = fullName;

                    oktmoComposition.SetSubjectRows(supportWorksheets.OKTMOWs.GetSubjectRows(fullName).ToList());
                    oktmoComposition.ResetToSubject();


                    breakFromRow = true;
                    return;
                }
            }

            //
            //----Населенный пункт
            //
            var switched = false;
            var endChanged = false;
            var regs = new List<Regex> {regexpHandler.NearCityToLeftRegex, regexpHandler.NearCityRegex};
            foreach (var regi in regs)
            {
                var reg = regi;

                var matches = reg.Matches(descrtContent);

                if (matches.Count == 0) continue;

                match = null;
                //Приоритет у любого негорода
                if (matches.Count > 1)
                {
                    //Приорите у любого негорода без рассстояния
                    match =
                        matches.Cast<Match>()
                            .FirstOrDefault(
                                m =>
                                    (m.Groups["type"].Value.IndexOf("г",
                                        StringComparison.OrdinalIgnoreCase) ==
                                     -1) && string.IsNullOrEmpty(m.Groups["out"].Value));

                    //Если все насел пункты в удаленности
                    //Взять хотя бы с точной удаленность
                    if (match == null)
                        match =
                            matches.Cast<Match>()
                                .FirstOrDefault(
                                    m =>
                                        (m.Groups["type"].Value.IndexOf("г",
                                            StringComparison.OrdinalIgnoreCase) ==
                                         -1) && !string.IsNullOrEmpty(m.Groups["dist"].Value));

                    //Если все насел пункты с приблизительной удалённостью
                    //Взять хотя бы не город
                    if (match == null)
                        match =
                            matches.Cast<Match>()
                                .FirstOrDefault(
                                    m =>
                                        (m.Groups["type"].Value.IndexOf("г",
                                            StringComparison.OrdinalIgnoreCase) == -1));

                    //Если все города
                    //Хотя бы не региональный центр
                    if (match == null)
                        match =
                            matches.Cast<Match>()
                                .FirstOrDefault(
                                    m =>
                                        (!string.Equals(m.Groups["name"].Value, regName,
                                            StringComparison.OrdinalIgnoreCase)));

                    //Если все города, берём первый
                    if (match == null)
                        match = matches[0];
                }

                if (match == null)
                    match = matches[0];


                var name = ReplaceYo(TemplateName(match.Groups["name"].Value));
                var type = ReplaceYo(TryDescriptTypeOfNasPunkt(match.Groups["type"].Value));

                if (NearCityCell.Value.EqualNoCase(name))
                {
                    descrtContent = reg.Replace(descrtContent, ", ");
                    continue;
                }

                if (!string.IsNullOrEmpty(match.Groups["out"].Value))
                {
                    InCityCell.Value = "нет";
                    if (!string.IsNullOrEmpty(match.Groups["dist"].Value))
                    {
                        var dist = TryDescriptDistance(match.Groups["dist"].Value);
                        if (string.Equals(name, regName, StringComparison.OrdinalIgnoreCase))
                        {
                            //Backup current value
                            if (DictToNearCityCell.Value != "")
                                LandMarkCell.Value +=
                                    string.Format("Расстояние до регионального центра \"{0}\"",
                                        DistToRegCenterCell.Value);
                            DistToRegCenterCell.Value = dist.ToString();
                        }
                        else
                            DictToNearCityCell.Value = dist.ToString();
                    }
                }

                var splitted = false;
                List<string> words = null;
                var startName = name;
                tryGetNearCityAgain:

                //Опеределяем нужно ли обрабатывать найденную информацию
                if (name.EqualNoCase(regName) && RegionCell.Valid) continue;

//                if(!string.IsNullOrEmpty(regName) && name.EqualNoCase(regName)) continue;

                if (NearCityCell.Value != "" && NearCityCell.Value.EqualNoCase(name)) continue;

                if (SubjectCell.Valid && RegionCell.Valid && NearCityCell.Valid &&
                    !(type != "город" && TypeOfNearCityCell.Value == "город")) continue;

                if (type == "город" && TypeOfNearCityCell.Value != "" &&
                    TypeOfNearCityCell.Value != "город")
                {
                    LandMarkCell.Value += name + " " + type + ", ";
                }
                else
                {
                    //BackUp current value
                    if (NearCityCell.Value != "")
                        LandMarkCell.Value += TypeOfNearCityCell.Value + " " +
                                              NearCityCell.Value;

                    //Обнуляем МунОбразование
                    //сейчас стоит региональный центр или просто город
                    //а найденный насел пункт подходит к другому мун образования
                    var itIsCity = CurrentRegionIsCity || CurrentSettlIsCity || CurrentNearCityIsCity;

                    var valueNeedsResetRegion = !oktmoComposition.HasEqualNearCity(name) &&
                                                oktmoComposition.SubjectHasEqualNearCity(name);

                    if (itIsCity && valueNeedsResetRegion)
                    {
                        oktmoComposition.ResetToSubject();

                        RegionCell.Value = string.Empty;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        SettlementCell.Value = string.Empty;
                        SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                    }

                    //найденный насел пункт подхоидт к нашей выборке (по субъекту и возможно по мунобразованию если оно есть)
                    if (oktmoComposition.HasEqualNearCity(name.ToLower()))
                    {
                        var spec = new NearCitySpecification(name);
                        oktmoComposition.SetSpecifications(spec);

                        NearCityCell.Value = name;
                        TypeOfNearCityCell.Value = type;
                        cellsFilled = false;
                    }
                    else
                    {
                        if (!switched)
                        {
                            if (SwitchDashBetweenWords(ref name))
                            {
                                switched = true;
                                goto tryGetNearCityAgain;
                            }
                        }

                        if (!endChanged)
                        {
                            if (CityNameEndChanged(type, ref name))
                            {
                                endChanged = true;
                                goto tryGetNearCityAgain;
                            }
                        }
                        //Дробим имя собственное если возможно для поиска по каждому имени отдельни
                        if (!splitted)
                        {
                            //Step one: we split it
                            if (words == null)
                            {
                                var patterns = new List<string> {DashPattern, SpacePattern};

                                foreach (var pattern in patterns.Where(pattern => Regex.IsMatch(startName, pattern)))
                                {
                                    words = Regex.Split(startName, pattern).ToList();
                                    name = words.Last();
                                    words[words.Count - 1] = null;
                                    goto tryGetNearCityAgain;
                                }
                            }

                            //Step two: we use it untill end
                            else
                            {
                                for (var i = words.Count - 1; i >= 0; i--)
                                {
                                    if (words[i] == null) continue;
                                    name = words[i];
                                    words[i] = null;
                                    goto tryGetNearCityAgain;
                                }
                            }
                        }
                        if (((CurrentNearCityIsCity || NearCityCell.Value == "") && !type.EqualNoCase("город")) ||
                            NearCityCell.Value != "" && !NearCityCell.Valid && TypeOfNearCityCell.Value == "")
                        {
                            NearCityCell.Value = name;
                            TypeOfNearCityCell.Value = type;
                        }
                    }
                }

                descrtContent = reg.Replace(descrtContent, ", ");
            }
            //
            //----Муниципальное образование
            //
            if (!cellsFilled)
            {
                regs = new List<Regex> {regexpHandler.RegionRegex, regexpHandler.RegionToLeftRegex};
                foreach (var regi in regs)
                {
                    TryFillRegion(ref descrtContent, regi);
                }
            }
        }

        private static bool CityNameEndChanged(string type, ref string name)
        {
            const string cityEnd = @"(е|а)\b";
            if (type != "город" || !Regex.IsMatch(name, cityEnd, RegexOptions.IgnoreCase)) return false;
            name = Regex.Replace(name, cityEnd, "");
            return true;
        }

        private static bool SwitchDashBetweenWords(ref string name)
        {
            if (Regex.IsMatch(name, DashPattern))
            {
                name = Regex.Replace(name, DashPattern, " ");
                return true;
            }

            if (Regex.IsMatch(name, SpacePattern))
            {
                name = Regex.Replace(name, SpacePattern, "-");
                return true;
            }
            return false;
        }

        private bool CurrentNearCityIsCity
        {
            get
            {
                if (!NearCityCell.Valid) return false;
                return TypeOfNearCityCell.Valid && TextContainsCityAsType(TypeOfNearCityCell.Value);
            }
        }

        private bool CurrentSettlIsCity
        {
            get { return SettlementCell.Valid && TextContainsCityAsType(SettlementCell.Value); }
        }

        private bool CurrentRegionIsCity
        {
            get { return RegionCell.Valid && TextContainsCityAsType(RegionCell.Value); }
        }

        private static bool TextContainsCityAsType(string value)
        {
            return value.ToLower().Contains("город");
        }

        private void CheckLandmarkdsCell()
        {
            if (string.IsNullOrEmpty(LandMarkCell.InitValue)) return;
            var value = LandMarkCell.InitValue;
            //Поиск мун образвания
            if (RegionCell.Value == "")
            {
                TryFillRegion(ref value);
                LandMarkCell.InitValue = value;
            }

            //поиск улицы
            if (StreetCell.Value == "")
            {
                TryFillStreet(ref value);
                LandMarkCell.InitValue = value;
            }

            //Поиск внутрегородской территории
            //Bug а не происходит ли такая же процедура в методе поиска мунОбразования (см 6 строк выше)
            if (VgtCell.Value == "")
            {
                var tmpMatch = regexpHandler.RegionRegex.Match(LandMarkCell.InitValue);
                if (tmpMatch.Success)
                {
                    var tmpValue = tmpMatch.Groups["name"].Value;
                    TryFillVGT(ref tmpValue);
                    LandMarkCell.InitValue = LandMarkCell.InitValue.Replace(tmpValue, ", ").Trim().Trim(',').Trim();
                }
            }

            if (NearCityCell.Value == "")
            {
//                TryFindProperName(ref value);
//                LandMarkCell.InitValue = value;
            }

            //обработка имен собственных
            TryFindProperName(ref value);

            if (LandMarkCell.Value.Length > 2)
                LandMarkCell.Value += LandMarkCell.InitValue + ", ";
        }

        private void CheckVgtCell()
        {
            if (string.IsNullOrEmpty(VgtCell.InitValue)) return;
            var value = VgtCell.InitValue;

            var tmpMatch = regexpHandler.WordWithHeadLetteRegex.Match(value);

            if (!tmpMatch.Success) return;

            var tmpValue = tmpMatch.Value;

            TryFillVGT(ref tmpValue);

            value = value.Replace(tmpValue, ", ").Trim().Trim(',').Trim();
            VgtCell.InitValue = value;
        }

        private void CheckNearCityCell()
        {
            //
            //Разбираем Населенный пункт
            //
            var value = NearCityCell.InitValue;

            //Удаляем дублируемуб инфомарцию о субъекте
//            if (SubjectCell.InitValue != "")
//            value = value.Replace(SubjectCell.InitValue, ", ").Trim().Trim(',').Trim();

            if (string.IsNullOrEmpty(value)) return;

            if (NearCityCell.Valid && NearCityCell.Value != "")
                value = value.Replace(NearCityCell.Value, "");

            if (string.IsNullOrEmpty(value)) return;

            if (regexpHandler.SignleLetterPerStringRegex.IsMatch(value.Trim()))
            {
                if (!TryAppendPropNameToNearCity(value))
                {
                    TryFindProperName(ref value);
                    NearCityCell.InitValue = value;
                    return;
                }
                NearCityCell.InitValue = "";
                return;
            }

            //Ищем субъект для сравнение с проставленным
            var tmpRegex = regexpHandler.SubjRegEx;
            var match = tmpRegex.Match(value);
            if (!match.Success)
            {
                tmpRegex = regexpHandler.SubjToLeftRegex;
                match = tmpRegex.Match(value);
            }
            if (match.Success)
            {
                //Собственно это главное, зачем мы входили в это условие. Исключаем Субъект для дальнейшего облегчения поиска других типов
                NearCityCell.InitValue = tmpRegex.Replace(NearCityCell.InitValue, ", ").Trim().Trim(',').Trim();
                var fullName = supportWorksheets.OKTMOWs.GetFullName(
                    TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumn.Subject);

                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value != "" &&
                    SubjectCell.Value
                        .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                {
                    rowsToDelete.Add(row);
                    SubjectCell.Value += " <=> " + fullName;
                    SubjectCell.SetStatus(DataCell.DataCellStatus.InValid);

                    AppendToLandMarkCell(fullName);

                    breakFromRow = true;
                    return;
                }
            }

            if (NearCityCell.InitValue == "") return;

            //Поиск муниципального образования
            tmpRegex = regexpHandler.RegionRegex;
            match = tmpRegex.Match(value); // "Дальнево р-н"
            if (!match.Success)
            {
                tmpRegex = regexpHandler.RegionToLeftRegex;
                match = tmpRegex.Match(value); // "р-н Дальнево"
            }
            if (match.Success)
            {
                var name = TemplateName(match.Groups["name"].Value);
                var type = match.Groups["type"].Value;
                if (type.IndexOf("г", StringComparison.OrdinalIgnoreCase) >= 0)
                    type = "город";
                else if (type.IndexOf("р", StringComparison.OrdinalIgnoreCase) >= 0)
                    type = "район";
                else if (type.IndexOf("б", StringComparison.OrdinalIgnoreCase) >= 0)
                    type = "область";


                //bug перенести прверку на ВГТ сюда?
                //bug двойной поиск ВГТ
                var fullName = supportWorksheets.OKTMOWs.GetFullName(type == "город" ? type + " " + name : name,
                    OKTMOColumn.Region);
                if (!string.IsNullOrEmpty(fullName)) //This is REGION
                {
                    if (oktmoComposition.HasEqualRegion(fullName))
                    {
                        RegionCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;

                        //Выборка
                        var spec = new RegionSpecification(fullName);
                        oktmoComposition.SetSpecifications(spec);
                    }
                    else
                    {
                        RegionCell.Valid = false;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }

                    //В зависимости заполнен ли уже Регион, пишем извлеченное значение в ячейку Региона или ДопИнформации
                    if (RegionCell.Value == "")
                        RegionCell.Value = fullName;
                    else if (
                        !string.Equals(fullName, RegionCell.Value,
                            StringComparison.OrdinalIgnoreCase))
                        //Ситуция когда при обработке столбца "Регион" мы уже нашли более менее подходящее к субъекту мун.образование
                        //И при обработке населн пункта (данный процесс) мы тоже нашли подходящее к субъекту мун.образование
                        LandMarkCell.Value = fullName + LandMarkCell.Value + ", ";
                }
                //------------Try Apeend to VGT-----------
                else if (!TryFillVGT(ref name))
                {
                    fullName = name + " " + type;

                    //В зависимости заполнен ли уже Регион, пишем извлеченное значение в ячейку Региона или ДопИнформации
                    if (RegionCell.Value == "")
                        RegionCell.Value = fullName;
                    else if (fullName != RegionCell.Value)
                        LandMarkCell.Value = fullName + LandMarkCell.Value + ", ";
                }
                value = tmpRegex.Replace(value, ", ").Trim().Trim(',').Trim();
                NearCityCell.InitValue = value;
                if (value.Length <= 2) return;
            }

            NearCityCell.InitValue = value;
            if (value.Length < 3) return;

            //Поиск киллометров до населенного пункта
            match = regexpHandler.DistToRegCenteRegex.Match(value);
            if (match.Success)
            {
                //Спихиваем всё в столбец "Расстояние до рег центра"
                //Разбирать будем в конце
                DistToRegCenterCell.Value += ", " + match.Value;
                value = regexpHandler.DistToRegCenteRegex.Replace(value, ", ").Trim().Trim(',').Trim();
            }

            //Поиск улиц
            TryFillStreet(ref value);

            //Поиск поселения
            match = regexpHandler.SettlementRegex.Match(value);
            if (!match.Success)
            {
                match = regexpHandler.SettlemenToLeftRegex.Match(value);
            }
            if (match.Success)
            {
                var name = TemplateName(match.Groups["name"].Value);
                var type = match.Groups["type"].Value;

                type = type.IndexOf("п", StringComparison.OrdinalIgnoreCase) >= 0 ? "сельское поселение" : "сельсовет";

                var fullName = name + " " + type;
                if (SettlementCell.Value == "")
                    SettlementCell.Value = fullName;
                else
                    LandMarkCell.Value += fullName + ", ";

                if (oktmoComposition.HasEqualSettlement(fullName.ToLower()))
                {
                    var spec = new SettlementSpecification(fullName);
                    oktmoComposition.SetSpecifications(spec);
                }
                else
                {
                    SettlementCell.SetStatus(DataCell.DataCellStatus.InValid);

                    if (RegionCell.Value != "")
                    {
                        RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                    }
                    else if (NearCityCell.Value != "")
                    {
                        NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                    }
                }
                var stringReplace = ",";

                if (match.Value.Trim().StartsWith("("))
                    stringReplace = "";

                value = value.Replace(match.Value, stringReplace).Trim().Trim(',').Trim();
            }

            NearCityCell.InitValue = value;
            if (value.Length < 3) return;

            //Поиск 100% дополнительной инфомрации (снт, сот, с/н)
            tmpRegex = regexpHandler.SntToLeftRegex;
            match = tmpRegex.Match(value); //"пурум снт"
            if (!match.Success)
            {
                //"снт Пурум"
                tmpRegex = regexpHandler.SntRegex;
                match = tmpRegex.Match(value);
            }
            if (match.Success)
            {
                while (match.Success)
                {
                    var name = TemplateName(match.Groups["name"].Value);
                    var type = match.Groups["type"].Value;
                    if (type == "дп")
                    {
                        if (oktmoComposition.HasEqualNearCity(name))
                        {
                            NearCityCell.Value = name;
                            TypeOfNearCityCell.Value = "дачный поселок";

                            var spec = new NearCitySpecification(name);
                            oktmoComposition.SetSpecifications(spec);
                            continue;
                        }
                    }

                    SntKpsCell.Value = SntKpsCell.Value == "" ? name : ", " + name;
                    value = value.Replace(match.Value, "").Trim().Trim(',').Trim();
                    match = match.NextMatch();
                }
            }

            NearCityCell.InitValue = value;
            if (value.Length < 3) return;

            //Поиск населенного пункта
            if (oktmoComposition.HasEqualNearCity(value))
            {
                var spec = new NearCitySpecification(value);
                oktmoComposition.SetSpecifications(spec);
                NearCityCell.Value = TemplateName(value);
                value = "";
            }
            else
            {
                tmpRegex = regexpHandler.NearCityToLeftRegex;
                var matches = tmpRegex.Matches(value); // "Дальнево с."
                var switched = false;
                if (matches.Count == 0)
                {
                    tmpRegex = regexpHandler.NearCityRegex;
                    matches = tmpRegex.Matches(value); // "с. Дальнево"
                }
                if (matches.Count > 0)
                {
                    if (!cellsFilled)
                    {
                        //Приоритет у любого негорода
                        if (matches.Count > 1)
                        {
                            match =
                                matches.Cast<Match>()
                                    .FirstOrDefault(
                                        m => !Regex.IsMatch(m.Groups["type"].Value, "\bг", RegexOptions.IgnoreCase)) ??
                                //Приорите у любого негорода
                                matches[0]; //Конечно если он есть, если его нет, берём первое совпадение
                        }
                        else
                            match = matches[0];

                        var name = TemplateName(match.Groups["name"].Value);
                        var type = TryDescriptTypeOfNasPunkt(match.Groups["type"].Value);

                        tryAgainNCInNC:
                        //Если мы впервые нашлим населенный пункт
                        if (NearCityCell.Value == "")
                        {
                            NearCityCell.Value = name;
                            TypeOfNearCityCell.Value = type;

                            if (oktmoComposition.HasEqualNearCity(name.ToLower()))
                            {
                                var spec = new NearCitySpecification(name);
                                oktmoComposition.SetSpecifications(spec);

                                RegionCell.Valid = true;
                                NearCityCell.Valid = true;
                                NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                            }
                            else
                            {
                                if (!switched)
                                {
                                    switched = true;
                                    if (name.Contains("-"))
                                    {
                                        name = name.Replace("-", " ");
                                        goto tryAgainNCInNC;
                                    }
                                    if (name.Contains(" "))
                                    {
                                        name = name.Replace(" ", "-");
                                        goto tryAgainNCInNC;
                                    }
                                }
                                NearCityCell.Valid = false;
                                RegionCell.Valid = false;
                                NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                        }
                        else if (NearCityCell.Value != name) //нашли ли мы новую информацию
                        {
                            //и подходит ли она к нам
                            if (oktmoComposition.HasEqualNearCity(name.ToLower()))
                            {
                                var spec = new NearCitySpecification(name);
                                oktmoComposition.SetSpecifications(spec);

                                LandMarkCell.Value += NearCityCell.Value + ", ";

                                NearCityCell.Valid = true;
                                RegionCell.Valid = true;
                                NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;

                                NearCityCell.Value = name;
                                TypeOfNearCityCell.Value = type;
                            }
                            else
                            {
                                if (!switched)
                                {
                                    switched = true;
                                    if (name.Contains("-"))
                                    {
                                        name = name.Replace("-", " ");
                                        goto tryAgainNCInNC;
                                    }

                                    if (name.Contains(" "))
                                    {
                                        name = name.Replace(" ", "-");
                                        goto tryAgainNCInNC;
                                    }
                                }
                                LandMarkCell.Value += name + " " + type + ", ";
                            }
                        }
                    }
                    value = tmpRegex.Replace(value, ", ").Trim().Trim(',').Trim();
                }

                //Обрабатываем имена собственные
                else
                {
                    if (!TryAppendPropNameToNearCity(value))
                        TryFindProperName(ref value);
                    else
                        value = "";
                }
            }

            NearCityCell.InitValue = value;
            //Если у нас что-то не разобрано, мы его пихаем в доп инфо или ту же ячейек
            if (NearCityCell.InitValue.Length > 2)
            {
                LandMarkCell.Value += NearCityCell.InitValue + ", ";
            }
        }

        private void CheckRegionCell()
        {
            if (string.IsNullOrEmpty(RegionCell.InitValue)) return;

            //Удаляем дублируем инфомарцию о субъекте из ячейки мун образование
            if (!string.IsNullOrEmpty(SubjectCell.InitValue))
                RegionCell.InitValue = RegionCell.InitValue.Replace(SubjectCell.InitValue, ", ").Trim().Trim(',').Trim();

            if (string.IsNullOrEmpty(RegionCell.InitValue)) return;

            if (NearCityCell.InitValue == RegionCell.InitValue)
            {
                NearCityCell.InitValue = "";
            }

            //Ищем СУБЪЕКТ для сравнение с текущим
            var tmpRegex = regexpHandler.SubjRegEx;
            var match = tmpRegex.Match(RegionCell.InitValue);
            if (!match.Success)
            {
                tmpRegex = regexpHandler.SubjToLeftRegex;
                match = tmpRegex.Match(RegionCell.InitValue);
            }

            //далее при определенных условиях мы помечаем строку на удаление
            if (match.Success)
            {
                //удаляем из строки найденное
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ").Trim().Trim(',').Trim();

                var fullName = supportWorksheets.OKTMOWs.GetFullName(
                    TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumn.Subject);

                //мы должны определить полное название субъекты, найденного в ячейке муниципального Образования
                //И если у нас уже стоит какие-либо шаблонное значение в субеъкте
                //и оно не сходится с найденным в муниц образовании
                //помечаем столбец на удаление
                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value == "" &&
                    SubjectCell.Value
                        .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                {
                    //потому что продавцы пишут одно а по факту другое
                    rowsToDelete.Add(row);
                    SubjectCell.Value = fullName;

                    breakFromRow = true;
                    return;
                }
            }

            if (RegionCell.InitValue == "") return;

            var value = RegionCell.InitValue;
            TryFillRegion(ref value);
            RegionCell.InitValue = value;
            if (RegionCell.InitValue.Length <= 2) return;

            //На наличие поселения
            match = regexpHandler.SettlementRegex.Match(RegionCell.InitValue);
            if (!match.Success)
            {
                match = regexpHandler.SettlemenToLeftRegex.Match(RegionCell.InitValue);
            }
            //Если есть совпадение и оно не на всю строку
            if (match.Success)
            {
                var name = TemplateName(match.Groups["name"].Value);
                var type = match.Groups["type"].Value;
                type = type.IndexOf("п", StringComparison.OrdinalIgnoreCase) >= 0
                    ? "сельское поселение"
                    : "сельсовет";

                var fullName = name + " " + type;
                SettlementCell.Value = fullName;

                //В выборке уже имеется субъект и возможно Регион(или ВГТ)
                if (oktmoComposition.HasEqualSettlement(fullName))
                {
                    var spec = new SettlementSpecification(fullName);
                    oktmoComposition.SetSpecifications(spec);
                }
                else
                {
                    SettlementCell.Valid = false;
                    SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    if (RegionCell.Value == "")
                    {
                        RegionCell.Valid = false;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else if (NearCityCell.Value == "")
                    {
                        NearCityCell.Valid = false;
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                }

                RegionCell.InitValue = RegionCell.InitValue.Replace(match.Value, ", ").Trim().Trim(',').Trim();
            }


            //Поиск  товарищств
            tmpRegex = regexpHandler.SntRegex;
            match = tmpRegex.Match(RegionCell.InitValue);
            if (!match.Success)
            {
                tmpRegex = regexpHandler.SntToLeftRegex;
                match = tmpRegex.Match(RegionCell.InitValue);
            }
            if (match.Success)
            {
                var newName = TemplateName(match.Groups["name"].Value);
                SntKpsCell.Value = SntKpsCell.Value == "" ? newName : ", " + newName;
                if (match.Groups["type"].Value != "дп")
                    RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ").Trim().Trim(',').Trim();
            }


            //На наличие населенного пункта и его типа
            var cityRegexes = new[] {regexpHandler.NearCityRegex, regexpHandler.NearCityToLeftRegex};

            var cityMatches =
                new List<Match>(
                    cityRegexes.SelectMany(r => r.Matches(RegionCell.InitValue).Cast<Match>().Select(m =>
                    {
                        RegionCell.InitValue = RegionCell.InitValue.Replace(m.Value, ", ").Trim().Trim(',').Trim();
                        return m;
                    }).ToList()));
            var switched = false;

            //Если есть совпадение
            if (cityMatches.Count > 0)
            {
                //Приоритет у любого негорода
                //если таковой есть
                if (cityMatches.Count > 1)
                {
                    match = GetSingleOktmoSuitedMatch(cityMatches) ?? GetNonCentralCityTypeMatch(cityMatches);
                }
                else
                    match = cityMatches.First();

                var name = TemplateName(match.Groups["name"].Value);
                var type = TryDescriptTypeOfNasPunkt(match.Groups["type"].Value);
                tryAgainNC:

                //В выборке уже имеется Субъект и вохможно Регион(или ВГТ) и возможно поселение
                //Урезаем выборку если возможно
                if (oktmoComposition.HasEqualNearCity(name.ToLower()))
                {
                    var spec = new NearCitySpecification(name);
                    oktmoComposition.SetSpecifications(spec);
                }
                else
                {
                    if (!switched)
                    {
                        switched = true;
                        if (name.Contains("-"))
                        {
                            name = name.Replace("-", " ");
                            goto tryAgainNC;
                        }
                        if (name.Contains(" "))
                        {
                            name = name.Replace(" ", "-");
                            goto tryAgainNC;
                        }
                    }

                    //Поиск только по субъекту, если в более частной выборке совпадений по населенному пункту не нашлось
                    if (oktmoComposition.HasEqualNearCity(name.ToLower()))
                    {
                        var spec = new NearCitySpecification(name);
                        var newCustomRowsList = oktmoComposition.CustomOktmoRows.FindAll(r => spec.IsSatisfiedBy(r));
                        //Обновляем тип по найденному нас пункту если возможно
                        if (newCustomRowsList.Count == 1)
                        {
                            var newType = newCustomRowsList.First().City.Type;

                            RegionCell.Valid = true;
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;

                            if (TypeOfNearCityCell.Value == "" || TypeOfNearCityCell.Value != newType)
                            {
                                type = newType;
                            }
                            //bug если в таблице 1 запись, может уже записать всё?
                        }
                    }
                    else
                    {
                        NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                        SubjectCell.SetStatus(DataCell.DataCellStatus.InValid);
                    }
                }

                NearCityCell.Value = name; //Пишем найденное наименование в нужную ячейку
                TypeOfNearCityCell.Value = type;

                cityMatches.ForEach(match1 =>
                    RegionCell.InitValue = RegionCell.InitValue.Replace(match1.Value, ", ").Trim().Trim(',').Trim());

                if (RegionCell.InitValue.Length <= 2) return;
            }

            //Для улиц
            var initValue = RegionCell.InitValue;
            TryFillStreet(ref initValue);
            RegionCell.InitValue = initValue;


            //Имена собственные
            var s = RegionCell.InitValue;
            TryFindProperName(ref s);
            RegionCell.InitValue = s;

            //Ту информацию, что мы не смогли разобрать вписываем в отдельную ячейку
            if (RegionCell.InitValue.Length > 2)
                LandMarkCell.Value += RegionCell.InitValue + ", ";
        }

        private Match GetNonCentralCityTypeMatch(List<Match> cityMatches)
        {
            return cityMatches
                .FirstOrDefault(m => !m.Groups["name"].Value.Equals(regCenter)) ??
                   cityMatches
                       .FirstOrDefault(
                           m => !Regex.IsMatch(m.Groups["type"].Value, "\bг", RegexOptions.IgnoreCase)) ??
                   cityMatches[0];
        }

        private Match GetSingleOktmoSuitedMatch(List<Match> cityMatches)
        {
            cityMatches = cityMatches.DistinctBy(m => m.Groups["name"].Value).ToList();
            return cityMatches.FirstOrDefault(
                m => oktmoComposition.HasEqualNearCity(TemplateName(m.Groups["name"].Value)));
        }

        private void TryFillRegion(ref string content, Regex reg = null)
        {
            Regex tmpRegex;

            MatchCollection matches;
            if (reg == null)
            {
                tmpRegex = regexpHandler.RegionRegex;
                matches = tmpRegex.Matches(content);
                if (!matches.Cast<Match>().Any())
                {
                    tmpRegex = regexpHandler.RegionToLeftRegex;
                    matches = tmpRegex.Matches(content);
                }
            }
            else
            {
                tmpRegex = reg;
                matches = tmpRegex.Matches(content);
            }

            //Если есть совпадение
            if (!matches.Cast<Match>().Any()) return;


            var match = matches.Count > 1
                ? matches.Cast<Match>().FirstOrDefault(m => !Regex.IsMatch(m.Groups["type"].Value, "(^|\b)г")) ??
                  matches.Cast<Match>().First()
                : matches.Cast<Match>().First();


            var name = TryChangeRegionEndness(TemplateName(match.Groups["name"].Value));
            var type = match.Groups["type"].Value;
            if (type.IndexOf("г", StringComparison.OrdinalIgnoreCase) >= 0)
                type = "город";
            else if (match.Groups["type"].Value.IndexOf("р", StringComparison.OrdinalIgnoreCase) >= 0)
                type = "район";
            else
            {
                Console.WriteLine(
                    @"Неизвестный тип муниципального образования: {0}. \r\n Строка:{1} \r\n Контекст: {2}",
                    match.Groups["type"].Value, row, content);
                return;
            }

            //Пытаемся найти полное наименование во всём ОКТМО
            var fullName = supportWorksheets.OKTMOWs.GetFullName(name, OKTMOColumn.Region, type);
            if (fullName == "")
            {
                if (name.EndsWith("ий"))
                {
                    var sb = new StringBuilder(name);
                    sb[sb.Length - 2] = 'о';
                    name = sb.ToString();
                    fullName = supportWorksheets.OKTMOWs.GetFullName(name, OKTMOColumn.Region, type);
                }
            }

            //Spet 1: Подходит ли регион к субъекту
            if ((!string.IsNullOrEmpty(fullName) && oktmoComposition.SubjectHasEqualRegion(fullName.ToLower())) ||
                ChangeRegTypeAndCheck(ref fullName))
            {
                //Отлично, найденное мунОбр-е относится к текущему субъекту
                //нам надо использовать найденный текст?
                if (RegionCell.Value == "" ||
                    (!RegionCell.Value.EqualNoCase(fullName) &&
                     RegionCell.Valid == false))
                {
                    if (RegionCell.Value != "")
                        AppendToLandMarkCell(RegionCell.Value);

                    RegionCell.Value = fullName;

                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                    SubjectCell.SetStatus(DataCell.DataCellStatus.Valid);

                    //Выборка
                    if (oktmoComposition.HasEqualRegion(fullName.ToLower()))
                    {
                        var spec = new RegionSpecification(fullName);
                        oktmoComposition.SetSpecifications(spec);

                        RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }

                    //Окрашиваем Регион если он подходит к субъекту но не подходит к выборке
                    else
                    {
                        RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                    }
                }
                //Запись в ориентир если текущий рег центр верный. Найденный нам просто не нужен
                else if (RegionCell.Value != "" &&
                         !string.Equals(RegionCell.Value, fullName, StringComparison.OrdinalIgnoreCase))
                {
                    AppendToLandMarkCell(fullName);
                }
            }
            else
            {
                //step 2: Проверяем на промежуточную принадлежность
                //Город к насел пункту
                //Район к ВГТ
                if (type == "город")
                {
                }
                else
                {
                    if (!TryFillVGT(ref name))
                    {
                        var sb = new StringBuilder(name);
                        sb[sb.Length - 2] = 'и';
                        name = sb.ToString();

                        if (!TryFillVGT(ref name))
                        {
                            //Step 3: проверяем принадлежность региона к ОКТМО
                            //Проверялось при заполнеии fullName. Если заполнено, значи есть в ОКТМО
                            if (!string.IsNullOrEmpty(fullName))
                            {
                                //Если стоит верный, найденный кидаем в ориентир
                                if (RegionCell.Value != "" && !RegionCell.Valid)
                                {
                                    AppendToLandMarkCell(fullName);
                                }
                                //Тут мы пишем неверный в пустую ячейку, либо заменяем один неверный на другой
                                else
                                {
                                    if (RegionCell.Value != "")
                                        AppendToLandMarkCell(RegionCell.Value);
                                    RegionCell.Value = fullName;

                                    RegionCell.Valid = false;
                                    SubjectCell.Valid = false;
                                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                    SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                }
                            }
                            //Запись в оринетир если найденный регион не существует в ОКТМО
                            else
                                AppendToLandMarkCell(name + " " + type);
                        }
                    }
                }
            }
            content = match.Value == content.Trim() ? "" : content.Replace(match.Value, ", ").Trim().Trim(',').Trim();
        }


        /// <summary>
        ///     Метод добавляет или убирает текст " муниципальное образование" из названия муниципального образования и проверяет его наличие в ОКТМО по субъекту
        /// </summary>
        /// <param name="fullName"></param>
        /// <returns></returns>
        private bool ChangeRegTypeAndCheck([NotNull] ref string fullName)
        {
            var newVal = fullName;
            if (string.IsNullOrEmpty(newVal)) return false;
            const string mynType = " муниципальный район";

            if (newVal.ToLower().ToLower().Contains(mynType))
                newVal = newVal.ToLower().Replace(mynType, "");
            else
                newVal = newVal + mynType;

            var result = oktmoComposition.SubjectHasEqualRegion(newVal.ToLower());

            //change source value if test passed
            if (result == true) fullName = newVal;

            return result;
        }

        private void CheckSubejctCell()
        {
            var subjectName = supportWorksheets.SoubjectSourceWorksheet.GetSubjectBySourceLink(SourceLinkCell.InitValue);


            if (string.IsNullOrEmpty(subjectName))
            {
                var justSubjName = regexpHandler.TryCutSubjName(SubjectCell.InitValue);
                subjectName = supportWorksheets.OKTMOWs.GetFullName(justSubjName, OKTMOColumn.Subject);
            }

            //определили
            if (!string.IsNullOrEmpty(subjectName))
            {
                //отбираем
                var subjRows = supportWorksheets.OKTMOWs.GetSubjectRows(subjectName).ToList();
                if (!subjRows.Any()) return;

                //вставляем
                SubjectCell.Value = subjectName;

                oktmoComposition.SetSubjectRows(subjRows);
                oktmoComposition.ResetToSubject();

                SubjectCell.Valid = true;

                //Get RegCenter
                regCenter = supportWorksheets.OKTMOWs.GetDefaultRegCenterFullName(SubjectCell.Value, ref regName);
                if (string.IsNullOrEmpty(regCenter))
                    Console.WriteLine(@"Не найден региональный центр по субъекту {0}", subjectName);
            }
            else
            {
                //субъекит у нас не нашёлся
                SubjectCell.Cell.Value = SubjectCell.InitValue;
                SubjectCell.Valid = false;
                SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
            }
        }

        /// <summary>
        ///     Вставить текст в ячейку "Ориентир" по указанной строке
        /// </summary>
        /// <param name="value">Текст для вставки</param>
        private void AppendToLandMarkCell(string value)
        {
            if (LandMarkCell.Value != "" &&
                LandMarkCell.Value.IndexOf(value, StringComparison.Ordinal) >= 0)
                return;


            if (value.IndexOf("район", StringComparison.Ordinal) >= 0)
                if (LandMarkCell.Value == "")
                    LandMarkCell.Value = value + ", ";
                else
                    LandMarkCell.Value = value + ", " + LandMarkCell.Value;
            else
                LandMarkCell.Value += value + ", ";
        }

        /// <summary>
        ///     Иетод возвращает расшифрованную дистанцию
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static int TryDescriptDistance(string value)
        {
            const string integer = @"\d(\d|\.|\,)*";
            var match = Regex.Match(value, integer);
            if (!match.Success) return 0; //Хотя вообще такого случаться в принципе не должно

            int result;
            int.TryParse(match.Value, out result);

            if (value.IndexOf("к", StringComparison.OrdinalIgnoreCase) == -1)
                result = result*1000;

            return result;
        }

        /// <summary>
        ///     Метод пытается найти Имена собственные в переданной строке, и пытается их опеределить к какому-либо тиипу (мунОбр,
        ///     населПункт, ВГТ и прочие)
        /// </summary>
        /// <param name="value"></param>
        internal void TryFindProperName(ref string value)
        {
            if (oktmoComposition.CustomOktmoRows == null) return;
            if (!oktmoComposition.CustomOktmoRows.Any()) return;

            var matches = regexpHandler.WordWithHeadLetteRegex.Matches(value);

            for (var i = 0; i < matches.Count; i++)
            {
                var match = matches[i];
                var propName = TemplateName(match.Value);
                var wordRegex = new Regex("(\\b|^)" + propName + "(\\b|$)", RegexOptions.IgnoreCase);

                //имя собственное уже есть где-то
                if (wordRegex.IsMatch(RegionCell.Value) || wordRegex.IsMatch(NearCityCell.Value) ||
                    (SntKpsCell.Value != "" && wordRegex.IsMatch(NearCityCell.Value))) continue;

                //Пробуем подогнать к каждой ячейке
                if (TryAppendPropName(propName))
                {
                    value = value.Replace(propName, ", ").Trim().Trim(',').Trim();
                    continue;
                }

                if (i > 0)
                {
                    var multiProp = "";
                    var appended = false;
                    for (var j = i - 1; i >= 0; i--)
                    {
                        multiProp = matches[j].Value + " " + multiProp;
                        appended = (TryAppendPropName(multiProp));
                        if (appended) break;
                    }
                    if (appended) continue;
                }


                //Если никуда не подошло то пишем в первую пустую
                var replaceWord = TryAppendToFirstEmptyCell(propName);

                if (replaceWord)
                    value = value.Replace(propName, ", ").Trim().Trim(',').Trim();
            }
        }

        /// <summary>
        ///     Метод вытается вписать Имя собственное в первую свободную ячейку
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        private bool TryAppendToFirstEmptyCell(string propName)
        {
            bool replaceWord = true;

            if (NearCityCell.Value == "")
            {
                NearCityCell.Value = TemplateName(propName);

                NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);

                if (RegionCell.Value != "")
                {
                    RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                }
                else
                {
                    SubjectCell.SetStatus(DataCell.DataCellStatus.InValid);
                }
            }
            else
                replaceWord = false;

            return replaceWord;
        }

        private bool TryAppendPropName(string propName)
        {
            //Try append to Region

            if (TryAppendPropNameToRegion(propName)) return true;

            //Try append to NearCity
            if (TryAppendPropNameToNearCity(propName)) return true;

            //Try Append To VGT
            if (supportWorksheets.VgtWorksheet.TerritotyExists(propName))
            {
                if (TryFillVGT(ref propName))
                    return true;
            }

            //try append to Street via @Kladr
            if (TryAppendPropNameToStreet(propName))
            {
                return true;
            }

            //To street by end
            if (StreetCell.Value == "" &&
                Regex.IsMatch(propName, @"ая\b", RegexOptions.IgnoreCase))
            {
                StreetCell.Value = TemplateName(propName);
                TypeOfStreetCell.Value = "улица";
                return true;
            }

            return false;
        }

        private bool TryAppendPropNameToRegion(string propName)
        {
            var fullName = OKTMORepository.GetFullName(oktmoComposition.CustomOktmoRows, propName,
                OKTMOColumn.Region);
            if (!string.IsNullOrEmpty(fullName))
            {
                if (!cellsFilled)
                {
                    if (oktmoComposition.HasEqualRegion(fullName.ToLower()))
                    {
                        RegionCell.Value = fullName;

                        RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                        SubjectCell.SetStatus(DataCell.DataCellStatus.Valid);

                        //Делаем выборку только если найденный регион не является региональным центром
                        if (!string.Equals(fullName, regCenter, StringComparison.OrdinalIgnoreCase))
                        {
                            var spec = new RegionSpecification(fullName);
                            oktmoComposition.SetSpecifications(spec);
                        }
                    }
                }
                return true;
            }
            return false;
        }

        private bool TryAppendPropNameToNearCity(string propName)
        {
            if (oktmoComposition.HasEqualNearCity(propName.ToLower()))
            {
                if (!cellsFilled)
                {
                    NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);

                    NearCityCell.Value = propName;

                    if (!string.Equals(propName, regName, StringComparison.OrdinalIgnoreCase))
                    {
                        var spec = new NearCitySpecification(propName);
                        oktmoComposition.SetSpecifications(spec);
                    }
                }
                return true;
            }

            if (!oktmoComposition.SubjectHasEqualNearCity(propName.ToLower())) return false;

            if (NearCityCell.Value != "") return false;
            
            NearCityCell.Value = propName;
            NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
            RegionCell.SetStatus(DataCell.DataCellStatus.InValid);

            return true;
        }

        private bool TryAppendPropNameToStreet(string streetName)
        {
            var itIsStreet = supportWorksheets.Kladr.IsStreetFromKladr(streetName);
            if (!itIsStreet) return false;

            if (NearCityCell.Value == "")
            {
                if (!RegionCell.Valid) return false;
                if (!supportWorksheets.Kladr.IsStreetFromRegion(streetName, RegionCell.Value)) return false;

                StreetCell.Value = streetName;
                StreetCell.SetStatus(DataCell.DataCellStatus.Valid);
                var type = supportWorksheets.Kladr.GetStreetTypeFromRegion(streetName, RegionCell.Value);
                if (type != "")
                    TypeOfStreetCell.Value = type;

                return true;
            }

            if (NearCityCell.Valid)
            {
                if (!supportWorksheets.Kladr.IsStreetFromNearCity(streetName, NearCityCell.Value)) return false;

                StreetCell.Value = streetName;
                StreetCell.SetStatus(DataCell.DataCellStatus.Valid);
                var type = supportWorksheets.Kladr.GetStreetTypeFromCity(streetName, NearCityCell.Value);
                if (type != "")
                    TypeOfNearCityCell.Value = type;
            }

            return true;
        }

        private bool TryFillVGT(ref string value)
        {
            //----Обрабатываем ВГТ-----
            if (string.IsNullOrEmpty(value)) return false;

            //Подтверждаем, что это ВГТ
            if (!supportWorksheets.VgtWorksheet.TerritotyExists(value)) return false;

            var vgt = value;

            if (VgtCell.Value == "")
            {
                VgtCell.Value = vgt;
            }
            else if (RegionCell.Value != "" && RegionCell.Valid &&
                     supportWorksheets.VgtWorksheet.CombinationExists(RegionCell.Value, vgt))
            {
                VgtCell.Value = vgt;
            }

            return true;

            //а не записывается ли он потом, по найденному городу

            //Далее идут ситации если текущий насел пункт пустой, или не подходит к найденному ВГТ

            #region old logic

            //Пробуем определить населенный пункт
//            var city = string.Empty;
//
//
//            //Пробуем извлечь текущий насел пункт из мунОбр
//            //И тем самым подтвердить мунОбр и проставить населПункт
//            if (RegionCell.Value != "" &&
//                RegionCell.Value.IndexOf("город", StringComparison.OrdinalIgnoreCase) >= 0)
//            {
//                city = TryTemplateName(RegionCell.Value.Replace("город", ""));
//                city = city.Trim();
//            }
//            if (!string.IsNullOrEmpty(city) && supportWorksheets.VgtWorksheet.CombinationExists(city, vgt))
//            {
//                NearCityCell.Value = city;
//                TypeOfNearCityCell.Value = "город";
//
//                //Проверяем найденный насел пункт
//                if (oktmoComposition.HasEqualNearCity(city))
//                {
//                    NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
//                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
//
//                    var spec = new NearCitySpecification(city);
//                    oktmoComposition.SetSpecifications(spec);
//                }
//                else
//                {
//                    NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
//                    RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
//                }
//            }
//
//            //В ином случае пробуем записать насел пункт через ВГТ
//            else
//            {
//                var newCity = cellsFilled
//                    ? string.Empty
//                    : supportWorksheets.VgtWorksheet.GetCityByTerritory(vgt);
//
//                //Строка будет  заполнена, если существует всего один насел пункт с таким районом
//                if (string.IsNullOrEmpty(newCity)) return true;
//
//                //нужно ли нам вообще проверять найденный
//                if (NearCityCell.Value != "" &&
//                    string.Equals(NearCityCell.Value, newCity,
//                        StringComparison.CurrentCultureIgnoreCase)) return true;
//
//                //Если текущий населенный пункт верный (он не пуст и не окрашен как неверный)
//                //мы его оставляем на месте, а найденный пишем в ориентир
//                if (NearCityCell.Value != "" && NearCityCell.Valid)
//                    //Пишем найденный насел пункт в ориентир
//                    LandMarkCell.Value += "город " + vgt + ", ";
//
//                //В остальных случаях найденный насел пункт попадёт в ячейку населенногоп пункта
//                else
//                {
//                    //Определяем, относится ли насел пункт к выборке
//                    if (oktmoComposition.HasEqualCityType(newCity))
//                    {
//                        NearCityCell.Valid = true;
//                        RegionCell.Valid = true;
//                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//
//
//                        //Try to fill
//                        var spec = new NearCitySpecification(newCity);
//                        oktmoComposition.SetSpecifications(spec);
//                    }
//                    else
//                    {
//                        NearCityCell.Valid = false;
//                        RegionCell.Valid = false;
//                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
//                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
//                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
//                    }
//
//                    //Перекидываем текущий насел пункт
//                    if (NearCityCell.Value != "")
//                        LandMarkCell.Value += NearCityCell + ", ";
//
//                    NearCityCell.Value = newCity;
//                }
//            }
//            return true;

            #endregion
        }

        private void TryFillStreet(ref string value)
        {
            //Поиск улиц
            var regs = new List<Regex> {regexpHandler.StreetToLeftRegex, regexpHandler.StreetRegex};
            foreach (var reg in regs)
            {
                var match = reg.Match(value);
                if (!match.Success) continue;
                //По сути если у нас уже проставлена улица, новую нужно игнорировать кроме нескольких случаев ниже

                //Берём только первое совпадение!
                var name = ReplaceYo(TemplateName(match.Groups["name"].Value));

                var type = "";
                if (NearCityCell.Valid && NearCityCell.Value != "")
                    type = supportWorksheets.Kladr.GetStreetTypeFromCity(name, NearCityCell.Value);

                if (type == "")
                    type = supportWorksheets.Kladr.GetStreetType(name);

                if (type == "")
                    if (supportWorksheets.Kladr.TypeFromBase(match.Groups["type"].Value))
                        type = match.Groups["type"].Value;

                if (type == "")
                    type = ReplaceYo(TryDescriptTypeOfStreet(match.Groups["type"].Value));

                if (StreetCell.Value == "" || StreetCell.Value == string.Empty ||
                    StreetCell.Value != name &&
                    (TypeOfStreetCell.Value == "" ||
                     TypeOfStreetCell.Value.ToLower() == "микрорайон".ToLower()))
                {
                    //Backups current INFO
                    //Когда стоит микрорайон, а найдена улица, приориет у улицы
                    if (TypeOfStreetCell.Value == "микрорайон" &&
                        type != "микрорайон")
                        LandMarkCell.Value += StreetCell.Value + " " + TypeOfStreetCell.Value + ", ";
                    //Когда стоит Именование, без типа
                    else if (TypeOfStreetCell.Value == "" && StreetCell.Value != "")
                        LandMarkCell.Value += StreetCell.Value + ", ";

                    StreetCell.Value = name;
                    TypeOfStreetCell.Value = type;
                }

                //Отдельная логика для информации о доме
                if (!string.IsNullOrEmpty(match.Groups["house_num"].Value))
                {
                    if (BuildsCell.Value == "")
                        BuildsCell.Value = match.Groups["house_num"].Value.Trim().Trim(',').Trim();
                }

                value = reg.Replace(value, ", ").Trim().Trim(',').Trim();
            }
        }


        /// <summary>
        ///     Метод на основе сложившейся выборки пытается заполнить 100%-тные поля
        /// </summary>
        private void TryFillClassificator()
        {
            if (oktmoComposition.CustomOktmoRows == null) return;
            if (!oktmoComposition.CustomOktmoRows.Any()) return;
            if (cellsFilled) return;
            oktmoComposition.FixDoubles();

            //
            //Записываем город если он у нас один 
            //
            var cities =
                oktmoComposition.CustomOktmoRows.Select(r => r.City.Name)
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToArray();
            if (cities.Count() == 1)
            {
                var validCity = cities.First();

                if (NearCityCell.Value == "")
                {
                    NearCityCell.Value = validCity;
                    NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                }
                else if (NearCityCell.Value.EqualNoCase(validCity))
                {
                    NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                }
                else
                {
                    NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                }
            }
            //Проверяем текущее значение на принадлежность к выборке
            else
            {
                if ((NearCityCell.Value != ""))
                {
                    if (cities.All(s => s != NearCityCell.Value.ToString()))
                    {
                        if (RegionCell.Value != "")
                        {
                            NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                            RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                        }
                    }
                    else
                    {
                        NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                        RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                }
            }


            if (NearCityCell.Value != "" &&
                NearCityCell.Valid)
            {
                var types =
                    oktmoComposition.CustomOktmoRows.Select(r => r.City.Type)
                        .Where(s => !string.IsNullOrEmpty(s))
                        .Distinct()
                        .ToArray();

                if (types.Count() == 1)
                {
                    var validType = types.First();
                    if (!TypeOfNearCityCell.Value.EqualNoCase(validType))
                    {
                        if (TypeOfNearCityCell.Value != "")
                            AppendToLandMarkCell(NearCityCell.Value + " " + TypeOfNearCityCell.Value);

                        TypeOfNearCityCell.Value = validType;
                        TypeOfNearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                }

                else if (TypeOfNearCityCell.Value != "" &&
                         TypeOfNearCityCell.Valid)
                {
                    if (TypeCanClarifyRows(types))
                    {
                        var spec =
                            new ExpressionSpecification<OktmoRowDTO>(
                                r => (r.City.Type ?? "").EqualNoCase(TypeOfNearCityCell.Value));

                        oktmoComposition.SetSpecifications(spec);
                        oktmoComposition.FixDoubles();

                        TypeOfNearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                    else
                    {
                        TypeOfNearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                    }
                }
                else
                {
                    if (types.Any(t => t.EqualNoCase(TypeOfNearCityCell.Value)))
                    {
                        TypeOfNearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                }
            }

            //записываем поселение
            var settlements = oktmoComposition.CustomOktmoRows.Select(r => r.Settlement.Name).Distinct().ToArray();

            if (settlements.Count() == 1)
            {
                var validSettlement = settlements.First();

                if (SettlementCell.Value == "" ||
                    !SettlementCell.Value.EqualNoCase(validSettlement))
                {
                    if (SettlementCell.Value != "")
                        AppendToLandMarkCell(SettlementCell.Value);

                    SettlementCell.Value = validSettlement;

                    SettlementCell.SetStatus(DataCell.DataCellStatus.Valid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                }
                else if (SettlementCell.Value != "")
                {
                    SettlementCell.SetStatus(DataCell.DataCellStatus.Valid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                }
            }
            else
            //Проверяем текущее значение на принадлежность к выборке
            {
                if (SettlementCell.Value != "")
                    if (settlements.All(s => s != SettlementCell.Value.ToString()))
                    {
                        if (RegionCell.Value != "")
                        {
                            SettlementCell.SetStatus(DataCell.DataCellStatus.InValid);
                            RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                        }
                    }
                    else
                    {
                        SettlementCell.SetStatus(DataCell.DataCellStatus.Valid);
                        RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                else
                {
                    if (settlements.Any(s => s == "") && !NearCityCell.Valid && RegionCell.Valid &&
                        NearCityCell.Value != "")
                    {
                        SettlementCell.SetStatus(DataCell.DataCellStatus.Valid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                }
            }


            //
            //Записываем регион (муниципальное образование)
            //
            var regions =
                oktmoComposition.CustomOktmoRows.Select(r => r.Region)
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToArray();
            if (regions.Count() == 1)
            {
                var validRegion = regions.First();

                if (RegionCell.Value != "" &&
                    string.Equals(RegionCell.Value, validRegion, StringComparison.OrdinalIgnoreCase))
                {
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                    SubjectCell.SetStatus(DataCell.DataCellStatus.Valid);
                }
                else
                {
                    if (RegionCell.Value != "")
                        LandMarkCell.Value += RegionCell.Value + ", ";
                    RegionCell.Value = validRegion;

                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                    SubjectCell.SetStatus(DataCell.DataCellStatus.Valid);
                }
            }
            //Проверяем текущее значение на принадлежность к выборке
            else
            {
                if (RegionCell.Value != "")
                    if (regions.All(s => s != RegionCell.Value.ToString()))
                    {
                        if (SubjectCell.Value != "")
                        {
                            RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                            SubjectCell.SetStatus(DataCell.DataCellStatus.InValid);
                        }
                    }
                    else
                    {
                        RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
                        SubjectCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
            }

            if (oktmoComposition.CustomOktmoRows.Count == 1)
                cellsFilled = true;
        }

        private bool TypeCanClarifyRows(string[] types)
        {
            return types.Any(
                s => TypeOfNearCityCell.Value.EqualNoCase(s)) &&
                   NearCityCell.Value != "" &&
                   oktmoComposition.CustomOktmoRows.Any(
                       r => (r.City.Name ?? "").EqualNoCase(NearCityCell.Value) &&
                            (r.City.Type ?? "").EqualNoCase(NearCityCell.Value));
        }

        private static string TryChangeSubjectEndness(string s)
        {
            var reg = new Regex(@"(ой)\b");

            var match = reg.Match(s);
            if (!match.Success) return s;

            var newString = reg.Replace(s, "ая");

            return newString;
        }

        public static string ReplaceYo(string s)
        {
            var s2 = s.Replace("ё", "е");
            var s3 = s2.Replace("Ё", "Е");
            return s3;
        }

        private static string TryDescriptTypeOfNasPunkt(string s)
        {
            if (s.ToLower() == "дп")
                s = "дачный поселок";
            else if (Regex.IsMatch(s, @"\bд(ер(евн[а-я]*)?)?\.?", RegexOptions.IgnoreCase))
                s = "деревня";
            else if (Regex.IsMatch(s, @"\bг(ород[а-я]*|\.|\b)?", RegexOptions.IgnoreCase))
                s = "город";
            else if (Regex.IsMatch(s, @"\bм\b", RegexOptions.IgnoreCase))
                s = "местечко";
            else if (Regex.IsMatch(s, @"дачн\w+\sп(ос((е|ё)л(о?к[а-я]{0,3})?)?)?\.?", RegexOptions.IgnoreCase))
                s = "дачный поселок";
            else if (Regex.IsMatch(s, @"\bр.?п\.?", RegexOptions.IgnoreCase))
                s = "рабочий поселок";
            else if (Regex.IsMatch(s, @"\b(с|c)(ел[а-я]*)?\.?", RegexOptions.IgnoreCase))
                s = "село";
            else if (Regex.IsMatch(s, @"\bх\.?", RegexOptions.IgnoreCase))
                s = "хутор";
            else if (Regex.IsMatch(s, @"\bпгт\.?", RegexOptions.IgnoreCase))
                s = "поселок городского типа";
            else if (Regex.IsMatch(s, @"п(ос((е|ё)л(о?к[а-я]{0,3})?)?)?\.?", RegexOptions.IgnoreCase))
                s = "поселок";
            else if (Regex.IsMatch(s, @"\bнп", RegexOptions.IgnoreCase))
                s = "поселок";

            return s;
        }

        private static string TryDescriptTypeOfStreet(string s)
        {
            if (Regex.IsMatch(s, @"\bм", RegexOptions.IgnoreCase))
                s = "микрорайон";
            else if (Regex.IsMatch(s, @"\bтр", RegexOptions.IgnoreCase))
                s = "тракт";
            else if (Regex.IsMatch(s, @"\bпер", RegexOptions.IgnoreCase))
                s = "переулок";
            else if (Regex.IsMatch(s, @"\bб", RegexOptions.IgnoreCase))
                s = "бульвар";
            else if (Regex.IsMatch(s, @"\bпр.*д\b", RegexOptions.IgnoreCase))
                s = "проезд";
            else if (Regex.IsMatch(s, @"\bпрос.*(л|к\b)", RegexOptions.IgnoreCase))
                s = "проселок";
            else if (Regex.IsMatch(s, @"\bпр.*т\b", RegexOptions.IgnoreCase))
                s = "проспект";
            else if (Regex.IsMatch(s, @"\bш", RegexOptions.IgnoreCase))
                s = "шоссе";
            else if (Regex.IsMatch(s, @"\bт", RegexOptions.IgnoreCase))
                s = "тупик";
            else if (Regex.IsMatch(s, @"\bп", RegexOptions.IgnoreCase))
                s = "переулок";
            else if (Regex.IsMatch(s, @"\bул", RegexOptions.IgnoreCase))
                s = "улица";

            return s;
        }

        /// <summary>
        ///     Метод возвращает переданную строку в формате Первая буква заглавная, остальные прописные
        /// </summary>
        /// <param name="s">Строка для приведния к формату имени собственного</param>
        /// <returns>Отформатированная строка</returns>
        private static string TemplateName(string s)
        {
            const string justWordPattern = @"\b([А-Яа-я])([А-Яа-я]+)\b";
            const string perfectWordPatterd = @"\b[А-Я][а-я]+\b";
            s = s.Trim();
            if (s.Length < 6 && Regex.IsMatch(s, "^[А-Я]+$")) return s; //Для АББРИВЕАТУР

            var words = Regex.Matches(s, justWordPattern);
            //Если все слова уже приведены в порядок
            if (words.Cast<Match>().All(m => Regex.IsMatch(m.Value, perfectWordPatterd)))
                return s;

            var result = s;
            foreach (Match match in words)
            {
//                var newWord = Regex.Replace(match.Value, justWordPattern,
//                    m => string.Format("{0}{1}", m.Groups[1].Value.ToUpper(), m.Groups[2].Value.ToLower()));
                var newWord = char.ToUpper(match.Value[0]) + match.Value.ToLower().Substring(1);
                result = result.Replace(match.Value, newWord);
            }
            return result;
        }

        private static string TryChangeRegionEndness(string s)
        {
            var reg = new Regex(@"(ого|ом|ем)\b");

            var match = reg.Match(s);
            if (!match.Success) return s;

            var newString = reg.Replace(s, "ий");

            return newString;
        }

        [Obsolete("Не доделано", true)]
        private void TryFillCommunications(long row, ref string val)
        {
            //Инфу, что мы заменяем, не удаляем а переносим с новый столбец "Блабла"


            //1. "Свет, вода, канализация, дорога"
            //2. "Проведён свет, вода"
            //3. "Свет, вода есть"
            //4. "Свет, канализация есть, вода в перспективе"
            //5. "Вода есть, свет легко провести

            // ком + статус; статус + ком
            // ком, ком + статус, статус + ком


            //=======================
            //1.Arrange
            //=======================
            //Формат "проведён свет, вода, канализация"
//            var sentenceReg = new Regex(@"(?n)(?<=(^|\b))(?!\.)[^\!\?$^]{5,}?(?=((?<!\s\w{1,2})\.|\!|\?|$))",
//                RegexOptions.Multiline);
//            var sentencesMatchCollection = sentenceReg.Matches(val);
//
//
//            //Общие регулярки
//            const string wordUnions = @"\s*(и|\,|;|:)\s*";
//            //Пунктуационанные знаки в предложении, объединяющие части предложения
//            const string sentenceEnds = @"\s*((<!\\s\w{1,4})\.|!|?)";
//            //Пунктационные знаки, обозначающие конец предложения
//
//            const string justWords = @"\(w(\w)*|\s)+"; //Паттерн для выялвения просто слов

            const string startCollocation = @"(?<=^|\""|(?<!\s\w{1,2})\.|\,|\)|\()";
            //Символы, обозначающие начало предложения
            const string endCollocation = @"(?=$|\""|(?<!\s\w{1,3})\.(\s|$|\,)|\,|\)|\()";
            //Символы, обозначающие конец предложения

            const string orDel = @"|"; //Символ Или
            const string spacesNRq = @"\s*"; //Наличие пробела в кол-ве от нуля до бесконечности

            //==========
            //Статус
            //Перечень фраз для подтверждения наличия коммуникации
            const string comValid =
                @"(?<valid>круглый\sгод|всегда|подведен(о|ы)|централизирован(а|о)|(?!в\sобществе\s)проводят|провед(ё|е)но?(?!\sк\sгранице)|на\s(участке|территории)|есть(?!\s*возможность)|име(е|ю)тся|(?<kvt>\d(\d|\.|\,)*)\s*квт)";
//            const string comCanConnectAlwaysLeft = @"";
//            const string comCanConnectAlwaysRight = @"";
            //Перечень фраз для подтверждения возможности провести коммуникацию
            const string comCanConnect =
                @"(?<canconn>в\sперспективах|\bТУ\b|проводится|будет|проведут|в\sобществе\sпроводят|легко\sпровести|оплачивается\sотдельно|(проведен\s(к|по))?границе|подключение(\sту)?|рядом\sпроходит|(есть\s)?возможно(сть)?|в\s\d+\sм(\.|етрах|\s)|актуально(\sпровести)?|разешени(е|я)|около|техусловия|соласовано|(на|по)\sулице|не\sдалеко)";
            const string comNo = @"(?<no>нет|отсутству(е|ю)т)"; //Фразы, подтверждающие отсутствие коммуникации
//            const string comTemp = @"(?<temp>летний|зимний)"; //Наличие сезонной коммуникации


//            const string delimCom = @"(\s*(\,|\.)\s*)"; //Символы разделители между преречисленными коммуникациями

            const string commonCommunicatuionNames = @"\b(?<all>коммуникации|удобства)\b";
            //Паттерн с перечнем всех коммуникация
            //==========
            //ЭлестроЭнергия
            const string electrNames = @"(?<elec>свет|эл\.свет|эл(-|лектрич(ест|\-|еск(ие|ая)\sсет(ь|и)))во)";

            //==========
            //Водопровод
            const string waterNames = @"\b(?<water>вода|водопровод|скважина|колодец|родник)\b";

            //==========
            //Газопровод
            const string gasNames = @"\b(?<gas>газ)\b";

            //==========
            //Канализация
            const string severageNames = @"(?<swrg>канализиция)";


            //======================
            //======================
            //В другой метод
            //==========
            //Рельеф
//            const string reliefNames = @"(?<relief>ровный)";

            //==========
            //Дорога
//            const string roadNames = @"(?<road>асфальт|грунтовая|засыпана)";
            //======================
            //======================


            //Строка-паттерн-перечень всех вохможных коммуникаций для выявления одного
            const string anyCom =
                "(?<anyCom>" + commonCommunicatuionNames + orDel + electrNames + orDel + waterNames + orDel + gasNames +
                orDel + severageNames + ")";

            //Строка-паттерн-перечень всех возможных коммуникаций для выявления их в прямой последновательности
//            const string stringOfAnyCom = "(" + delimCom + anyCom + @"|\s*\,\s*";


            //=======================
            //2.ACT
            //=======================
            //Список вариантов=паттернов, вощвращающих паттерн для регулярного выражения
            var patterns = new List<string>
            {
                // ",свет,вода,ещё что"
                startCollocation + spacesNRq + anyCom + spacesNRq + endCollocation,


                //Bug "Доступен свет, канализация, вода проведётся не скоро"
                //Уточнения слева в тексте от коммуникаций
                //                  "Доступен               свет, вода"
                startCollocation + comValid + spacesNRq + anyCom + spacesNRq + endCollocation,

                //              "Недоступен"
                startCollocation + comNo + spacesNRq + anyCom + spacesNRq + endCollocation,

                //              "Возможно проведение"
                startCollocation + comCanConnect + spacesNRq + anyCom + spacesNRq + endCollocation,


                //Уточнения справа в тексте от коммуникация
                //                  "свет, вода          Доступен"
                startCollocation + anyCom + spacesNRq + comValid + spacesNRq + endCollocation,

                //                  "свет, вода          Недоступен"
                startCollocation + anyCom + spacesNRq + comNo + spacesNRq + endCollocation,

                //                  "свет, вода          Возможно проведение"
                startCollocation + anyCom + spacesNRq + comCanConnect + spacesNRq + endCollocation
            };

            //Цикл для обработки всех вариантов, представленных выше
            foreach (var reg in patterns.Select(funcS => new Regex(funcS)))
            {
            }
        }

        private bool disposed = false;

        public void Dispose()
        {
            if (disposed) return;
            SubjectCell.Dispose();
            RegionCell.Dispose();
            SettlementCell.Dispose();
            NearCityCell.Dispose();
            TypeOfNearCityCell.Dispose();
            StreetCell.Dispose();
            BuildsCell.Dispose();
            DictToNearCityCell.Dispose();
            DistToRegCenterCell.Dispose();
            HouseNumCell.Dispose();
            InCityCell.Dispose();
            LandMarkCell.Dispose();
            LetterCell.Dispose();
            SntKpsCell.Dispose();
            SourceLinkCell.Dispose();
            VgtCell.Dispose();
        }

        public void SplitConcatenatedCellS()
        {
            SubjectCell.InitValue = TrySpilitConcatenatedWords(SubjectCell.InitValue);
            RegionCell.InitValue = TrySpilitConcatenatedWords(RegionCell.InitValue);
            SettlementCell.InitValue = TrySpilitConcatenatedWords(SettlementCell.InitValue);
            NearCityCell.InitValue = TrySpilitConcatenatedWords(NearCityCell.InitValue);
            StreetCell.InitValue = TrySpilitConcatenatedWords(StreetCell.InitValue);
        }

        private string TrySpilitConcatenatedWords(string text)
        {
            var matches = regexpHandler.ConcatenatedWordsRegex.Matches(text);
            if (matches.Count == 0) return text;
            var newText = text;
            foreach (var match in matches.Cast<Match>())
            {
                var name1 = match.Groups["name1"].Value;
                newText = newText.Replace(name1, name1 + ", ");
            }
            return newText;
        }

        public string actualType { get; set; }
    }

    public class SupportWorksheets : IDisposable
    {
        public SupportWorksheets(CatalogWorksheet catalogWorksheet, OKTMORepository oktmo,
            SubjectSourceWorksheet subjectSourceWorksheet, VGTWorksheet vgtWorksheet, KladrRepository kladr)
        {
            CatalogWs = catalogWorksheet;
            OKTMOWs = oktmo;
            SoubjectSourceWorksheet = subjectSourceWorksheet;
            VgtWorksheet = vgtWorksheet;
            Kladr = kladr;
        }

        public CatalogWorksheet CatalogWs { get; private set; }
        public OKTMORepository OKTMOWs { get; private set; }
        public SubjectSourceWorksheet SoubjectSourceWorksheet { get; private set; }
        public VGTWorksheet VgtWorksheet { get; private set; }
        public KladrRepository Kladr { get; private set; }

        private bool disposed = false;

        public void Dispose()
        {
            if (disposed) return;

            if (CatalogWs != null) CatalogWs.Dispose();
            if (OKTMOWs != null) OKTMOWs.Dispose();
            if (SoubjectSourceWorksheet != null) SoubjectSourceWorksheet.Dispose();
            if (VgtWorksheet != null) VgtWorksheet.Dispose();
            if (Kladr != null) Kladr.Dispose();
            disposed = true;
        }
    }

    public class DataCell : IDisposable
    {
        public enum DataCellStatus
        {
            Valid,
            InValid
        }

        public DataCell(ExcelRange cell)
        {
            Cell = cell;
            InitValue = (Cell.Value ?? "").ToString().Replace("ё", "е").Replace("Ё", "Е");
            Value = "";
        }

        public ExcelRange Cell { get; private set; }
        public string InitValue { get; set; }
        public string Value { get; set; }

        public bool Valid { get; set; }

        public void SetDefaultValue()
        {
            Value = InitValue;
        }

        public void SetStatus(DataCellStatus status)
        {
            switch (status)
            {
                case DataCellStatus.InValid:
                    Valid = false;
                    Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    break;
                case DataCellStatus.Valid:
                    Valid = true;
                    Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                    break;
            }
        }

        public void Save()
        {
            Cell.Value = Value;
        }

        public void Dispose()
        {
            Cell.Dispose();
        }
    }
}