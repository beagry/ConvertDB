using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using ExcelRLibrary.SupportEntities.Oktmo;
using Formater.SupportWorksheetsClasses;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PatternsLib;

namespace Formater
{
    public class ExcelLocationRow:IDisposable
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
                oktmoHelper = new OktmoHelper();
                var columns = db.TemplateWorkbooks.First(w => w.WorkbookType == wbType).Columns.ToList();

                var subjColumn = (byte)columns.First(c => c.CodeName.Equals("SUBJECT")).ColumnIndex;
                var regionColumn = (byte)columns.First(c => c.CodeName.Equals("REGION")).ColumnIndex;
                var settlementColumn = (byte)columns.First(c => c.CodeName.Equals("SETTLEMENT")).ColumnIndex;
                var nearCityColumn = (byte)columns.First(c => c.CodeName.Equals("NEAR_CITY")).ColumnIndex;
                typeOfNearCityColumn = (byte)columns.First(c => c.CodeName.Equals("TERRITORY_TYPE")).ColumnIndex;
                var vgtColumn = (byte)columns.First(c => c.CodeName.Equals("VGT")).ColumnIndex;
                var streetColumn = (byte)columns.First(c => c.CodeName.Equals("STREET")).ColumnIndex;
                var typeOfStreetColumn = (byte)columns.First(c => c.CodeName.Equals("STREET_TYPE")).ColumnIndex;
                var sourceLinkColumn = (byte)columns.First(c => c.CodeName.Equals("URL_SALE")).ColumnIndex;
                var distToRegCenterColumn = (byte)columns.First(c => c.CodeName.Equals("DIST_REG_CENTER")).ColumnIndex;
                var distToNearCityColumn = (byte)columns.First(c => c.CodeName.Equals("DIST_NEAR_CITY")).ColumnIndex;
                var inCityColumn = (byte)columns.First(c => c.CodeName.Equals("IN_CITY")).ColumnIndex;
                var houseNumColumn = (byte)columns.First(c => c.CodeName.Equals("HOUSE_NUM")).ColumnIndex;
                var letterColumn = (byte)columns.First(c => c.CodeName.Equals("LETTER")).ColumnIndex;
                var sntKpDnpColumn = (byte)columns.First(c => c.CodeName.Equals("ASSOCIATIONS")).ColumnIndex;
                var additionalInfoColumn = (byte)columns.First(c => c.CodeName.Equals("ADDITIONAL")).ColumnIndex;
                var buildColumn = (byte)columns.First(c => c.CodeName.Equals("HOUSE_NUM")).ColumnIndex;
                descriptionColumn = (byte)columns.First(c => c.CodeName.Equals("DESCRIPTION")).ColumnIndex;


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
            oktmoHelper = new OktmoHelper();
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
        private readonly OktmoHelper oktmoHelper;

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

            CheckLandmarkdsCell();
            if (breakFromRow) return;

            CheckDescriptionCell();
            if (breakFromRow) return;
            FillDefaultValues();
            if (!cellsFilled)
                FillOldvalues();

            SaveCells();
        }

        private void FillOldvalues()
        {
            var cells = new[]
            {
                SubjectCell, RegionCell,SettlementCell, NearCityCell
            };

            cells.AsParallel().ForAll(cell =>
            {
                if (cell.Valid) return;
                if (cell.InitValue == "") return;
                cell.SetDefaultValue();
                cell.SetStatus(DataCell.DataCellStatus.InValid);
            });


        }

        private void FillDefaultValues()
        {
            //Вписываем дефолтные значения Если населенный пункт так и не заполнен
            if (NearCityCell.Value == "")
            {
                //Находим дефолтный населенный пункт по ссылке на объявление
                var newCity = supportWorksheets.SoubjectSourceWorksheet.GetDefaultNearCityByLink(SourceLinkCell.InitValue);

                if (!string.IsNullOrEmpty(newCity))
                {
                    //Мы пишем насел пункт только если он подходит к нашей выборке
                    //Т.е. подходит и к субъекту и к муниципальному образованию, есть таковой есть
                    if (oktmoHelper.HasEqualNearCity(newCity))
                    {
                        NearCityCell.Value = newCity;
                        TypeOfNearCityCell.Value = "город";

                        var spec = new NearCitySpecification(newCity);
                        oktmoHelper.SetSpecifications(spec);

                        TryFillClassificator();
                    }
                }
                //или ставим муниципальное образование как город
                //При условии что это не региональный центр
                else if (RegionCell.Value != "" &&
                         RegionCell.Valid
                         && RegionCell.Value.Contains("город",StringComparison.OrdinalIgnoreCase))
                {
                    var name = RegionCell.Value.Replace("город", "");
                    name = name.Replace("(ЗАТО)", "");
                    name = name.Trim();
                    if (oktmoHelper.HasEqualNearCity(name))
                    {
                        cellsFilled = false;
                        NearCityCell.Value = name;
                        TypeOfNearCityCell.Value = "город";

                        var spec = new NearCitySpecification(name);
                        oktmoHelper.SetSpecifications(spec);
                        TryFillClassificator();
                    }
                }
            }
            //Ставим дефолтное значение для муниципального образования, если оно пустое, а текущий насленный пункт у нас является региональным центро
            else if (RegionCell.Value == "" &&
                     string.Equals(NearCityCell.Value, regName, StringComparison.OrdinalIgnoreCase))
            {
                var spec = new NearCitySpecification(regName);
                oktmoHelper.SetSpecifications(spec);

                TryFillClassificator();
            }
            //Дефолное значение для типа населенного пункта, если найденный насел пункт совпадает по названию с региональным центром
            else if (TypeOfNearCityCell.Value == "" &&
                     string.Equals(NearCityCell.Value, regName, StringComparison.OrdinalIgnoreCase))
            {
                TypeOfNearCityCell.Value = "город";
            }
        }

        private void CheckDescriptionCell()
        {
            if (!DoDescription) return;
            //Вначале мы ищем наименования по типу
            //После мы пытаемся отнести найдненные в описании Именования без типов
            var cell = worksheet.Cells[row, descriptionColumn];
            if ((string) cell.Value == "") return;

            var descrtContent = ReplaceYo((cell.Value??"").ToString()).Trim().Trim(',').Trim();
            

            //
            //----Товарищества
            //

            var match = regexpHandler.SntToLeftRegex.Match(descrtContent);
            while (match.Success)
            {
                //Берём только первое совпадение!
                var name = TryTemplateName(match.Groups["name"].Value);

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
                var fullName = supportWorksheets.OKTMOWs.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumn.Subject);

                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value != "" &&
                    !string.Equals(SubjectCell.Value.Trim(), fullName.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    rowsToDelete.Add(row);
                    SubjectCell.Value = fullName;

                    oktmoHelper.SetSubjectRows(supportWorksheets.OKTMOWs.GetSubjectRows(fullName).ToList());
                    oktmoHelper.ResetToSubject();
                    

                    breakFromRow = true;
                    return;
                }
            }

            //
            //----Населенный пункт
            //
            var switched = false;
            var endChanged = false;
            var regs = new List<Regex> { regexpHandler.NearCityToLeftRegex, regexpHandler.NearCityRegex };
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


                var name = ReplaceYo(TryTemplateName(match.Groups["name"].Value));
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

                    var valueNeedsResetRegion = !oktmoHelper.HasEqualNearCity(name) && oktmoHelper.SubjectHasEqualNearCity(name);

                    if (itIsCity && valueNeedsResetRegion)
                    {
                        oktmoHelper.ResetToSubject();

                        RegionCell.Value = string.Empty;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        SettlementCell.Value = string.Empty;
                        SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                    }

                    //найденный насел пункт подхоидт к нашей выборке (по субъекту и возможно по мунобразованию если оно есть)
                    if (oktmoHelper.HasEqualNearCity(name.ToLower()))
                    {
                        var spec = new NearCitySpecification(name);
                        oktmoHelper.SetSpecifications(spec);

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
                        if ((CurrentNearCityIsCity || NearCityCell.Value =="") && !type.EqualNoCase("город"))
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
                //TODO недоделано
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
            if (SubjectCell.InitValue != "")
            value = value.Replace(SubjectCell.InitValue, ", ").Trim().Trim(',').Trim();

            if (string.IsNullOrEmpty(value)) return;

            if (NearCityCell.Valid && NearCityCell.Value != "")
                value = value.Replace(NearCityCell.Value,"");

            if (string.IsNullOrEmpty(value)) return;

            if (regexpHandler.SignleLetterPerStringRegex.IsMatch(value.Trim()))
            {
                TryFindProperName(ref value);
                if (value.Length < 3) return;
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
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ").Trim().Trim(',').Trim();
                var fullName = supportWorksheets.OKTMOWs.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumn.Subject);

                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value != "" &&
                    SubjectCell.Value
                        .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                {
                    rowsToDelete.Add(row);
                    SubjectCell.Value = fullName;

                    oktmoHelper.SetSubjectRows(supportWorksheets.OKTMOWs.GetSubjectRows(fullName).ToList());
                    oktmoHelper.ResetToSubject();

                    breakFromRow = false;
                    return;
                }
            }

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
                var name = TryTemplateName(match.Groups["name"].Value);
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
                    if (oktmoHelper.HasEqualRegion(fullName))
                    {
                        RegionCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;

                        //Выборка
                        var spec = new RegionSpecification(fullName);
                        oktmoHelper.SetSpecifications(spec);
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
                else if (!TryFillVGT( ref name))
                {
                    fullName = name + " " + type;

                    //В зависимости заполнен ли уже Регион, пишем извлеченное значение в ячейку Региона или ДопИнформации
                    if (RegionCell.Value == "")
                        RegionCell.Value = fullName;
                    else if (fullName != RegionCell.Value)
                        LandMarkCell.Value = fullName + LandMarkCell.Value + ", ";
                }
                value = tmpRegex.Replace(value, ", ").Trim().Trim(',').Trim();
                if (value.Length <= 2) return;
            }

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
            if (match.Success)
            {
                var name = TryTemplateName(match.Groups["name"].Value);
                var type = match.Groups["type"].Value;

                type = type.IndexOf("п", StringComparison.OrdinalIgnoreCase) >= 0 ? "сельское поселение" : "сельсовет";

                var fullName = name + " " + type;
                if (SettlementCell.Value == "")
                    SettlementCell.Value = fullName;
                else
                    LandMarkCell.Value += fullName + ", ";

                if (oktmoHelper.HasEqualSettlement(fullName.ToLower()))
                {
                    var spec = new SettlementSpecification(fullName);
                    oktmoHelper.SetSpecifications(spec);
                }
                else
                {
                    SettlementCell.Valid = false;
                    SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    if (RegionCell.Value != "")
                    {
                        RegionCell.Valid = false;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else if (NearCityCell.Value != "")
                    {
                        NearCityCell.Valid = false;
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                }
                value = regexpHandler.SettlementRegex.Replace(value, ",").Trim().Trim(',').Trim();
            }

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
                    var name = TryTemplateName(match.Groups["name"].Value);
                    var type = match.Groups["type"].Value;
                    if (type == "дп")
                    {
                        if (oktmoHelper.HasEqualNearCity(name))
                        {
                            NearCityCell.Value = name;
                            TypeOfNearCityCell.Value = "дачный поселок";

                            var spec = new NearCitySpecification(name);
                            oktmoHelper.SetSpecifications(spec);
                            continue;
                        }
                    }

                    SntKpsCell.Value = SntKpsCell.Value == "" ? name : ", " + name;

                    match = match.NextMatch();
                }
                value = tmpRegex.Replace(value, ",");
            }


            //Поиск населенного пункта
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

                    var name = TryTemplateName(match.Groups["name"].Value);
                    var type = TryDescriptTypeOfNasPunkt(match.Groups["type"].Value);

                    tryAgainNCInNC:
                    //Если мы впервые нашлим населенный пункт
                    if (NearCityCell.Value == "" || NearCityCell.Value == string.Empty)
                    {
                        NearCityCell.Value = name;
                        TypeOfNearCityCell.Value = type;

                        if (oktmoHelper.HasEqualNearCity(name.ToLower()))
                        {
                            var spec = new NearCitySpecification(name);
                            oktmoHelper.SetSpecifications(spec);

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
                        if(oktmoHelper.HasEqualNearCity(name.ToLower()))
                        {                            
                            var spec = new NearCitySpecification(name);
                            oktmoHelper.SetSpecifications(spec);

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
                TryFindProperName(ref value);
            }

            NearCityCell.InitValue = value;
            //Если у нас что-то не разобрано, мы его пихаем в доп инфо или ту же ячейек
            if (NearCityCell.InitValue.Length > 2)
            {
                LandMarkCell.Value += NearCityCell.InitValue + ", ";
            }
            //Если у нас разобрано всё, а в ячейку населенного пункта ничего записано не было
            //Мы очищаем ячейку
        }

        private void CheckRegionCell()
        {
            if (string.IsNullOrEmpty(RegionCell.InitValue)) return;

            //Удаляем дублируем инфомарцию о субъекте из ячейки мун образование
            if (!string.IsNullOrEmpty(SubjectCell.InitValue))
                RegionCell.InitValue = RegionCell.InitValue.Replace(SubjectCell.InitValue, ", ").Trim().Trim(',').Trim();

            if (string.IsNullOrEmpty(RegionCell.InitValue)) return;

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

                var fullName = supportWorksheets.OKTMOWs.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value),
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

            var value = RegionCell.InitValue;
            TryFillRegion(ref value);
            RegionCell.InitValue = value;
            if (RegionCell.InitValue.Length <= 2) return;

            //На наличие поселения
            match = regexpHandler.SettlementRegex.Match(RegionCell.InitValue);
            //Если есть совпадение и оно не на всю строку
            if (match.Success)
            {
                var name = TryTemplateName(match.Groups["name"].Value);
                var type = match.Groups["type"].Value;
                type = type.IndexOf("п", StringComparison.OrdinalIgnoreCase) >= 0
                    ? "сельское поселение"
                    : "сельсовет";

                var fullName = name + " " + type;
                SettlementCell.Value = fullName;

                //В выборке уже имеется субъект и возможно Регион(или ВГТ)
                if (oktmoHelper.HasEqualSettlement(fullName))
                {
                    var spec = new SettlementSpecification(fullName);
                    oktmoHelper.SetSpecifications(spec);
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
                    else if (NearCityCell.Value == "") //bug ячейка ещё не проверена
                    {
                        NearCityCell.Valid = false;
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                }

                RegionCell.InitValue = regexpHandler.SettlementRegex.Replace(RegionCell.InitValue, ", ").Trim().Trim(',').Trim();
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
                var newName = TryTemplateName(match.Groups["name"].Value);
                SntKpsCell.Value = SntKpsCell.Value == "" ? newName : ", " + newName;
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ").Trim().Trim(',').Trim();
            }


            //На наличие населенного пункта и его типа
            tmpRegex = regexpHandler.NearCityRegex;
            var matches = tmpRegex.Matches(RegionCell.InitValue);
            var switched = false;
            if (matches.Count == 0)
            {
                tmpRegex = regexpHandler.NearCityToLeftRegex;
                matches = tmpRegex.Matches(RegionCell.InitValue);
            }
            //Если есть совпадение
            if (matches.Count > 0)
            {
                //Приоритет у любого негорода
                //если таковой есть
                if (matches.Count > 1)
                {
                    match =
                        matches.Cast<Match>()
                            .FirstOrDefault(
                                m => !Regex.IsMatch(m.Groups["type"].Value, "\bг", RegexOptions.IgnoreCase)) ??
                        matches[0];
                }
                else
                    match = matches[0];

                var name = TryTemplateName(match.Groups["name"].Value);
                var type = TryDescriptTypeOfNasPunkt(match.Groups["type"].Value);
            tryAgainNC:
                
                //В выборке уже имеется Субъект и вохможно Регион(или ВГТ) и возможно поселение
                //Урезаем выборку если возможно
                if (oktmoHelper.HasEqualNearCity(name.ToLower()))
                {
                    var spec = new NearCitySpecification(name);
                    oktmoHelper.SetSpecifications(spec);
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
                    if (oktmoHelper.HasEqualNearCity(name.ToLower()))
                    {
                        var spec = new NearCitySpecification(name);
                        var newCustomRowsList = oktmoHelper.CustomOktmoRows.FindAll(r => spec.IsSatisfiedBy(r));
                        //Обновляем тип по найденному нас пункту если возможно
                        if (newCustomRowsList.Count == 1)
                        {
                            var newType = newCustomRowsList.First().TypeOfNearCity;

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
                        NearCityCell.Valid = false;
                        SubjectCell.Valid = false;
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                }

                NearCityCell.Value = name; //Пишем найденное наименование в нужную ячейку
                TypeOfNearCityCell.Value = type;

                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ").Trim().Trim(',').Trim();
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

        private void TryFillRegion(ref string content,  Regex reg = null)
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


            var match = matches.Count > 1 ?
                matches.Cast<Match>().FirstOrDefault(m => !Regex.IsMatch(m.Groups["type"].Value, "(^|\b)г")) ??
                matches.Cast<Match>().First() : matches.Cast<Match>().First();


            var name = TryChangeRegionEndness(TryTemplateName(match.Groups["name"].Value));
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


            //Spet 1: Подходит ли регион к субъекту
            if (!string.IsNullOrEmpty(fullName) && oktmoHelper.SubjectHasEqualRegion(fullName.ToLower()))
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
                    if(oktmoHelper.HasEqualRegion(fullName.ToLower()))
                    {
                        var spec = new RegionSpecification(fullName);
                        oktmoHelper.SetSpecifications(spec);

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
                    //todo проверка на населенный пункт
                    //Или просто оставляем как есть до проверки населенного пункта
                }
                else
                {
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
                                SubjectCell.Valid = true;
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
            content = match.Value == content.Trim() ? "" : content.Replace(match.Value,", ").Trim().Trim(',').Trim();
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

                oktmoHelper.SetSubjectRows(subjRows);
                oktmoHelper.ResetToSubject();

                SubjectCell.Valid = true;

                //Get RegCenter
                regCenter = supportWorksheets.OKTMOWs.GetDefaultRegCenterFullName(subjectName, ref regName);
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
                result = result * 1000;

            return result;
        }

        /// <summary>
        ///     Метод пытается найти Имена собственные в переданной строке, и пытается их опеределить к какому-либо тиипу (мунОбр,
        ///     населПункт, ВГТ и прочие)
        /// </summary>
        /// <param name="value"></param>
        private void TryFindProperName(ref string value)
        {
            if (oktmoHelper.CustomOktmoRows == null) return;
            if (!oktmoHelper.CustomOktmoRows.Any()) return;

            var match = regexpHandler.WordWithHeadLetteRegex.Match(value);
            while (match.Success)
            {
                //does not match region and near city
                //and does not match SNT (or it`s just empty)
                if (match.Value != RegionCell.Value && match.Value != NearCityCell.Value &&
                    (SntKpsCell.Value == "" ||
                     (SntKpsCell.Value.IndexOf(match.Value,
                         StringComparison.OrdinalIgnoreCase) == -1)))
                {
                    //Пробуем подогнать к каждой ячейке
                    //Если никуда не подошло то пишем в первую пустую

                    //Try append to Region
                    var fullName = OKTMORepository.GetFullName(oktmoHelper.CustomOktmoRows, "город" + " " + match.Value,
                        OKTMOColumn.Region); //Tty to find on whole OKTMO
                    if (!string.IsNullOrEmpty(fullName))
                    {
                        if (!cellsFilled)
                        {
                            //Найденный регион пишем только если он подходит к выборке
//                            if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, fullName, OKTMOColumn.Region))
                            if (oktmoHelper.HasEqualRegion(fullName.ToLower()))
                            {
                                RegionCell.Value = fullName;

                                RegionCell.Valid = true;
                                SubjectCell.Valid = true;
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                                SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;

                                //Делаем выборку только если найденный регион не является региональным центром
                                if (!string.Equals(fullName, regCenter, StringComparison.OrdinalIgnoreCase))
                                {
                                    var spec = new RegionSpecification(fullName);
                                    oktmoHelper.SetSpecifications(spec);
                                }
                            }
                        }
                    }
                    //Try append to NearCity
                    else if (oktmoHelper.HasEqualNearCity(TryTemplateName(match.Value).ToLower()))
                    {
                        if (!cellsFilled)
                        {
                            var newName = TryTemplateName(match.Value);

                            NearCityCell.Valid = true;
                            RegionCell.Valid = true;
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;

                            NearCityCell.Value = newName;

                            if (!string.Equals(newName, regName, StringComparison.OrdinalIgnoreCase))
                            {
                                var spec = new NearCitySpecification(newName);
                                oktmoHelper.SetSpecifications(spec);
                            }
                        }
                    }
                    //Try Append To VGT
                    else if (supportWorksheets.VgtWorksheet.TerritotyExists(match.Value))
                    {
                        var v = TryTemplateName(match.Value);
                        TryFillVGT(ref v);
                    }
                    //Just Wtire to first epmty cell
                    else
                    {
                        if (StreetCell.Value == "" &&
                            Regex.IsMatch(match.Value, @"ая\b", RegexOptions.IgnoreCase))
                        {
                            StreetCell.Value = TryTemplateName(match.Value);
                            TypeOfStreetCell.Value = "улица";
                        }
                        else if (NearCityCell.Value == "")
                        {
                            NearCityCell.Value = TryTemplateName(match.Value);
                            NearCityCell.Valid = false;
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            if (RegionCell.Value != "")
                            {
                                RegionCell.Valid = false;
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                            else
                            {
                                SubjectCell.Valid = false;
                                SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                        }
                        else
                            goto skipWordReplace;
                    }

                    value = regexpHandler.WordWithHeadLetteRegex.Replace(value, ", ").Trim().Trim(',').Trim();
                }
            skipWordReplace:
                match = match.NextMatch();
            }
        }

        private bool TryFillVGT(ref string value)
        {
            //----Обрабатываем ВГТ-----
            if (string.IsNullOrEmpty(value)) return false;
            
            //Подтверждаем, что это ВГТ
            if (!supportWorksheets.VgtWorksheet.TerritotyExists(value)) return false;

            var vgt = value;

            if (RegionCell.Value != "" && RegionCell.Valid &&
                supportWorksheets.VgtWorksheet.CombinationExists(RegionCell.Value, vgt))
            {
                VgtCell.Value = vgt;
            }
            
            return true;

            //todo Нужно ли нам записывть населенный пункт?
            //а не записывается ли он потом, по найденному городу

            //Далее идут ситации если текущий насел пункт пустой, или не подходит к найденному ВГТ

            //Пробуем определить населенный пункт
            var city = string.Empty;


            //Пробуем извлечь текущий насел пункт из мунОбр
            //И тем самым подтвердить мунОбр и проставить населПункт
            if (RegionCell.Value != "" &&
                RegionCell.Value.IndexOf("город", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                city = TryTemplateName(RegionCell.Value.Replace("город", ""));
                city = city.Trim();
            }
            if (!string.IsNullOrEmpty(city) && supportWorksheets.VgtWorksheet.CombinationExists(city, vgt))
            {
                NearCityCell.Value = city;
                TypeOfNearCityCell.Value = "город";

                //Проверяем найденный насел пункт
                if (oktmoHelper.HasEqualNearCity(city))
                {
                    NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.Valid);

                    var spec = new NearCitySpecification(city);
                    oktmoHelper.SetSpecifications(spec);
                }
                else
                {
                    NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
                    RegionCell.SetStatus(DataCell.DataCellStatus.InValid);
                }
            }

            //В ином случае пробуем записать насел пункт через ВГТ
            else
            {
                var newCity = cellsFilled
                    ? string.Empty
                    : supportWorksheets.VgtWorksheet.GetCityByTerritory(vgt);

                //Строка будет  заполнена, если существует всего один насел пункт с таким районом
                if (string.IsNullOrEmpty(newCity)) return true;

                //нужно ли нам вообще проверять найденный
                if (NearCityCell.Value != "" &&
                    string.Equals(NearCityCell.Value, newCity,
                        StringComparison.CurrentCultureIgnoreCase)) return true;

                //Если текущий населенный пункт верный (он не пуст и не окрашен как неверный)
                //мы его оставляем на месте, а найденный пишем в ориентир
                if (NearCityCell.Value != "" && NearCityCell.Valid)
                    //Пишем найденный насел пункт в ориентир
                    LandMarkCell.Value += "город " + vgt + ", ";

                //В остальных случаях найденный насел пункт попадёт в ячейку населенногоп пункта
                else
                {
                    //Определяем, относится ли насел пункт к выборке
                    if (oktmoHelper.HasEqualCityType(newCity))
                    {
                        NearCityCell.Valid = true;
                        RegionCell.Valid = true;
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;


                        //Try to fill
                        var spec = new NearCitySpecification(newCity);
                        oktmoHelper.SetSpecifications(spec);
                    }
                    else
                    {
                        NearCityCell.Valid = false;
                        RegionCell.Valid = false;
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }

                    //Перекидываем текущий насел пункт
                    if (NearCityCell.Value != "")
                        LandMarkCell.Value += NearCityCell + ", ";

                    NearCityCell.Value = newCity;
                }
            }
            return true;
        }

        private void TryFillStreet(ref string value)
        {

            //Поиск улиц
            var regs = new List<Regex> { regexpHandler.StreetToLeftRegex, regexpHandler.StreetRegex };
            foreach (var reg in regs)
            {
                var match = reg.Match(value);
                if (match.Success)
                {
                    //По сути если у нас уже проставлена улица, новую нужно игнорировать кроме нескольких случаев ниже

                    //Берём только первое совпадение!
                    var name = ReplaceYo(TryTemplateName(match.Groups["name"].Value));
                    var type = ReplaceYo(TryDescriptTypeOfStreet(match.Groups["type"].Value));

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
                            BuildsCell.Value = match.Groups["house_num"].Value;
                    }

                    value = reg.Replace(value, ", ").Trim().Trim(',').Trim();
                }
            }
        }


        /// <summary>
        ///     Метод на основе сложившейся выборки пытается заполнить 100%-тные поля
        /// </summary>
        private void TryFillClassificator()
        {

            if (oktmoHelper.CustomOktmoRows == null) return;
            if (!oktmoHelper.CustomOktmoRows.Any()) return;
            if (cellsFilled) return;

            //
            //Записываем город если он у нас один 
            //
            var cities = oktmoHelper.CustomOktmoRows.Select(r => r.NearCity).Where(s => !string.IsNullOrEmpty(s)).Distinct().ToArray();
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
                    oktmoHelper.CustomOktmoRows.Select(r => r.TypeOfNearCity)
                        .Where(s => !string.IsNullOrEmpty(s))
                        .Distinct()
                        .ToArray();

                if (types.Count() == 1)
                {
                    var validType = types.First();
                    if (!TypeOfNearCityCell.Value.EqualNoCase(validType))
                    {
//                        if (TypeOfNearCityCell.Value == "")
//                        {
                            TypeOfNearCityCell.Value = validType;
                            TypeOfNearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                            NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
//                        }
//                        else
//                        {
//                            TypeOfNearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
//                            NearCityCell.SetStatus(DataCell.DataCellStatus.InValid);
//                        }
                    }
                }

                else if (TypeOfNearCityCell.Value != "" &&
                        TypeOfNearCityCell.Valid)
                {
                    if (TypeCanClarifyRows(types))
                    {
                        var spec =
                            new ExpressionSpecification<OktmoRow>(
                                r => (r.TypeOfNearCity ?? "").EqualNoCase(TypeOfNearCityCell.Value));

                        oktmoHelper.SetSpecifications(spec);

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
            var settlements = oktmoHelper.CustomOktmoRows.Select(r => r.Settlement).Distinct().ToArray();
                
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

//                    if(string.Equals(SettlementCell.Value, validSettlement,
//                        StringComparison.OrdinalIgnoreCase))
//                    {
                        SettlementCell.SetStatus(DataCell.DataCellStatus.Valid);
                        RegionCell.SetStatus(DataCell.DataCellStatus.Valid);
//                    }
//                    else
//                    {
//                        SettlementCell.SetStatus(DataCell.DataCellStatus.InValid);
//                        RegionCell.SetStatus(DataCell.DataCellStatus.InValid);    
//                    }
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
                    if (settlements.Any(s => s == "") && !NearCityCell.Valid && NearCityCell.Value != "")
                    {
                        SettlementCell.SetStatus(DataCell.DataCellStatus.Valid);
                        NearCityCell.SetStatus(DataCell.DataCellStatus.Valid);
                    }
                }
            }


            //
            //Записываем регион (муниципальное образование)
            //
            var regions = oktmoHelper.CustomOktmoRows.Select(r => r.Region).Where(s => !string.IsNullOrEmpty(s)).Distinct().ToArray();
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
                        if (SubjectCell.Value != "" )
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

            if (oktmoHelper.CustomOktmoRows.Count == 1)
                cellsFilled = true;
        }

        private bool TypeCanClarifyRows(string[] types)
        {
            return types.Any(
                s => TypeOfNearCityCell.Value.EqualNoCase(s)) &&
                   NearCityCell.Value != "" && 
                   oktmoHelper.CustomOktmoRows.Any(
                       r => (r.NearCity ?? "").EqualNoCase(NearCityCell.Value) &&
                            (r.TypeOfNearCity ?? "").EqualNoCase(NearCityCell.Value));
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
            if (Regex.IsMatch(s, @"\bд(ер(евн[а-я]*)?)?\.?", RegexOptions.IgnoreCase))
                s = "деревня";
            else if (Regex.IsMatch(s, @"г(ород[а-я]*|\.|\b)?", RegexOptions.IgnoreCase))
                s = "город";
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
            else if (Regex.IsMatch(s, @"\bб", RegexOptions.IgnoreCase))
                s = "бульвар";
            else if (Regex.IsMatch(s, @"\bпрос", RegexOptions.IgnoreCase))
                s = "проселок";
            else if (Regex.IsMatch(s, @"\bпр\-?т?", RegexOptions.IgnoreCase))
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
        private static string TryTemplateName(string s)
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
                var newWord = Regex.Replace(match.Value, justWordPattern,
                    m => string.Format("{0}{1}", m.Groups[1].Value.ToUpper(), m.Groups[2].Value.ToLower()));
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
            var sentenceReg = new Regex(@"(?n)(?<=(^|\b))(?!\.)[^\!\?$^]{5,}?(?=((?<!\s\w{1,2})\.|\!|\?|$))",
                RegexOptions.Multiline);
            var sentencesMatchCollection = sentenceReg.Matches(val);


            //Общие регулярки
            const string wordUnions = @"\s*(и|\,|;|:)\s*";
            //Пунктуационанные знаки в предложении, объединяющие части предложения
            const string sentenceEnds = @"\s*((<!\\s\w{1,4})\.|!|?)";
            //Пунктационные знаки, обозначающие конец предложения

            const string justWords = @"\(w(\w)*|\s)+"; //Паттерн для выялвения просто слов

            const string startCollocation = @"(?<=^|\""|(?<!\s\w{1,2})\.|\,|\)|\()";
            //Символы, обозначающие начало предложения
            const string endCollocation = @"(?=$|\""|(?<!\s\w{1,3})\.(\s|$|\,)|\,|\)|\()";
            //Символы, обозначающие конец предложения

            const string orDel = @"|"; //Символ Или
            const string spacesNRq = @"\s*"; //Наличие пробела в кол-ве от нуля до бесконечности

            //==========
            //Статус
            //TODO дефолтный значения при точномм, не точном наличии или отсутствии
            //Перечень фраз для подтверждения наличия коммуникации
            const string comValid =
                @"(?<valid>круглый\sгод|всегда|подведен(о|ы)|централизирован(а|о)|(?!в\sобществе\s)проводят|провед(ё|е)но?(?!\sк\sгранице)|на\s(участке|территории)|есть(?!\s*возможность)|име(е|ю)тся|(?<kvt>\d(\d|\.|\,)*)\s*квт)";
            const string comCanConnectAlwaysLeft = @"";
            const string comCanConnectAlwaysRight = @"";
            //Перечень фраз для подтверждения возможности провести коммуникацию
            const string comCanConnect =
                @"(?<canconn>в\sперспективах|\bТУ\b|проводится|будет|проведут|в\sобществе\sпроводят|легко\sпровести|оплачивается\sотдельно|(проведен\s(к|по))?границе|подключение(\sту)?|рядом\sпроходит|(есть\s)?возможно(сть)?|в\s\d+\sм(\.|етрах|\s)|актуально(\sпровести)?|разешени(е|я)|около|техусловия|соласовано|(на|по)\sулице|не\sдалеко)";
            const string comNo = @"(?<no>нет|отсутству(е|ю)т)"; //Фразы, подтверждающие отсутствие коммуникации
            const string comTemp = @"(?<temp>летний|зимний)"; //Наличие сезонной коммуникации


            const string delimCom = @"(\s*(\,|\.)\s*)"; //Символы разделители между преречисленными коммуникациями

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
            const string reliefNames = @"(?<relief>ровный)";

            //==========
            //Дорога
            //Bug есть подъезд = есть дорога?
            const string roadNames = @"(?<road>асфальт|грунтовая|засыпана)";
            //======================
            //======================


            //Строка-паттерн-перечень всех вохможных коммуникаций для выявления одного
            const string anyCom =
                "(?<anyCom>" + commonCommunicatuionNames + orDel + electrNames + orDel + waterNames + orDel + gasNames +
                orDel + severageNames + ")";

            //Строка-паттерн-перечень всех возможных коммуникаций для выявления их в прямой последновательности
            const string stringOfAnyCom = "(" + delimCom + anyCom + @"|\s*\,\s*";


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
    }

    public class SupportWorksheets
    {
        public SupportWorksheets(CatalogWorksheet catalogWorksheet, OKTMORepository oktmo,
            SubjectSourceWorksheet subjectSourceWorksheet, VGTWorksheet vgtWorksheet)
        {
            CatalogWs = catalogWorksheet;
            OKTMOWs = oktmo;
            SoubjectSourceWorksheet = subjectSourceWorksheet;
            VgtWorksheet = vgtWorksheet;
        }

        public CatalogWorksheet CatalogWs { get; private set; }
        public OKTMORepository OKTMOWs { get; private set; }
        public SubjectSourceWorksheet SoubjectSourceWorksheet { get; private set; }
        public VGTWorksheet VgtWorksheet { get; private set; }
    }

    public class DataCell:IDisposable
    {
        public enum DataCellStatus
        {
            Valid,
            InValid
        }
        public DataCell(ExcelRange cell)
        {
            Cell = cell;
            InitValue = (Cell.Value ?? "").ToString().Replace("ё","е").Replace("Ё","Е");
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
