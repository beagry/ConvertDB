using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary.SupportEntities.Oktmo;
using Formater.SupportWorksheetsClasses;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Formater
{
    public class ExcelLocationRow
    {
        private readonly ExcelWorksheet worksheet;
        private readonly SupportWorksheets supportWorksheets;
        private readonly int row;
        private readonly int descriptionColumn;

        private DataTable customTable;
        private DataTable subjectTable;
        private List<OktmoRow> customOktmoRows; 
        private List<OktmoRow> subjectOktmoRows; 
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
                TypeOfnearCityCell = new DataCell(worksheet.Cells[row, typeOfNearCityColumn]);
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
        }

        private ExcelLocationRow(XlTemplateWorkbookType wbType, SupportWorksheets supportWorksheets)
        {
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
        public DataCell TypeOfnearCityCell { get; set; }
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
            TypeOfnearCityCell.Save();
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
            

            SaveCells();
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
                    if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, newCity, OKTMOColumns.NearCity))
                    {
                        NearCityCell.Value = newCity;
                        TypeOfnearCityCell.Value = "город";

                        customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                            new SearchParams(newCity, OKTMOColumns.NearCity));
                        TryFillClassificator();
                    }
                }
                //или ставим муниципальное образование как город
                //При условии что это не региональный центр
                else if (RegionCell.Value != "" &&
                         RegionCell.Cell.Style.Fill.BackgroundColor.Rgb != ExcelExtensions.BadColor.ToArgb().ToString()
                         && RegionCell.Value.IndexOf("город") >= 0)
                {
                    var name = RegionCell.Value.Replace("город", "");
                    name = name.Replace("(ЗАТО)", "");
                    name = name.Trim();
                    if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                    {
                        cellsFilled = false;
                        NearCityCell.Value = name;
                        TypeOfnearCityCell.Value = "город";

                        customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                            new SearchParams(name, OKTMOColumns.NearCity));
                        TryFillClassificator();
                    }
                }
            }
            //Ставим дефолтное значение для муниципального образования, если оно пустое, а текущий насленный пункт у нас является региональным центро
            else if (RegionCell.Value == "" &&
                     string.Equals(NearCityCell.Value, regName, StringComparison.OrdinalIgnoreCase))
            {
                customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                    new SearchParams(regName, OKTMOColumns.NearCity));
                TryFillClassificator();
            }
            //Дефолное значение для типа населенного пункта, если найденный насел пункт совпадает по названию с региональным центром
            else if (TypeOfnearCityCell.Value == "" &&
                     string.Equals(NearCityCell.Value, regName, StringComparison.OrdinalIgnoreCase))
            {
                TypeOfnearCityCell.Value = "город";
            }
        }

        private void CheckDescriptionCell()
        {
//            var descriptionColumn = GetColumnIndex("DESCRIPTION");

            //Вначале мы ищем наименования по типу
            //После мы пытаемся отнести найдненные в описании Именования без типов
            var cell = worksheet.Cells[row, descriptionColumn];
            if ((string) cell.Value == "") return;

            var descrtContent = ReplaceYo((cell.Value??"").ToString());

            //
            //----Товарищества
            //

            var match = regexpHandler.SntToLeftRegex.Match(descrtContent);
            if (match.Success)
            {
                do
                {
                    //Берём только первое совпадение!
                    var name = TryTemplateName(match.Groups["name"].Value);

                    SntKpsCell.Value = SntKpsCell.Value == "" ||
                                       SntKpsCell.Value.Length < 3
                        ? name
                        : ", " + name;
                    descrtContent = regexpHandler.SntToLeftRegex.Replace(descrtContent, ", ");
                    match = match.NextMatch();
                } while (match.Success);
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
                    OKTMOColumns.Subject);
                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value != "" &&
                    !string.Equals(SubjectCell.Value.Trim(), fullName.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    rowsToDelete.Add(row);
                    SubjectCell.Value = fullName;
                    subjectTable = supportWorksheets.OKTMOWs.GetCustomDataTable(new SearchParams(fullName, OKTMOColumns.Subject));
                    customTable = subjectTable != null ? subjectTable.Copy() : supportWorksheets.OKTMOWs.Table.Copy();
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
            Regex reg;
            foreach (var regi in regs)
            {
                reg = regi;

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
                if ((string.Equals(name, regName, StringComparison.OrdinalIgnoreCase) &&
                     (RegionCell.Cell.Style.Fill.BackgroundColor.Rgb !=
                      ExcelExtensions.BadColor.ToArgb().ToString()))) continue;

                if(!string.Equals(name, regName, StringComparison.OrdinalIgnoreCase)) continue;

                if (NearCityCell.Value != "" && (string.Equals(NearCityCell.Value, name,
                         StringComparison.OrdinalIgnoreCase))) continue;
                if (SubjectCell.Valid && RegionCell.Valid && NearCityCell.Valid) continue;

                if (type == "город" && TypeOfnearCityCell.Value != "" &&
                    TypeOfnearCityCell.Value != "город")
                {
                    LandMarkCell.Value += name + " " + type + ", ";
                }
                else
                {
                    //BackUp current value
                    if (NearCityCell.Value != "")
                        LandMarkCell.Value += TypeOfnearCityCell.Value + " " +
                                              NearCityCell.Value;

                    //Обнуляем МунОбразование
                    //сейчас стоит региональный центр или просто город
                    //а найденный насел пункт подходит к другому мун образованию
                    var itIsCity = (RegionCell.Value != "" &&
                                    string.Equals(RegionCell.Value, regCenter,
                                        StringComparison.OrdinalIgnoreCase) ||
                                    (RegionCell.Value != "" &&
                                     RegionCell.Value
                                         .IndexOf("город", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                     type != "город") ||
                                    (NearCityCell.Value != "" && TypeOfnearCityCell.Value == ""));

                    var valueNeedsResetRegion =
                        !supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity) &&
                        supportWorksheets.OKTMOWs.StringMatchInColumn(subjectTable, name, OKTMOColumns.NearCity);

                    if (itIsCity && valueNeedsResetRegion)
                    {
                        customTable = subjectTable != null ? subjectTable.Copy() : supportWorksheets.OKTMOWs.Table.Copy();
                        RegionCell.Value = string.Empty;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        SettlementCell.Value = string.Empty;
                        SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }

                    const string dashPattern = @"\s*\-\s*";
                    const string spacePattern = @"\s+";
                    const string cityEnd = @"(е|а)\b";

                    //найденный насел пункт подхоидт к нашей выборке (по субъекту и возможно по мунобразованию если оно есть)
                    if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                        customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                            new SearchParams(name, OKTMOColumns.NearCity));
                    else
                    {
                        if (!switched)
                        {
                            switched = true;
                            if (Regex.IsMatch(name, dashPattern))
                            {
                                name = Regex.Replace(name, dashPattern, " ");
                                goto tryGetNearCityAgain;
                            }

                            if (Regex.IsMatch(name, spacePattern))
                            {
                                name = Regex.Replace(name, spacePattern, "-");
                                goto tryGetNearCityAgain;
                            }
                        }

                        if (!endChanged)
                        {
                            if (type == "город" && Regex.IsMatch(name, cityEnd, RegexOptions.IgnoreCase))
                            {
                                name = Regex.Replace(name, cityEnd, "");
                                goto tryGetNearCityAgain;
                            }

                            endChanged = true;
                            name = startName;
                        }
                        //Дробим имя собственное если возможно для поиска по каждому имени отдельни
                        if (!splitted)
                        {
                            //Step one: we split it
                            if (words == null)
                            {
                                var patterns = new List<string> { dashPattern, spacePattern };

                                foreach (var pattern in patterns)
                                {
                                    if (!Regex.IsMatch(startName, pattern)) continue;

                                    words = Regex.Split(startName, pattern).ToList();
                                    name = words.Last();
                                    words[words.Count - 1] = null;
                                    goto tryGetNearCityAgain; //just break
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
                                splitted = true;
                                name = startName;
                            }
                        }
                    }

                    NearCityCell.Value = name;
                    TypeOfnearCityCell.Value = type;

                    cellsFilled = false;
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

            //=================
            //Коммуникацияя
            //=================
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
                    LandMarkCell.InitValue = LandMarkCell.InitValue.Replace(tmpValue, ", ");
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

            value = value.Replace(tmpValue, ", ");
            VgtCell.InitValue = value;
        }

        private void CheckNearCityCell()
        {
            //
            //Разбираем Населенный пункт
            //
            var value = NearCityCell.InitValue;
            //Удаляем дублируемуб инфомарцию о субъекте
            if (!string.IsNullOrEmpty(SubjectCell.InitValue))
                value = value.Replace(SubjectCell.InitValue, ", ");

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
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ");
                var fullName = supportWorksheets.OKTMOWs.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumns.Subject);

                if (!string.IsNullOrEmpty(fullName) &&
                    SubjectCell.Value != "" &&
                    SubjectCell.Value
                        .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                {
                    rowsToDelete.Add(row);
                    SubjectCell.Value = fullName;
                    subjectTable = supportWorksheets.OKTMOWs.GetCustomDataTable(new SearchParams(fullName, OKTMOColumns.Subject));
                    customTable = subjectTable != null ? subjectTable.Copy() : supportWorksheets.OKTMOWs.Table.Copy();

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
                    OKTMOColumns.Region);
                if (!string.IsNullOrEmpty(fullName)) //This is REGION
                {
                    if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, fullName,
                        OKTMOColumns.Region))
                    {
                        RegionCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                        //Выборка
                        customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                            new SearchParams(fullName, OKTMOColumns.Region));
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
                value = tmpRegex.Replace(value, ", ");
                if (value.Length <= 2) return;
            }

            //Поиск киллометров до населенного пункта
            match = regexpHandler.DistToRegCenteRegex.Match(value);
            if (match.Success)
            {
                //Спихиваем всё в столбец "Расстояние до рег центра"
                //Разбирать будем в конце
                DistToRegCenterCell.Value += ", " + match.Value;
                value = regexpHandler.DistToRegCenteRegex.Replace(value, ", ");
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

                if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, fullName, OKTMOColumns.Settlement))
                {
                    customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                        new SearchParams(fullName, OKTMOColumns.Settlement));
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
                value = regexpHandler.SettlementRegex.Replace(value, ",");
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
                        TypeOfnearCityCell.Value = type;

                        if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                        {
                            customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                new SearchParams(name, OKTMOColumns.NearCity));

                            RegionCell.Valid = true;
                            NearCityCell.Valid = true;
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
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
                        if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                            //и подходит ли она к нам
                        {
                            customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                new SearchParams(name, OKTMOColumns.NearCity));

                            LandMarkCell.Value += NearCityCell.Value + ", ";

                            NearCityCell.Valid = true;
                            RegionCell.Valid = true;
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                            NearCityCell.Value = name;
                            TypeOfnearCityCell.Value = type;
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
                value = tmpRegex.Replace(value, ", ");
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
                //Как бы зачем оставлять "3б" в населенном пункте
                //В зависимости от была ли внесена полезная инфомация в ячеку "населенный пункт"
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
                RegionCell.InitValue = RegionCell.InitValue.Replace(SubjectCell.InitValue, ", ");


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
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ");

                var fullName = supportWorksheets.OKTMOWs.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value),
                    OKTMOColumns.Subject);

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
                if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, fullName, OKTMOColumns.Settlement))
                    customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                        new SearchParams(fullName, OKTMOColumns.Settlement));
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

                RegionCell.InitValue = regexpHandler.SettlementRegex.Replace(RegionCell.InitValue, ", ");
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
                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ");
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
                if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                    customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                        new SearchParams(name, OKTMOColumns.NearCity));
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
                    if (supportWorksheets.OKTMOWs.StringMatchInColumn(subjectTable, name, OKTMOColumns.NearCity))
                    {
                        //BUG поселение уже может быть окрашено в красный
                        var newTable = supportWorksheets.OKTMOWs.GetCustomDataTable(subjectTable,
                            new SearchParams(name, OKTMOColumns.NearCity));
                        //Обновляем тип по найденному нас пункту если возможно
                        if (newTable.Rows.Count == 1)
                        {
                            string newType;
                            try
                            {
                                newType =
                                    newTable.Rows.Cast<DataRow>().First()[typeOfNearCityColumn - 1].ToString();

                                RegionCell.Valid = true;
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                                RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            }
                            catch (InvalidOperationException e)
                            {
                                throw e;
                            }

                            if (TypeOfnearCityCell.Value == "" || TypeOfnearCityCell.Value != newType)
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
                TypeOfnearCityCell.Value = type;

                RegionCell.InitValue = tmpRegex.Replace(RegionCell.InitValue, ", ");
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
            Match match;
            if (reg == null)
            {
                tmpRegex = regexpHandler.RegionRegex;
                match = tmpRegex.Match(content);
                if (!match.Success)
                {
                    tmpRegex = regexpHandler.RegionToLeftRegex;
                    match = tmpRegex.Match(content);
                }
            }
            else
            {
                tmpRegex = reg;
                match = tmpRegex.Match(content);
            }

            //Если есть совпадение
            if (match.Success)
            {
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
                var fullName = supportWorksheets.OKTMOWs.GetFullName(name, OKTMOColumns.Region, type);


                //Spet 1: Подходит ли регион к субъекту
                if (!string.IsNullOrEmpty(fullName) &&
                    supportWorksheets.OKTMOWs.StringMatchInColumn(subjectTable, fullName, OKTMOColumns.Region))
                {
                    //Отлично, найденное мунОбр-е относится к текущему субъекту
                    //подтверждаем что нам надо использовать найденный текст
                    if (RegionCell.Value == "" ||
                        (!string.Equals(RegionCell.Value, fullName, StringComparison.OrdinalIgnoreCase) &&
                         RegionCell.Cell.Style.Fill.BackgroundColor.Rgb == ExcelExtensions.BadColor.ToArgb().ToString()))
                    {
                        if (RegionCell.Value != "")
                            AppendToLandMarkCell(RegionCell.Value);

                        RegionCell.Value = fullName;

                        RegionCell.Valid = true;
                        SubjectCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                        //Выборка
                        if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, fullName, OKTMOColumns.Region))
                        {
                            customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                new SearchParams(fullName, OKTMOColumns.Region));


                            RegionCell.Valid = true;
                            NearCityCell.Valid = true;
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        }
                        //Окрашиваем Регион если он подходит к субъекту но не подходит к выборке
                        else
                        {

                            RegionCell.Valid = false;
                            NearCityCell.Valid = false;
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
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
                                if (RegionCell.Value != "" &&
                                    RegionCell.Cell.Style.Fill.BackgroundColor.Rgb ==
                                    ExcelExtensions.BadColor.ToArgb().ToString())
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
                content = match.Value == content.Trim() ? "" : tmpRegex.Replace(content, ", ");
            }
        }

        private void CheckSubejctCell()
        {
            var subjectName = supportWorksheets.SoubjectSourceWorksheet.GetSubjectBySourceLink(SourceLinkCell.InitValue);


            if (string.IsNullOrEmpty(subjectName))
                subjectName = supportWorksheets.OKTMOWs.GetFullName(SubjectCell.InitValue, OKTMOColumns.Subject);

            //определили
            if (!string.IsNullOrEmpty(subjectName))
            {
                //вставляем
                SubjectCell.Value = subjectName;

                //отбираем
                if (!supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, subjectName,
                    OKTMOColumns.Subject)) return;

                customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                    new SearchParams(subjectName, OKTMOColumns.Subject));
                subjectTable = customTable.Copy();
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
            if (customTable == null) return;
            if (subjectTable == null) return;

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
                    var fullName = OKTMORepository.GetFullName(subjectTable, "город" + " " + match.Value,
                        OKTMOColumns.Region); //Tty to find on whole OKTMO
                    if (!string.IsNullOrEmpty(fullName))
                    {
                        if (!cellsFilled)
                        {
                            //Найденный регион пишем только если он подходит к выборке
                            if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, fullName, OKTMOColumns.Region))
                            {
                                RegionCell.Value = fullName;

                                RegionCell.Valid = true;
                                SubjectCell.Valid = true;
                                RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                                RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                                SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                                //Делаем выборку только если найденный регион не является региональным центром
                                if (!string.Equals(fullName, regCenter, StringComparison.OrdinalIgnoreCase))
                                    customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                        new SearchParams(fullName, OKTMOColumns.Region));
                            }
                        }
                    }
                    //Try append to NearCity
                    else if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, TryTemplateName(match.Value),
                        OKTMOColumns.NearCity))
                    {
                        if (!cellsFilled)
                        {
                            var newName = TryTemplateName(match.Value);

                            NearCityCell.Valid = true;
                            RegionCell.Valid = true;
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                            NearCityCell.Value = newName;

                            if (!string.Equals(newName, regName, StringComparison.OrdinalIgnoreCase))
                                customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                    new SearchParams(newName, OKTMOColumns.NearCity));
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

                    value = regexpHandler.WordWithHeadLetteRegex.Replace(value, ", ");
                }
            skipWordReplace:
                match = match.NextMatch();
            }
        }

        private bool TryFillVGT(ref string value)
        {

            var res = false;
            //----Обрабатываем ВГТ-----
            if (!string.IsNullOrEmpty(value))
            {
                //Подтверждаем, что это ВГТ
                if (supportWorksheets.VgtWorksheet.TerritotyExists(value))
                {
                    var vgt = value;
                    VgtCell.Value = vgt;
                    res = true;

                    if (NearCityCell.Value != "" &&
                        supportWorksheets.VgtWorksheet.CombinationExists(NearCityCell.Value, vgt))
                        return true;

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
                        TypeOfnearCityCell.Value = "город";

                        //Проверяем найденный насел пункт
                        if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, city, OKTMOColumns.NearCity))
                        {
                            NearCityCell.Valid = true;
                            RegionCell.Valid = true;
                            NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                            RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);


                            customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                new SearchParams(city, OKTMOColumns.NearCity));
                            //                            TryFillClassificator(row, customTable, ref cellsFilled);
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
                    }
                    //В ином случае пробуем записать насел пункт через ВГТ
                    else
                    {
                        var newCity = cellsFilled
                            ? string.Empty
                            : supportWorksheets.VgtWorksheet.GetCityByTerritory(vgt);
                        if (!string.IsNullOrEmpty(newCity))
                        //Строка будет  заполнена, если существует всего один насел пункт с таким районом
                        {
                            //нужно ли нам вообще проверять найденный
                            if (NearCityCell.Value != "" &&
                                string.Equals(NearCityCell.Value, newCity,
                                    StringComparison.CurrentCultureIgnoreCase)) return res;

                            //Если текущий населенный пункт верный (он не пуст и не окрашен как неверный)
                            //мы его оставляем на месте, а найденный пишем в ориентир
                            if (NearCityCell.Value != "" &&
                                NearCityCell.Cell.Style.Fill.BackgroundColor.Rgb !=
                                ExcelExtensions.BadColor.ToArgb().ToString())
                                //Пишем найденный насел пункт в ориентир
                                LandMarkCell.Value += "город " + vgt + ", ";

                            //В остальных случаях найденный насел пункт попадёт в ячейку населенногоп пункта
                            else
                            {
                                //Определяем, относится ли насел пункт к выборке
                                if (supportWorksheets.OKTMOWs.StringMatchInColumn(customTable, newCity,
                                    OKTMOColumns.TypeOfNearCity))
                                {
                                    NearCityCell.Valid = true;
                                    RegionCell.Valid = true;
                                    NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                                    NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);


                                    //Try to fill
                                    customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                                        new SearchParams(newCity, OKTMOColumns.NearCity));
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
                    }
                }
            }
            return res;
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

                    value = reg.Replace(value, ", ");
                }
            }
        }


        /// <summary>
        ///     Метод на основе сложившейся выборки пытается заполнить 100%-тные поля
        /// </summary>
        private void TryFillClassificator()
        {

            if (customTable == null) return;
            if (cellsFilled) return;

            //
            //Записываем город если он у нас один 
            //
            var cities = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMORepository.Columns.NearCity - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (cities.Count == 1)
            {
                if (NearCityCell.Value == "")
                {
                    if (NearCityCell.Value != "")
                        LandMarkCell.Value += NearCityCell.Value + ", ";
                    NearCityCell.Value = cities.First();

                    NearCityCell.Valid = true;
                    RegionCell.Valid = true;
                    NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    NearCityCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                    NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
                else if (NearCityCell.Value != "" &&
                         string.Equals(NearCityCell.Value, cities.First(), StringComparison.OrdinalIgnoreCase))
                {
                    NearCityCell.Valid = true;
                    RegionCell.Valid = true;
                    NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    NearCityCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                    NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
            }
            //Проверяем текущее значение на принадлежность к выборке
            else
            {
                if ((NearCityCell.Value != ""))
                    if (cities.All(s => s != NearCityCell.Value.ToString()))
                    {
                        NearCityCell.Valid = false;
                        RegionCell.Valid = false;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else
                    {
                        NearCityCell.Valid = true;
                        RegionCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        NearCityCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
            }
            //
            //Записываем тип если он один
            //
            var types = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMORepository.Columns.TypeOfNearCity - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (types.Count == 1 && NearCityCell.Value != "")
            {
                if (TypeOfnearCityCell.Value == "" ||
                    !string.Equals(TypeOfnearCityCell.Value, types.First(), StringComparison.OrdinalIgnoreCase))
                {
                    TypeOfnearCityCell.Value = types.First();
                    TypeOfnearCityCell.Valid = true;
                    TypeOfnearCityCell.Valid = true;
                    TypeOfnearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    TypeOfnearCityCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                    TypeOfnearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    NearCityCell.Cell.Style.Fill.BackgroundColor.Indexed = 0;
//                    NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
            }

            //Если в нашей выборке нельсколько типов населенных пунктов
            //И мы уже имеем какой-то тип
            else if (TypeOfnearCityCell.Value != "")
                //Пробуем использовать тип для уточнения выборки
                if (
                    types.Any(
                        s => string.Equals(TypeOfnearCityCell.Value.ToString(), s, StringComparison.OrdinalIgnoreCase)) &&
                    //1. Наш тип находится в пределах выборки
                    NearCityCell.Value != "" && //2. у нас есть насел пункт
                    customTable.Rows.Cast<DataRow>()
                        .Any( //3. В выборке есть комбинация текущий насел пункт + текущий тип
                            r =>
                                string.Equals(r[OKTMORepository.Columns.NearCity - 1].ToString(),
                                    NearCityCell.Value.ToString(),
                                    StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(r[OKTMORepository.Columns.TypeOfNearCity - 1].ToString(),
                                    TypeOfnearCityCell.Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    //И тогда  мы уточняем выборку по типу населенного пункта
                    customTable = supportWorksheets.OKTMOWs.GetCustomDataTable(customTable,
                        new SearchParams(TypeOfnearCityCell.Value, OKTMOColumns.TypeOfNearCity));
                    {
                        TypeOfnearCityCell.Valid = true;
                        NearCityCell.Valid = true;
                        TypeOfnearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        TypeOfnearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
                }
                else
                {
                    TypeOfnearCityCell.Valid = false;
                    NearCityCell.Valid = false;
                    TypeOfnearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    TypeOfnearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    NearCityCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    NearCityCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                }

            //По возможности записываем поселение
            var settlements = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMORepository.Columns.Settlement - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (settlements.Count == 1)
            {
                if (SettlementCell.Value == "" ||
                    !string.Equals(SettlementCell.Value, settlements.First(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    if (SettlementCell.Value != "")
                        LandMarkCell.Value += SettlementCell.Value + ", ";
                    SettlementCell.Value = settlements.First();

                    SettlementCell.Valid = true;
                    RegionCell.Valid = true;
                    SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
                else if (SettlementCell.Value != "" &&
                         string.Equals(SettlementCell.Value, settlements.First(),
                             StringComparison.OrdinalIgnoreCase))
                {
                    SettlementCell.Valid = true;
                    RegionCell.Valid = true;
                    SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
            }
            else
            //Проверяем текущее значение на принадлежность к выборке
            {
                if (SettlementCell.Value != "")
                    if (settlements.All(s => s != SettlementCell.Value.ToString()))
                    {
                        SettlementCell.Valid = false;
                        RegionCell.Valid = false;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else
                    {
                        SettlementCell.Valid = true;
                        RegionCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        SettlementCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        SettlementCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
            }


            //
            //Записываем регион (муниципальное образование)
            //
            var regions = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMORepository.Columns.Region - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (regions.Count == 1)
            {
                if (RegionCell.Value == "" ||
                    !string.Equals(RegionCell.Value, regions.First(), StringComparison.OrdinalIgnoreCase))
                {
                    if (RegionCell.Value != "")
                        LandMarkCell.Value += RegionCell.Value + ", ";
                    RegionCell.Value = regions.First();
                    RegionCell.Valid = true;
                    SubjectCell.Valid = true;
                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
                else if (RegionCell.Value != "" &&
                         string.Equals(RegionCell.Value, regions.First(), StringComparison.OrdinalIgnoreCase))
                {
                    RegionCell.Valid = true;
                    SubjectCell.Valid = true;
                    RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                    SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
            }
            //Проверяем текущее значение на принадлежность к выборке
            else
            {
                if (RegionCell.Value != "")
                    if (regions.All(s => s != RegionCell.Value.ToString()))
                    {
                        RegionCell.Valid = false;
                        SubjectCell.Valid = false;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else
                    {
                        RegionCell.Valid = true;
                        SubjectCell.Valid = true;
                        RegionCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        RegionCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        SubjectCell.Cell.Style.Fill.PatternType = ExcelFillStyle.None;
//                        SubjectCell.Cell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
            }

            if (customTable.Rows.Count == 1)
                cellsFilled = true;
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

        public void Save()
        {
            Cell.Value = Value;
        }

        public void Dispose()
        {
            Cell.Dispose();
        }
    }

    public class OktmoCheckStatuss
    {
        public class TextBoolObject
        {
            public TextBoolObject(string text)
            {
                Text = text;
                Valid = false;
            }

            public TextBoolObject()
            {
                
            }

            public string Text { get; set; }
            public bool Valid { get; set; }
        }

        public OktmoCheckStatuss()
        {
            Subject = new TextBoolObject();
            Region = new TextBoolObject();
            Settlement = new TextBoolObject();
            NearCity = new TextBoolObject();
            CityType = new TextBoolObject();
        }

        public TextBoolObject Subject { get; set; }
        public TextBoolObject Region { get; set; }
        public TextBoolObject Settlement { get; set; }
        public TextBoolObject NearCity { get; set; }
        public TextBoolObject CityType { get; set; }
    }
}
