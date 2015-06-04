#define doRegion
#define doNearCity
#define doDescriptionPlace
#define fillDefault
#define checkVGT
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelRLibrary.TemplateWorkbooks;
using Formater.SupportWorksheetsClasses;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Action = System.Action;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

// ReSharper disable UnusedMember.Local

namespace Formater
{
    public partial class DbToConvert
    {
        //TODO начала и окончания регулярок не захватывать, заменить на (?<=) (?=)
        #region регулярки

        //это адовый пиздец
        //я копирвовал всю строку меж кавычек
        //и использовал http://regexhero.net/tester/
        private readonly Regex sntRegex =
            new Regex(
                @"(?:^|\""|\.|\,|\s|\))+\(?\""?(?<name>(\b\w[^\,\\\/\(\)\b\s]{2,}\b\s?){1,2})\)?\s+(?<type>(?i)\b(?:с|c)((\\|\/|\.)(т|с)(?<=\.)|(?:нт|от|п|т)|днп|днт|кп))(?:$|\.|\""|\,|\s)+");


        private readonly Regex sntToLeftRegex =
            new Regex(
                @"(?:^|\.|\,|\s|\)|\()+(?<type>(?<!(между|\bи|рядом(\sс)?|у|около|недалеко(\sот)?)\s)\b(?:СНТ|СКТ|ДНТ|ТИЗ|с-во|(?i:(?<!\d+\s)(?:с|c)(?:\\(?=\w\s)|\/(?=\w\s)|н|о(?=\s(?!Всем)[А-Я])|ад(?:ов(?!ая|ый)(?:од(?:че|ств(?:о|е)))?[а-я]{0,5}|\-о?м|ы|-в(?:о|е))|(?:\.|-)?т|от|нт|-в(о|e)|(?:\\|/)?о)\.?\s*(?:п|(некоммерческ[а-я]*\s)?т(?:ов?(?:арищ[а-я]{0,6}|\-в?е))?|общ(?:ест|\-)?в?о)?|днп|товариществ[а-я]{1,3}|смт|схк|д(ачное\s)?н(екоммерческое\s)?т(оварищество)?|кооп(ератив[а-я]*)|д(?:ач(?:н[а-я]{0,3})?)?(?(?<=\w{3,})\s|\s?)(к(?:ооп(?:ер[а-я]{0,4})?)?в?|п(ос(ело?ке?)?)?)\b|к(от(теджн[а-я]{0,2})?\s?)?(?:\s|\-|\\|\/|\.\s?|\,)*п(ос(ело?ке?)?)?)\b))(?:\.|\s|\()?\s*(?:\""|“|”|'|`|«|\&quot;)*(?<name>(?!(?:\b((?:(?#Само название не может начинатсья или оканчиватсья как...=>)Прозер|\d|Уник|с-т|Чист|Улиц|Дорог|Участ|Постройки)[а-я]*|[а-я]*(ом|ой|ого)))\b)((?#Приставки, что нужно брать)\s*завода\s*)?(?:\b?(?(?=\w+(\s+[а-я]{3,}){2,})[А-Я]|\w)[^\,\\\/\(\)\b\s\.""]+(?!(?#Проверка на слово справа)\s\b(морем|об\.?(л(\.|аст[а-я]{1,2}|\s))?|р(\-|ай)о?на?)\b)\b\s?|\s?№\d+){1,2}(?:\s\d+)?)(?:$|\""|”|“|»|\.|\,|\s|\)|'|`|&quot;|\()+");



        private Regex subjectRegex = new Regex(
            @"(?:^|\""|\.|\,|\s|\))+\""?(?<name>\b(?:\w|\s){2,})(?<type>(?i)\bобласть\b|\bреспубласть\b|\bокруг\b|\bкрай\b)(?:$|\.|\""|\,|\s|\)|\()+");


        private readonly Regex settlementRegex =
            new Regex(
                @"(?n)\(?\""?(?<name>(\b(?!\d)\w[^\,\\\/\(\)\b\s]{2,}\b\s?){1,2})(?<type>(?i)\b(с(\/|\\)?с|с(\/|\\)?(п)|с(\/|\\)о)\b)\)?($|\.|\""|\,|\s|\)|\()+");


        private readonly Regex distToRegCenteRegex =
            new Regex(
                @"(?i)(\(?(?:^|\""|\.|\,|\s)?(?<num>\d(?:\,(?=\d)|\d|\.(?=\d))*)\s*(?:км\b)?\s*(?:от|до)?\s*\""?(?<name>(?:\b\w[^\,\\\/\(\)\b\s]{2,}\b\s?){1,2})?(?:$|\.|\""|'|\,|\s|\))+|в\sчерте|за\s(чертой|городом))");


        private readonly Regex streetRegex =
            new Regex(
                @"(?n)(^|\.|\""|\,|\s|\)|\()+\""?(?<name>\b(?<!((пос)\.)\s)(?!(?i)коттеджн|имеет|Район|до\b|есть|зимой|от\b|круглый\sгод|за|\d+км|проход|как\b|так\b|по\b|при\b|ул\b|провед|км\b|без\b|ведут|газ\b|для|дач\b|дом\b|круглый\sгод|этажн|сел(а|о)|посел|дом|до\b|вы|цена|или\b|деревни|напротив|участ|водопров|сеть|коммуникац|канализ)(\w|\s){3,})\s(?<type>(?i)\b(ул(и(ц(а|(?<=(по|на)\s\[А-Я]w+\sулиц)е)?)?)?|(?<=(ий|ый)\s)п(?=\w)((р(ос(пе|\.|\-)\.|\-)?|\-)?к?(?(\.)т?|т))|пер|ми?(кр)?\-?(орай)?о?н?\.?\))\b\.?)(\.|\,|\s|\))*(д(ом)?\.?\s*(?<house_num>\d(\d|\-|\\|\/){0,4})(?<letter>\w+)?)*($|\.|\""|\,|\s|\)|\()*");


        private readonly Regex streetToLeftRegex =
            new Regex(
                @"(?n)(^|\-|\.|\""|\,|\s|\)|\()+(?<!(?i)(рядом(\sс)?|окколо|двух)\s)(?<type>(?i)\b(ул(и(ц(а|(?<=(по|на|в)\sулиц)е|е(?=\s(?-i)[А-Яа-я])|ы)?)?|\.(?((?=\s?\w+[А-Яа-я])|(?!\w+\b\s?(&|\,|\.|\(|\))))\s?|\s)|\s\.(?=[А-Яа-я][А-Яа-я]+))?(?=\s?\.?\s*\w+)|п(?=\w)((р(ос(пе)?)?|)?\-?к?т)|пер(еулок|\.)?|м(?=\w)(и(?!н))?(кр)?\-?(орай)?о?(н([А-Яа-я]{0,2}))?|б(ульва|\-)р|тупик|(про|разъ|въ)езд(?!\s(авто|на\b|в\b|вдоль\b))|ш(оссе(?!\s\d)|\.)?|алл(ея|\.))\)?\b)(\s|\""|'|`)*(?<name>\b(?!(?i)коттеджн|имеет|Район|до\b|есть|зимой|от\b|круглый\sгод|за|\d+км|круглый\sгод|проход|как\b|так\b|по\b|при\b|ул\b|провед|км\b|без\b|ведут|газ\b|для|дач\b|дом\b|этажн|сел(а|о)|посел|дом|до\b|вы|цена|или\b|деревни|напротив|участ|водопров|сеть|коммуникац|канализ)([А-Яа-я][А-Яа-я]{0,3}\.[А-Яа-я][А-Яа-я]+|\d+(?!(\d*?\s?км|\s|$))((\sлет)|(\s?\-?\s?[А-Яа-я]+)?\s?\-?\s?[А-Яа-я]{0,})|[А-Яа-я]{2,}|([А-Яа-я]|\s?\-\s?){2,}\d+?)(?(\s[А-Яа-я]\w+\s?(\)|:|-|\.|\,|\""))\s[А-Яа-я]\w+))(\.|\,|:|\s|\)|\()*(((д(ом)?|уч)?\.?)\s*(?<house_num>(\d|\\|/){1,5}(?!(\s?(\.|\,)\s?\d+)?\s(сот|га)))(?(house_num)(\s?к(орп(ус))?\.?\s?(?<korp>\d+))?))?($|\.|\""|\,|\s|\)|\()+");


        private readonly Regex nearCityRegex =
            new Regex(
                @"(?n)(?:^|\""|\.|\,|\s|\))+\""?(?<name>(?#Слева не должны стоять)(?<!(напротив|20\d\d\s|\bдо\s|\bбуквой\b|(поворот\s)?на\s|\bза\s|\bул\.?\s|км\.?\sв\s|видно\s|Ч\.|ж.|т\.|\bжизнь\s+((с|в)\s+)?)\s*)(?!(?i)(?#Название насел пункта не может начинаться с следующих вариантов)((Прозер|\-|км\b|располож|категор|\d|балкон[а-я]*\b|коттедж|рядом|продает|есть|находит|продам\b|8|поселен|поля\b|под\b|\bэто\b|соврем|органич|Цен|По\b|При\b|Торг\b|район|р-н|газ\b|Уник|МКАД|недалеко|вид\b|(?<=\s)Чист|интернет|Улиц|территор|продаю|днп|Дорог|школа|гэс|уже\b|ст\b|земля|Дачный|перспективн|вода\b|Участ|очень|производ|поселк|ост\b|около|включает|Продаж|огорожен|СНТ\b|граничит|сделан|полдома|площадь|Жилой|охраняет|охрана|для\b|от\b|за\b|предприят|готов|напротив|днт|город|офис|сада|пмж|кафе|пос\b|продаю|фото|М\-?\d+|АЗС|развитая|СПБ|зу\b|шум\b|об\b|центр\b|бизнес|выход|газифиц|весь\b|места|срочно|ОАО|Участк(ами)?|Категори|Земли|Гаранти|документ|Звоните|Лес|На|между|Продаются|ул\b|пос\b|МО\b|Дом|Домом|Бл\b|мкр)[а-я]*|(?#слова не моугт оканчиватсья на)[А-Яа-я]*(?(?<=с\s\w+)(ым|им|ой|ей)|(ом|ого)))\b)(?!(?i)\b(по|пмж|сот|снт|недалеко)\b)(\b(?!\d)[А-Я][^\,\\\/\(\)\b\s]{2,}\b\s?){1,2})\)?\s+(?<type>(?i)(?:\bд(?:п)?\b|\b(по)?(?:(?<!(?i)рядом\s)(с|c)(?!\s\w+(ым|им|ом|ой|ей)))(?!\s\w+(ом|ем|им|ой|ей))(?:(е|ё)л[а-я]{1,2})?\b|\bп(?:гт|ос)?\b|\bч\b|\bнп\b|\bрп\b|\bх\b))(?:$|\.|\,|\""|\s|\))+");


        private readonly Regex nearCityToLeftRegex =
            new Regex(
                @"(?n)(^|\.|\""|\,|\)|\()?\s*(\s(?<in>в)\s)?(?<out>(?i)((в\s)?р(айоне?|-не)|в?близи?(\sот)?|\bу|возле|около|рядом\sс|(в\s)?(?<dist>\d(\d|\-|\,|\.|х)*\s?(к(ило)?)?м(етр[а-я]{0,3})?)\sот|от\b)\s)?(?<type>(?i)(?#слева НЕ должны быть следующие слова)(?<!(напротив|20\d\d\s|\bдо\s|\bбуквой\b|(поворот\s)?на\s|\bза\s|\bул\.?\s|км\.?\sв\s|видно\s|Ч\.|ж.|т\.|договор|\bжизнь\s+((с|в)\s+)?)\s*)(\b((?<!(\\|\/))д(\,)?(?!\s\d+)(\.|п|ер(евн[а-я]*)?)?|((?<!\bкот(т?еджн[а-я]{1,4})?\s)((дачн)(ый|ом|ого)\s)?по)?(?<!(съ|за)езд[а-я]?\s|с(\\|\/)|близост(ь|и)\s|рядом\s)(с|c)((\\|/)п)?(?!\s\w+(ом|им|ем|ой|ей|ми))((е|ё)л(?!ь?(ско)?хоз)?[а-я]{0,4}|(?((\.?\s+\w+(\s*\-\s*\w+)?\s*($|\""|\.|\,|\s|\)|\'|\())|(?<=в\sс\s))\.?|\.))\b|п(гт|ос|\.)?|ч(?!\.?п)|(?<out>при)?г(ор(од[а-я]{0,3}(?!\.))?)?|н(\.|асел(\.|енным)?\s?)?п(\.|унктом)|р\.?п\.?|(?<!(\d(\-|\s)?)|(c|с)(\\|/))х(?!\s?\d))\b))(?!\s?((\d{2,})|\d+(х|x)\d+))((?<=\b\w{1,4})\s?\.|\s|\""|\\?\&quot;|«|')*(?<name>(им\.?\s)?(\b(?!(?i)(?#Название насел пункта не может начинаться с следующих вариантов)((Прозер|\-|км\b|располож|категор|\d|балкон[а-я]*\b|коттедж|рядом|продает|есть|находит|продам\b|8|поселен|поля\b|под\b|\bэто\b|соврем|органич|Цен|По\b|При\b|Торг\b|район|р-н|газ\b|Уник|МКАД|недалеко|вид\b|(?<=\s)Чист|интернет|Улиц|возле|жилой|полдома|территор|продаю|назначен|днп|Дорог|школа|гэс|уже\b|ст\b|земля|Дачный|перспективн|вода\b|Участ|очень|производ|поселк|ост\b|около|включает|Продаж|огорожен|СНТ\b|граничит|сделан|площадь|охраняет|охрана|для\b|от\b|за\b|предприят|готов|напротив|днт|город|офис|сада\b|пмж|кафе|пос\b|продаю|фото|М\-?\d+|АЗС|развитая|СПБ|зу\b|шум\b|об\b|центр\b|бизнес|выход|газифиц|весь\b|места|срочно|ОАО|Участк(ами)?|Категори|Земли|Гаранти|документ|Звоните|Лес\b|На\b|между|Продаются|ул\b|пос\b|МО\b|Дом\b|Домом|Бл\b|мкр)[а-я]*|(?#слова не моугт оканчиватсья на)[А-Яа-я]*(?(?<=с\s\w+)(ым|им|ой|ей)|(ом|ого)))\b)(?!\d)(?(?-i)((?<=с\s)|\w+\s+[а-я]\w+)(?(?<!\-\s?)(?((?<=с(\.|ело)\s?)(\w+(ое)))\w|[А-Я])|\w)|\w)(?#символы исключения)[^\,\\\s\/\(\)\b\.\']+(?!\s*\b(?<!в\s(\k<type>)\s?\w+\s)(?i:снт(?!\s[А-Я])|также|сот\b|ул(?!\.\s*(\'|\`|\"")?[А-Я])|морем|об(?!\-?во)(\.|л(\.|аст[а-я]{1,2}(?!\.?(?(?<=\.)\s*|\s+)(\'|\`|\"")?[А-Я])|\s))?|ра?(\-|ай)о?на?(?!\s(\'|\`|\"")?[А-Я])|20(\d|\.|\,){2,})[а-я]*)\b(\s*\-)?\s*(\d+(?!\s?(км|сот|га))\s?)?){1,2})($|\\?\&quot;|\""|»|\.|!|\,|\s|\)|\'|\()+");

        //Всякие исключения что слово не может начинатсья с Обл, район, пример: "село Прокофье Района Русский"  захватит "Прокофьево", но не "Района"

        private readonly Regex regionRegex =
            new Regex(
                @"(?n)(^|\""|\.|\,|\s|\)|\()+(?<!\b(д(п)?|(по)?(с|c)(ел[а-я]{1,2})?(?!\s\w+\s(р(-н|айон)[а-я]{0,3}|г(\.|ород[а-я]{0,2})))|п(гт)?|ч|нп|рп|х)\b\s|\bснт\s)((?<pre>в)\s)?(?<name>(?!(?i)(?#Наименование не одно из следующих вариантов)\b((?i)(чистый|жило(й|м)|недалеко|местонахожд|хороший|Перспект|Новый|Прописк|данном|Зелены|Любой)[а-я]*|До|от|М\d+)\b)(\b(?(pre)[А-Яа-я]([А-Яа-я]|\s?\-\s?)+(ом|ем)|(?((\w|\-)+(ый|ий|ой|ом|ем))[А-Яа-я](\w|\-)+|(?(\w+(ого)\sр([а-я]|\-)+а)[А-Яа-я]|[А-Я])(\w|\s?\-\s?)+)\b){1,2}))\s(?<type>(?i)(\bр(\.|\-|айо)?о?н(?(pre)е|(а|е)?)|\b(?<!(20|19)\d\d\s?)г(\.|ород(?(pre)е))?\b))(?(?<!(ого\sр(\w|\-)+а)|(ий\sр(\w|\-)+н))($|\""|\.(\s|$|\,)|\,|\s|\)|\()+)");


        private readonly Regex regionToLeftRegex =
            new Regex(
                @"(?n)(^|\""|\.|\,|\s|\))+((?<pre>\bв)\s)?(?<type>(?i)(\bр(\.|\-|айо|)?о?н(?(pre)е)\b|\bг(\.|ород(?(pre)е))?\b))($|\,|\'|\""|\s)+(?<name>\b(?!(деревн))(\w|\s|\-|\.){2,}?)($|\.|\'|\""|\,|\s)(д(\.|ом)\s*(?<house_num>\d+)(?<letter>\w+)?)?($|\.|\,|\""|\s|\))+");

        private readonly Regex vgtToLeftRegex =
            new Regex(
                @"(?:^|\""|\.|\,|\s|\))+(?<type>(?i)\bр(\.|\-|айо|)?о?н\b)(?:$|\.|\,|\s)\s*\""?(?<name>\b(?:\w|\s|\.){2,})(?:$|\.|\,|\s)(?:$|\.|\,|\""|\s|\))+");

        //Кинули в улицу
        private readonly Regex microRrRegex =
            new Regex(
                @"(?:^|\.|\,|\s|\)|\()+(?<type>(?i)(?<!(от|между|\bи|рядом(\sс)?|у|около|недалеко(\sот)?)\s)(?:\bми?(?:кр)?\-?(?:орай)?о?н?\.?\b))(?:\.|\,|\s|\)|\()?\s*(?:\""|'|`|\&quot;)?(?<name>(?!(?:\b((?:(?#Само название не может начинатсья или оканчиватсья как...=>)Прозер|СХК|Уник|Чист|Улиц|Дорог|Участ|Постройки)[а-я]*|[а-я]*(ом|ой|ого)))\b)((?#Приставки, что нужно брать)\s*завода\s*)?(?:\b?[А-Я][^\,\\\/\(\)\b\s\.]+(?!(?#Проверка на слово справа)\s\b(морем|об\.?(л(\.|аст[а-я]{1,2}|\s))?|р(\-|ай)о?на?)\b)\b\s?|\s?№\d+){1,2})(?:$|\""|\.|\,|\s|\)|'|`|&quot;|\()+");


        private readonly Regex wordWithHeadLetteRegex =
            new Regex(
                @"(?n)(?n)\b(?<!(от)\s)(?!(?i)(земе?л|продаж))(?!(?#Слова исключения)(?i)участок|ИЖС|М\d+|юг\b|снт\b|село\b|пос\b|км\b|сад\b|ГК\b|до\b|база\b)[А-Я](\w|\-)+\b(?(\s[А-Я]\w+($|\s?\(|\)|""|'))\s\w+)(?=\s*($|\""|”|“|»|\.|\,|\s|\)|'|`|&quot;|\())");

//        private readonly Regex nameWordRegex =
//
//            new Regex(
//                @"(?<=(?:^|\,|\.|(\())\s*)(?<name>(\b(?(?<=\()[а-яА-Я]|[а-яА-Я])[^\,\\\/\(\)\b\s\.]{3,}\b\s?){1,2})(?=\s*(?(\1)\)|(?:\.|\,|\(|$)))");
        private readonly Regex subjRegEx =
            new Regex(
                @"(?n)((?<pre>(?i)(по|в|за))\s)?(?<name>\b[А-Яа-я\-]+\b)\s(автономная(?=\sобл)\s)?(?<type>(?i)(?<=(ая|ой)\s)(обл(\.|(аст(ь|и))))|(?<=(ий|ый)\s)край)");

        private readonly Regex subjToLeftRegex =
            new Regex(
                @"(?n)((?<pre>(?i)(по|в|за))\s)?(?<type>(?i)респ(\.|ублик(а|е|и)))(?(?<=\.)\s?|\s+)(?<name>\b[А-Яа-я\-]+\b)");

        private dynamic distValue;

        #endregion


        private void FormatClassification()
        {
// ReSharper disable once UnusedVariable
            Dictionary<string, string> checkedValues = new Dictionary<string, string>();

            Regex tmpRegex;
            //Replaces
            var cellr = worksheet.Cells[2, nearCityColumn, lastUsedRow, nearCityColumn];
            cellr.Value = cellr.Value.ToString().Replace("\"\"", "");
//            ReserveColumns();

            const int percentForThisMethod = 70;
            var everyEachStep = lastUsedRow/percentForThisMethod;
            var currStep = 0;

            DateTime per10 = new DateTime();
            DateTime per100 = new DateTime();
            DateTime per500 = new DateTime();
            DateTime per1000 = new DateTime();


            for (var row = 2; row <= lastUsedRow; row++)
            {
                if ((row - 2)%10 == 0)
                {
                    if (row != 2)
                        Console.WriteLine(@"10 объектов обработано за {0}", (DateTime.Now - per10).ToString("g"));
                    per10 = DateTime.Now;
                }
                if ((row - 2)%100 == 0)
                {
                    if (row != 2)
                        Console.WriteLine(@"100 объектов обработано за {0}", (DateTime.Now - per100).ToString("g"));
                    per100 = DateTime.Now;
                }
                if ((row - 2)%500 == 0)
                {
                    if (row != 2)
                        Console.WriteLine(@"500 объектов обработано за {0}", (DateTime.Now - per500).ToString("g"));
                    per500 = DateTime.Now;
                }
                if ((row - 2)%1000 == 0)
                {
                    if (row != 2)
                        Console.WriteLine(@"1000 объектов обработано за {0}", (DateTime.Now - per1000).ToString("g"));
                    per1000 = DateTime.Now;
                }


                #region Инициализация строки

                //ВВести принцип, перед работой с ячейкой мы её очищаем
                var subjCell = worksheet.Cells[row, subjColumn];
                var regionCell = worksheet.Cells[row, regionColumn];
                var settlementCell = worksheet.Cells[row, settlementColumn];
                var nearCityCell = worksheet.Cells[row, nearCityColumn];
                var vgtCell = worksheet.Cells[row, vgtColumn];
                var streetCell = worksheet.Cells[row, streetColumn];
                var typeOfNearCity = worksheet.Cells[row, typeOfNearCityColumn];
                var landmarkCell = worksheet.Cells[row, additionalInfoColumn];
                var typeOfStreetCell = worksheet.Cells[row, typeOfStreetColumn];
                var distToRegCenterCell = worksheet.Cells[row, distToRegCenterColumn];
                var distToNearCityCell = worksheet.Cells[row, distToNearCityColumn];
                var sntKpDnpCell = worksheet.Cells[row, sntKpDnpColumn];
                var inCityCell = worksheet.Cells[row, inCityColumn];
                var houseNumCell = worksheet.Cells[row, houseNumColumn];
                var letterCell = worksheet.Cells[row, letterColumn];

                var cellsFilled = false;

                //Выносим значения в память
                //и затираем ячейки
                string subjValue = subjCell.Value is string ? ReplaceYO(subjCell.Value.ToString()) : String.Empty;
                subjCell.Value = String.Empty;
                string regionValue = regionCell.Value is string
                    ? ReplaceYO(regionCell.Value.ToString())
                    : String.Empty;
                regionCell.Value = String.Empty;
                string nearCityValue = nearCityCell.Value is string
                    ? ReplaceYO(nearCityCell.Value.ToString())
                    : String.Empty;
                nearCityCell.Value = String.Empty;
                string vgtValue = vgtCell.Value is string ? ReplaceYO(vgtCell.Value.ToString()) : String.Empty;
                vgtCell.Value = String.Empty;
                string sourceLinkValue =  worksheet.Cells[row, sourceLinkColumn].Value is string
                    ? ReplaceYO( worksheet.Cells[row, sourceLinkColumn].Value.ToString())
                    : String.Empty;
                string landmarkValue = landmarkCell.Value is string
                    ? ReplaceYO(landmarkCell.Value.ToString())
                    : String.Empty;
                landmarkCell.Value = String.Empty;

                Match match;
                DataTable subjectTable = null;
                DataTable customTable = null;
                var regCenter = string.Empty;
                var regName = string.Empty;

                #endregion

                //
                //Ячейка субъект
                //
                //Берём субъект по источнику если такое возможно
                string subjcetName = subjectSourceWorksheet.GetSubjectBySourceLink(sourceLinkValue);

                if (string.IsNullOrEmpty(subjcetName))
                    subjcetName = oktmo.GetFullName(subjValue, OKTMOColumns.Subject);


                if (!String.IsNullOrEmpty(subjcetName))
                {
                    subjCell.Value = subjcetName;

                    //Выборка
                    if (oktmo.StringMatchInColumn(customTable, subjcetName,
                        OKTMOColumns.Subject))
                    {
                        customTable = oktmo.GetCustomDataTable(customTable,
                            new SearchParams(subjcetName, OKTMOColumns.Subject));
                        subjectTable = customTable.Copy();

                        //Get RegCenter
                        regCenter = oktmo.GetDefaultRegCenterFullName(subjcetName, ref regName);
                        if (string.IsNullOrEmpty(regCenter))
                            Console.WriteLine(@"Не найден региональный центр по субъекту {0}", subjcetName);
                    }
                }
                else
                {
                    subjCell.Value = subjValue;
                    subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                }



            #region ячейка Муниципальное образование

#if doRegion
                //
                //Ячейка Муниципальное обращзование
                //На наличие муниципального образования
                if (!string.IsNullOrEmpty(regionValue))
                {
                    //Удаляем дублируемуб инфомарцию о субъекте
                    if (!String.IsNullOrEmpty(subjValue))
                        regionValue = regionValue.Replace(subjValue, ", ");


                    //Ищем субъект для сравнение с проставленным
                    tmpRegex = subjRegEx;
                    match = tmpRegex.Match(regionValue);
                    if (!match.Success)
                    {
                        tmpRegex = subjToLeftRegex;
                        match = tmpRegex.Match(regionValue);
                    }
                    if (match.Success)
                    {
                        //Собственно это главное, зачем мы входили в это условие. Исключаем Субъект для дальнейшего облегчения поиска других типов
                        regionValue = tmpRegex.Replace(regionValue, ", ");
                        var fullName = oktmo.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value), OKTMOColumns.Subject);
                        if (!String.IsNullOrEmpty(fullName) &&
                            subjCell.Value != null &&
                            subjCell.Value.ToString()
                                .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                        {
                            rowsToDelete.Add(row);
                            subjCell.Value = fullName;
                            subjectTable = oktmo.GetCustomDataTable(new SearchParams(fullName, OKTMOColumns.Subject));
                            customTable = subjectTable.Copy();
                            continue;
                        }
                    }

                    TryFillRegion(row, ref regionValue, ref customTable, subjectTable,
                        ref cellsFilled);

                    //На наличие поселения
                    match = settlementRegex.Match(regionValue);
                    //Если есть совпадение и оно не на всю строку
                    if (match.Success)
                    {
                        var name = TryTemplateName(match.Groups["name"].Value);
                        var type = match.Groups["type"].Value;
                        type = type.IndexOf("п", StringComparison.OrdinalIgnoreCase) >= 0
                            ? "сельское поселение"
                            : "сельсовет";

                        var fullName = name + " " + type;
                        settlementCell.Value = fullName;

                        //В выборке уже имеется субъект и возможно Регион(или ВГТ)
                        if (oktmo.StringMatchInColumn(customTable, fullName, OKTMOColumns.Settlement))
                            customTable = oktmo.GetCustomDataTable(customTable,
                                new SearchParams(fullName, OKTMOColumns.Settlement));
                        else
                        {
                            settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            if (regionCell.Value != null)
                            {
                                regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                            else if (nearCityCell.Value != null) //bug ячейка ещё не проверена
                            {
                                nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                        }


                        regionValue = settlementRegex.Replace(regionValue, ", ");
                    }

                    //Поиск  товарищств
                    tmpRegex = sntRegex;
                    match = tmpRegex.Match(regionValue);
                    if (!match.Success)
                    {
                        tmpRegex = sntToLeftRegex;
                        match = tmpRegex.Match(regionValue);
                    }
                    if (match.Success)
                    {
                        var newName = TryTemplateName(match.Groups["name"].Value);
                        sntKpDnpCell.Value = sntKpDnpCell.Value == null ? newName : ", " + newName;
                        regionValue = tmpRegex.Replace(regionValue, ", ");
                    }


                    //На наличие населенного пункта и его типа
                    tmpRegex = nearCityRegex;
                    var matches = tmpRegex.Matches(regionValue);
                    bool switched = false;
                    if (matches.Count == 0)
                    {
                        tmpRegex = nearCityToLeftRegex;
                        matches = tmpRegex.Matches(regionValue);
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
                        if (oktmo.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                            customTable = oktmo.GetCustomDataTable(customTable,
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
                            if (oktmo.StringMatchInColumn(subjectTable, name, OKTMOColumns.NearCity))
                            {
                                //BUG поселение уже может быть окрашено в красный
                                var newTable = oktmo.GetCustomDataTable(subjectTable,
                                    new SearchParams(name, OKTMOColumns.NearCity));
                                //Обновляем тип по найденному нас пункту если возможно
                                if (newTable.Rows.Count == 1)
                                {
                                    string newType;
                                    try
                                    {
                                        newType =
                                            newTable.Rows.Cast<DataRow>().First()[typeOfNearCityColumn - 1].ToString();
                                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    }
                                    catch (InvalidOperationException e)
                                    {

                                        throw e;
                                    }

                                    if (typeOfNearCity.Value == null || typeOfNearCity.Value.ToString() != newType)
                                    {
                                        type = newType;
                                    }
                                    //bug если в таблице 1 запись, может уже записать всё?
                                }
                            }
                            else
                            {
                                nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                        }

                        nearCityCell.Value = name; //Пишем найденное наименование в нужную ячейку
                        typeOfNearCity.Value = type;

                        regionValue = tmpRegex.Replace(regionValue, ", ");
                    }

                    //Для улиц
                    TryFillStreet(row, ref regionValue);


                    //Имена собственные
                    TryFindProperName(row, ref regionValue, ref customTable, subjectTable, ref cellsFilled, regCenter,
                        regName);

                    //Ту информацию, что мы не смогли разобрать вписываем в отдельную ячейку
                    if (regionValue.Length > 2)
                        landmarkCell.Value += regionValue + ", ";
                }
#endif
#endregion

#region Населенный пункт

#if doNearCity
                //
                //Разбираем Населенный пункт
                //
                var value = nearCityValue;
                //Удаляем дублируемуб инфомарцию о субъекте
                if (!String.IsNullOrEmpty(subjValue))
                    value = value.Replace(subjValue, ", ");

                if (!String.IsNullOrEmpty(value))
                {

                    //Ищем субъект для сравнение с проставленным
                    tmpRegex = subjRegEx;
                    match = tmpRegex.Match(value);
                    if (!match.Success)
                    {
                        tmpRegex = subjToLeftRegex;
                        match = tmpRegex.Match(value);
                    }
                    if (match.Success)
                    {
                        //Собственно это главное, зачем мы входили в это условие. Исключаем Субъект для дальнейшего облегчения поиска других типов
                        regionValue = tmpRegex.Replace(regionValue, ", ");
                        var fullName = oktmo.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value), OKTMOColumns.Subject);
                        if (!String.IsNullOrEmpty(fullName) &&
                            subjCell.Value != null &&
                            subjCell.Value.ToString()
                                .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                        {
                            rowsToDelete.Add(row);
                            subjCell.Value = fullName;
                            subjectTable = oktmo.GetCustomDataTable(new SearchParams(fullName, OKTMOColumns.Subject));
                            customTable = subjectTable != null ? subjectTable.Copy() : oktmo.Table.Copy();
                            continue;
                        }
                    }

                    //Поиск муниципального образования
                    tmpRegex = regionRegex;
                    match = tmpRegex.Match(value); // "Дальнево р-н"
                    if (!match.Success)
                    {
                        tmpRegex = regionToLeftRegex;
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
                        var fullName = oktmo.GetFullName(type == "город" ? type + " " + name : name,
                            OKTMOColumns.Region);
                        if (!String.IsNullOrEmpty(fullName)) //This is REGION
                        {
                            if (oktmo.StringMatchInColumn(customTable, fullName,
                                OKTMOColumns.Region))
                            {
//                                if (!string.Equals(fullName, regCenter, StringComparison.OrdinalIgnoreCase))
//                                {
                                regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                                //Выборка
                                customTable = oktmo.GetCustomDataTable(customTable,
                                    new SearchParams(fullName, OKTMOColumns.Region));

//                                }
                            }
                            else
                            {
                                regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }

                            //В зависимости заполнен ли уже Регион, пишем извлеченное значение в ячейку Региона или ДопИнформации
                            if (regionCell.Value == null)
                                regionCell.Value = fullName;
                            else if (
                                !string.Equals(fullName, regionCell.Value.ToString(),
                                    StringComparison.OrdinalIgnoreCase))
                                //Ситуция когда при обработке столбца "Регион" мы уже нашли более менее подходящее к субъекту мун.образование
                                //И при обработке населн пункта (данный процесс) мы тоже нашли подходящее к субъекту мун.образование
                                landmarkCell.Value = fullName + landmarkCell.Value + ", ";

                        }
                            //------------Try Apeend to VGT-----------
                        else if (!TryFillVGT(row, ref name, ref customTable, ref cellsFilled))
                        {
                            fullName = name + " " + type;

                            //В зависимости заполнен ли уже Регион, пишем извлеченное значение в ячейку Региона или ДопИнформации
                            if (regionCell.Value == null)
                                regionCell.Value = fullName;
                            else if (fullName != regionCell.Value.ToString())
                                landmarkCell.Value = fullName + landmarkCell.Value + ", ";
                        }
                        value = tmpRegex.Replace(value, ", ");
                    }

                    //Поиск киллометров до населенного пункта
                    match = distToRegCenteRegex.Match(value);
                    if (match.Success)
                    {

                        //Спихиваем всё в столбец "Расстояние до рег центра"
                        //Разбирать будем в конце
                        distToRegCenterCell.Value += ", " + match.Value;
                        value = distToRegCenteRegex.Replace(value, ", ");
                    }

                    //Поиск улиц
                    TryFillStreet(row, ref value);

                    //Поиск поселения
                    match = settlementRegex.Match(value);
                    if (match.Success)
                    {
                        string name = TryTemplateName(match.Groups["name"].Value);
                        string type = match.Groups["type"].Value;

                        if (type.IndexOf("п", StringComparison.OrdinalIgnoreCase) >= 0)
                            type = "сельское поселение";
                        else
                            type = "сельсовет";

                        var fullName = name + " " + type;
                        if (settlementCell.Value == null)
                            settlementCell.Value = fullName;
                        else
                            landmarkCell.Value += fullName + ", ";

                        if (oktmo.StringMatchInColumn(customTable, fullName, OKTMOColumns.Settlement))
                        {
                            customTable = oktmo.GetCustomDataTable(customTable,
                                new SearchParams(fullName, OKTMOColumns.Settlement));

                        }
                        else
                        {
                            settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            if (regionCell.Value != null)
                            {
                                regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                            else if (nearCityCell.Value != null)
                            {
                                nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                        }
                        value = settlementRegex.Replace(value, ",");
                    }

                    //Поиск 100% дополнительной инфомрации (снт, сот, с/н)
                    tmpRegex = sntToLeftRegex;
                    match = tmpRegex.Match(value); //"пурум снт"
                    if (!match.Success)
                    {
                        //"снт Пурум"
                        tmpRegex = sntRegex;
                        match = tmpRegex.Match(value);
                    }
                    if (match.Success)
                    {
                        while (match.Success)
                        {
                            string name = TryTemplateName(match.Groups["name"].Value);
                            string type = match.Groups["type"].Value;

                            sntKpDnpCell.Value = sntKpDnpCell.Value == null ? name : ", " + name;

                            match = match.NextMatch();
                        }
                        value = tmpRegex.Replace(value, ",");
                    }


                    //Поиск населенного пункта
                    tmpRegex = nearCityToLeftRegex;
                    var matches = tmpRegex.Matches(value); // "Дальнево с."
                    bool switched = false;
                    if (matches.Count == 0)
                    {
                        tmpRegex = nearCityRegex;
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
                            if (nearCityCell.Value == null || nearCityCell.Value.ToString() == String.Empty)
                            {

                                nearCityCell.Value = name;
                                typeOfNearCity.Value = type;

                                if (oktmo.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                                {
                                    customTable = oktmo.GetCustomDataTable(customTable,
                                        new SearchParams(name, OKTMOColumns.NearCity));

                                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    
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
                                        else if (name.Contains(" "))
                                        {
                                            name = name.Replace(" ", "-");
                                            goto tryAgainNCInNC;
                                        }
                                    }
                                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                }
                            }
                            else if (nearCityCell.Value.ToString() != name) //нашли ли мы новую информацию
                            {
                                if (oktmo.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                                    //и подходит ли она к нам
                                {
                                    customTable = oktmo.GetCustomDataTable(customTable,
                                        new SearchParams(name, OKTMOColumns.NearCity));

                                    landmarkCell.Value += nearCityCell.Value + ", ";

                                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    
                                    nearCityCell.Value = name;
                                    typeOfNearCity.Value = type;
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
                                    landmarkCell.Value += name + " " + type + ", ";
                                }
                            }
                        }
                        value = tmpRegex.Replace(value, ", ");
                    }
                        //Обрабатываем имена собственные
                    else
                    {
                        TryFindProperName(row, ref value, ref customTable, subjectTable, ref cellsFilled, regCenter,
                            regName);
                    }

                    nearCityValue = value;
                    //Если у нас что-то не разобрано, мы его пихаем в доп инфо или ту же ячейек
                    if (nearCityValue.Length > 2)
                    {
                        //Как бы зачем оставлять "3б" в населенном пункте
                        //В зависимости от была ли внесена полезная инфомация в ячеку "населенный пункт"
                        landmarkCell.Value += nearCityValue + ", ";
                    }
                    //Если у нас разобрано всё, а в ячейку населенного пункта ничего записано не было
                    //Мы очищаем ячейку
                }




                if (!String.IsNullOrEmpty(vgtValue))
                {
                    var tmpMatch = wordWithHeadLetteRegex.Match(vgtValue);
                    if (tmpMatch.Success)
                    {
                        var tmpValue = tmpMatch.Value;
                        TryFillVGT(row, ref tmpValue, ref customTable, ref cellsFilled);
                        value = value.Replace(tmpValue, ", ");
                    }
                }
                #region Доискиваем недостающую инфомрацию в столбце Ориентир (который вводится в выгрузек от октября 14 года)

                if (!String.IsNullOrEmpty(landmarkValue))
                {
                    //Поиск мун образвания
                    if (regionCell.Value == null)
                        TryFillRegion(row, ref landmarkValue, ref customTable, subjectTable,
                            ref cellsFilled);
                    //поиск улицы
                    if (streetCell.Value == null)
                        TryFillStreet(row, ref landmarkValue);
                    //Поиск внутрегородской территории
                    //Bug а не происходит ли такая же процедура в методе поиска мунОбразования (см 6 строк выше)
                    if (vgtCell.Value == null)
                    {
                        var tmpMatch = regionRegex.Match(landmarkValue);
                        if (tmpMatch.Success)
                        {
                            var tmpValue = tmpMatch.Groups["name"].Value;
                            TryFillVGT(row, ref tmpValue, ref customTable, ref cellsFilled);
                            landmarkValue = landmarkValue.Replace(tmpValue, ", ");
                        }

                    }

                    if (nearCityCell.Value == null)
                    {
                        //TODO недоделано
                    }


                    //обработка имен собственных
                    TryFindProperName(row, ref value, ref customTable, subjectTable, ref cellsFilled, regCenter, regName);

                    if (landmarkValue.Length > 2)
                        landmarkCell.Value += landmarkValue + ", ";
                }
#endregion
#endif
#endregion
#region Разбираем Описание на предмет Местоположения

#if doDescriptionPlace

                var descriptionColumn = GetColumnIndex("DESCRIPTION");

                //Вначале мы ищем наименования по типу
                //После мы пытаемся отнести найдненные в описании Именования без типов
                var cell = worksheet.Cells[row, descriptionColumn];
                if (cell.Value != null)
                {
                    String descrtContent = ReplaceYO(cell.Value.ToString());

                    //
                    //----Товарищества
                    //

                    match = sntToLeftRegex.Match(descrtContent);
                    if (match.Success)
                    {
                        do
                        {
                            //Берём только первое совпадение!
                            var name = TryTemplateName(match.Groups["name"].Value);

                            sntKpDnpCell.Value = sntKpDnpCell.Value == null ||
                                                  sntKpDnpCell.Value.ToString().Length < 3
                                ? name
                                : ", " + name;
                            descrtContent = sntToLeftRegex.Replace(descrtContent, ", ");
                            match = match.NextMatch();
                        } while (match.Success);
                    }

                    TryFillStreet(row, ref descrtContent);

                    //
                    //---Субъект для сравнение с проставленным
                    //
                    tmpRegex = subjRegEx;
                    match = tmpRegex.Match(descrtContent);
                    if (!match.Success)
                    {
                        tmpRegex = subjToLeftRegex;
                        match = tmpRegex.Match(descrtContent);
                    }
                    if (match.Success)
                    {
                        //Собственно это главное, зачем мы входили в это условие. Исключаем Субъект для дальнейшего облегчения поиска других типов
                        regionValue = tmpRegex.Replace(regionValue, ", ");
                        var fullName = oktmo.GetFullName(TryChangeSubjectEndness(match.Groups["name"].Value), OKTMOColumns.Subject);
                        if (!String.IsNullOrEmpty(fullName) &&
                            subjCell.Value != null &&
                            !string.Equals(subjCell.Value.ToString().Trim(),fullName.Trim(),StringComparison.OrdinalIgnoreCase))
                            //subjCell.Value.ToString()
                            //    .IndexOf(match.Groups["name"].Value, StringComparison.OrdinalIgnoreCase) == -1)
                        {
                            rowsToDelete.Add(row);
                            subjCell.Value = fullName;
                            subjectTable = oktmo.GetCustomDataTable(new SearchParams(fullName, OKTMOColumns.Subject));
                            customTable = subjectTable != null ? subjectTable.Copy() : oktmo.Table.Copy();
                            continue;
                        }
                    }

                    //
                    //----Населенный пункт
                    //
                    bool switched = false;
                    bool endChanged = false;
                    var regs = new List<Regex> {nearCityToLeftRegex, nearCityRegex};
                    Regex reg;
                    foreach (Regex regi in regs)
                    {
                        reg = regi;

                        var matches = reg.Matches(descrtContent);

                        if (matches.Count > 0)
                        {
                            match = null;
                            //Приоритет у любого негорода
                            if (matches.Count > 1)
                            {
                                //Приорите у любого негорода без рассстояния
                                match =
                                    matches.Cast<Match>()
                                        .FirstOrDefault(
                                            m =>
                                                (m.Groups["type"].Value.IndexOf("г", StringComparison.OrdinalIgnoreCase) ==
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



                            var name = ReplaceYO(TryTemplateName(match.Groups["name"].Value));
                            var type = ReplaceYO(TryDescriptTypeOfNasPunkt(match.Groups["type"].Value));


                            if (!string.IsNullOrEmpty(match.Groups["out"].Value))
                            {
                                inCityCell.Value = "нет";
                                if (!string.IsNullOrEmpty(match.Groups["dist"].Value))
                                {
                                    var dist = TryDescriptDistance(match.Groups["dist"].Value);
                                    if (string.Equals(name, regName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        //Backup current value
                                        if (distToNearCityCell.Value != null)
                                            landmarkCell.Value +=
                                                String.Format("Расстояние до регионального центра \"{0}\"",
                                                    distToRegCenterCell.Value);
                                        distToRegCenterCell.Value = dist;
                                    }
                                    else
                                        distToNearCityCell.Value = dist;
                                }
                            }

                            bool splitted = false;
                            List<string> words = null;
                            var startName = name;
                            tryGetNearCityAgain:

                            //Опеределяем нужно ли обрабатывать найденную информацию
                            if ((string.Equals(name, regName, StringComparison.OrdinalIgnoreCase) &&
                                 (regionCell.Style.Fill.BackgroundColor.Rgb !=
                                  ExcelExtensions.BadColor.ToArgb().ToString() ||
                                  !string.Equals(name, regName, StringComparison.OrdinalIgnoreCase))) ||
                                (nearCityCell.Value != null && (string.Equals(nearCityCell.Value.ToString(), name,
                                    StringComparison.OrdinalIgnoreCase)))) continue;

                            if (type == "город" && typeOfNearCity.Value != null &&
                                typeOfNearCity.Value.ToString() != "город")
                            {
                                landmarkCell.Value += name + " " + type + ", ";
                            }
                            else
                            {
                                //BackUp current value
                                if (nearCityCell.Value != null)
                                    landmarkCell.Value += typeOfNearCity.Value + " " +
                                                          nearCityCell.Value.ToString();

                                //Обнуляем МунОбразование
                                //сейчас стоит региональный центр или просто город
                                //а найденный насел пункт подходит к другому мун образованию
                                var itIsCity = (regionCell.Value != null &&
                                                String.Equals(regionCell.Value.ToString(), regCenter,
                                                    StringComparison.OrdinalIgnoreCase) ||
                                                (regionCell.Value != null &&
                                                 regionCell.Value.ToString()
                                                     .IndexOf("город", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                                 type != "город") ||
                                                (nearCityCell.Value != null && typeOfNearCity.Value == null));

                                var valueNeedsResetRegion =
                                    !oktmo.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity) &&
                                    oktmo.StringMatchInColumn(subjectTable, name, OKTMOColumns.NearCity);

                                if (itIsCity && valueNeedsResetRegion)
                                {
                                    customTable = subjectTable != null ? subjectTable.Copy() : oktmo.Table.Copy();
                                    regionCell.Value = string.Empty;
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    settlementCell.Value = string.Empty;
                                    settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                }

                                const string dashPattern = @"\s*\-\s*";
                                const string spacePattern = @"\s+";
                                const string cityEnd = @"(е|а)\b";
                                //найденный насел пункт подхоидт к нашей выборке (по субъетк и возможно по мунобразованию если оно есть)
                                if (oktmo.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                                    customTable = oktmo.GetCustomDataTable(customTable,
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
                                            var patterns = new List<string> {dashPattern, spacePattern};

                                            foreach (string pattern in patterns)
                                            {
                                                if (Regex.IsMatch(startName, pattern))
                                                {
                                                    words = Regex.Split(startName, pattern).ToList();
                                                    name = words.Last();
                                                    words[words.Count - 1] = null;
                                                    goto tryGetNearCityAgain; //just break
                                                }
                                            }
                                        }
                                        //Step two: we use it untill end
                                        else
                                        {
                                            for (int i = words.Count - 1; i >= 0; i--)
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

                                //записываем в любом случае
                                nearCityCell.Value = name;
                                typeOfNearCity.Value = type;

                                cellsFilled = false;
                            }

                            descrtContent = reg.Replace(descrtContent, ", ");
                        }
                    }
                    //
                    //----Муниципальное образование
                    //
                    regs = new List<Regex> {regionRegex, regionToLeftRegex};
                    foreach (Regex regi in regs)
                    {

                        if (!cellsFilled)
                        {
                            TryFillRegion(row, ref descrtContent, ref customTable, subjectTable,
                                ref cellsFilled, regi);

                        }
                    }

                    //=================
                    //Коммуникацияя
                    //=================
                }
#endif

                #endregion

                //
//              //Подтираем столбец Ориентир
                //
                const string del = @"(\,\s+)";
                var r = new Regex(del + @"{2,}", RegexOptions.IgnorePatternWhitespace);
                if (landmarkCell.Value != null)
                {
                    landmarkCell.Value = r.Replace(landmarkCell.Value.ToString(), ", ");
                }
#if fillDefault

                //Вписываем дефолтные значения Если населенный пункт так и не заполнен
                if (nearCityCell.Value == null)
                {
                    //Находим дефолтный населенный пункт по ссылке на объявление
                    var newCity = subjectSourceWorksheet.GetDefaultNearCityByLink(sourceLinkValue);

                    if (!String.IsNullOrEmpty(newCity))
                    {
                        //Мы пишем насел пункт только если он подходит к нашей выборке
                        //Т.е. подходит и к субъекту и к муниципальному образованию, есть таковой есть
                        if (oktmo.StringMatchInColumn(customTable, newCity, OKTMOColumns.NearCity))
                        {
                            nearCityCell.Value = newCity;
                            typeOfNearCity.Value = "город";

                            customTable = oktmo.GetCustomDataTable(customTable,
                                new SearchParams(newCity, OKTMOColumns.NearCity));
                            TryFillClassificator(row, ref customTable, ref cellsFilled, regCenter, regName);
                        }
                    }
                        //или ставим муниципальное образование как город
                        //При условии что это не региональный центр
                    else if (regionCell.Value != null && regionCell.Style.Fill.BackgroundColor.Rgb != ExcelExtensions.BadColor.ToArgb().ToString()
                             && regionCell.Value.ToString().IndexOf("город") >= 0)
                    {
                        string name = regionCell.Value.ToString().Replace("город", "");
                        name = name.Replace("(ЗАТО)", "");
                        name = name.Trim();
                        if (oktmo.StringMatchInColumn(customTable, name, OKTMOColumns.NearCity))
                        {
                            cellsFilled = false;
                            nearCityCell.Value = name;
                            typeOfNearCity.Value = "город";

                            customTable = oktmo.GetCustomDataTable(customTable,
                                new SearchParams(name, OKTMOColumns.NearCity));
                            TryFillClassificator(row, ref customTable, ref cellsFilled, regCenter, regName);
                        }
                    }
                }
                //Ставим дефолтное значение для муниципального образования, если оно пустое, а текущий насленный пункт у нас является региональным центро
                else if (regionCell.Value == null &&
                         string.Equals(nearCityCell.Value.ToString(), regName, StringComparison.OrdinalIgnoreCase))
                {
                    customTable = oktmo.GetCustomDataTable(customTable,
                        new SearchParams(regName, OKTMOColumns.NearCity));
                    TryFillClassificator(row, ref customTable, ref cellsFilled, regCenter, regName);
                }
                //Дефолное значение для типа населенного пункта, если найденный насел пункт совпадает по названию с региональным центром
                else if (typeOfNearCity.Value == null && 
                         string.Equals(nearCityCell.Value.ToString(), regName, StringComparison.OrdinalIgnoreCase))
                {
                    typeOfNearCity.Value = "город";
                }
#endif
#if checkVGT
                //Проверяем текущий ВГТ
                //Дописываем по возможности населенный пункт, опираясь на Муниципальное образование (если оно таки есть)
                //Так же удаляем если в справочнике ВГТ нет комбинации НаселПункт+ВГТ
                if (vgtCell.Value != null)
                {
                    //т.к у нас есть инфомрация о ВГТ, мы можем записать Регион(город) вручную или наоборот, записать насел пункт по региону
                    //но нам не надо будет этого делать, если насел пункт валиден относительно ВГТ, т.к. Регион проставится сам по населенному пункту
                    if (nearCityCell.Value == null ||
                        !vgtWorksheet.CombinationExists(nearCityCell.Value.ToString(), vgtCell.Value.ToString()))
                    {
                        //Список городов в которых присутствует район, с таким же наименованием
                        List<string> cities = vgtWorksheet.GetCitiesListByTerritory(vgtCell.Value.ToString());

                        //Когда населенный пункт пустой, а муниципальное образовать есть
                        if (subjectTable != null && regionCell.Value != null && nearCityCell.Value == null)
                        {
                            //Среди всех городов, в которых есть текущий ВГТ
                            //Находим те, которые подходят к нашей текущей выборке
                            var validCities =
                                cities.Where(city => oktmo.StringMatchInColumn(customTable, city, OKTMOColumns.NearCity))
                                    .ToList();

                            //И если подходит всего один город
                            //Мы ставим его как населенный пункт
                            if (validCities.Count == 1)
                            {
                                var newCity = validCities.First();
                                nearCityCell.Value = newCity;
                                typeOfNearCity.Value = "город";
                                //и дописываем выборку
                                customTable = oktmo.GetCustomDataTable(customTable,
                                    new SearchParams(newCity, OKTMOColumns.NearCity));
                                TryFillClassificator(row,ref customTable,ref cellsFilled,regCenter,regName);
                            }
                        }
                        //Обрабатываем населенный пункы = не города
                        if (typeOfNearCity.Value == null ||
                            !string.Equals(typeOfNearCity.Value.ToString(), "город", StringComparison.OrdinalIgnoreCase))
                        {
                            //При пустом Регионе смело пишем ВГТ, т.к. проставленный насел пункт может находиться в районе города, и не быть в ВГТ и ОКТМО справочниках
                            if (regionCell.Value == null)
                            {
                                //Если текущий ВГТ находится и в региональном центре, то смело пишем региональный центр как МунОбр
                                if (!string.IsNullOrEmpty(regCenter) && cities.Contains(regName))
                                {
                                    regionCell.Value = regCenter;
                                    customTable = oktmo.GetCustomDataTable(customTable,
                                            new SearchParams(regCenter, OKTMOColumns.NearCity));
                                    if (nearCityCell.Value == null)
                                    {
                                        nearCityCell.Value = regName;
                                        typeOfNearCity.Value = "город";
                                        customTable = oktmo.GetCustomDataTable(customTable,
                                            new SearchParams(regName, OKTMOColumns.NearCity));
                                    }
                                    TryFillClassificator(row,ref customTable,ref cellsFilled,regCenter,regName);
                                }
                                else if (subjectTable != null) //А вообще априори всегда заполнено
                                {
                                    //Ищем дугие пересечения между списокм городов субъекта и списком городов в текущим районом ВГТ
                                    //Список городов по нашему субъекту
                                    var subjCitiesRows =
                                        subjectTable.Rows.Cast<DataRow>()
                                            .Where(
                                                rw =>
                                                    string.Equals(
                                                        rw[OKTMOWorksheet.Columns.TypeOfNearCity - 1].ToString(),
                                                        "город"))
                                            .ToList();

                                    //Находим города, что если в обоих списках
                                    var same =
                                        subjCitiesRows.Where(
                                            rw => cities.Contains(rw[OKTMOWorksheet.Columns.NearCity - 1].ToString()))
                                            .ToList();
                                    if (same.Count() == 1)
                                    {
                                        //Bug или лучше найти МунОбр через населенный пункт?
                                        var newReg = same[0][OKTMOWorksheet.Columns.Region - 1].ToString();

                                        regionCell.Value = newReg;
                                    }
                                }
                            }
                            else
                            {
                                AppendToLandMarkCell(vgtCell.Value.ToString() + " район", row);
                                vgtCell.Value = string.Empty;

                            }
                        }
                            //Если стоит ВГТ, не относящийся в проставленному городу
                            //Bug может ли такое быть, и правильно ли так делать?
                        else
                        {
                            AppendToLandMarkCell(vgtCell.Value.ToString() + " район", row);
                            vgtCell.Value = string.Empty;
                        }
                    }
                }
                TryFillClassificator(row, ref customTable, ref cellsFilled, regCenter, regName);
#endif
                //
                //Прогресс бар
                //
                currStep++; //Инкрементируем групповой счётчик для единицы статус бара
                if (currStep == everyEachStep) //Если мы обработали строк на единицу статусбара
                {
                    progressBar.BeginInvoke(new VoidDelegate(() => progressBar.PerformStep()));
                        //Инкрементируем прогрессбар
                    currStep = 0; //Сбрасываем счётчик
                }

            }
            progressBar.Invoke(new Action(() => progressBar.Value = percentForThisMethod));
        }

        /// <summary>
        /// Вставить текст в ячейку "Ориентир" по указанной строке
        /// </summary>
        /// <param name="value">Текст для вставки</param>
        /// <param name="row">Строка для поиска</param>
        private void AppendToLandMarkCell(string value, int row)
        {
            var landmarkCell = worksheet.Cells[row, additionalInfoColumn];

            if (landmarkCell.Value != null &&
                landmarkCell.Value.ToString().IndexOf(value, System.StringComparison.Ordinal) >= 0)
                return;


            if (value.IndexOf("район", System.StringComparison.Ordinal) >= 0)
                if (landmarkCell.Value == null)
                    landmarkCell.Value = value + ", ";
                else
                    landmarkCell.Value = value + ", " + landmarkCell.Value.ToString();
            else
                landmarkCell.Value += value + ", ";

        }

        /// <summary>
        /// Иетод возвращает расшифрованную дистанцию
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int TryDescriptDistance(string value)
        {
            const string integer = @"\d(\d|\.|\,)*";
            var match = Regex.Match(value, integer);
            if (!match.Success) return 0; //Хотя вообще такого случаться в принципе не должно

            var result = 0;
            int.TryParse(match.Value, out result);

            if (value.IndexOf("к", StringComparison.OrdinalIgnoreCase) == -1)
                result = result*1000;

            return result;
        }

        /// <summary>
        /// Метод пытается найти Имена собственные в переданной строке, и пытается их опеределить к какому-либо тиипу (мунОбр, населПункт, ВГТ и прочие)
        /// </summary>
        /// <param name="row"></param>
        /// <param name="value"></param>
        /// <param name="customTable"></param>
        /// <param name="subjectTable"></param>
        /// <param name="cellsFilled"></param>
        /// <param name="regCenter"></param>
        /// <param name="regName"></param>
        private void TryFindProperName(int row, ref string value, ref DataTable customTable, DataTable subjectTable,
            ref bool cellsFilled, string regCenter, string regName)
        {
            if (customTable == null) return;
            if (subjectTable == null) return;


            var regionCell = worksheet.Cells[row, regionColumn];
            var nearCityCell = worksheet.Cells[row, nearCityColumn];
            var streetCell = worksheet.Cells[row, streetColumn];
            var typeOfStreetCell = worksheet.Cells[row, typeOfStreetColumn];
            var sntKpDnpCell = worksheet.Cells[row, sntKpDnpColumn];
            var subjCell = worksheet.Cells[row, subjColumn];

            var match = wordWithHeadLetteRegex.Match(value);
            while (match.Success)
            {
                //does not match region and near city
                //and does not match SNT (or it`s just empty)
                if (match.Value != (string) regionCell.Value && match.Value != (string) nearCityCell.Value &&
                    (sntKpDnpCell.Value == null ||
                     (((string) sntKpDnpCell.Value.ToString()).IndexOf(match.Value,
                         StringComparison.OrdinalIgnoreCase) == -1)))
                {
                    //Пробуем подогнать к каждой ячейке
                    //Если никуда не подошло то пишем в первую пустую

                    //Try append to Region
                    var fullName = OKTMOWorksheet.GetFullName(subjectTable, "город" + " " + match.Value,
                        OKTMOColumns.Region); //Tty to find on whole OKTMO
                    if (!String.IsNullOrEmpty(fullName))
                    {
                        if (!cellsFilled)
                        {
                            //Найденный регион пишем только если он подходит к выборке
                            if (oktmo.StringMatchInColumn(customTable, fullName, OKTMOColumns.Region))
                            {
                                regionCell.Value = fullName;
                                regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                                //Делаем выборку только если найденный регион не является региональным центром
                                if (!string.Equals(fullName, regCenter, StringComparison.OrdinalIgnoreCase))
                                    customTable = oktmo.GetCustomDataTable(customTable,
                                        new SearchParams(fullName, OKTMOColumns.Region));
                            }
                        }
                    }
                        //Try append to NearCity
                    else if (oktmo.StringMatchInColumn(customTable, TryTemplateName(match.Value),
                        OKTMOColumns.NearCity))
                    {
                        if (!cellsFilled)
                        {
                            var newName = TryTemplateName(match.Value);
                            nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            
                            nearCityCell.Value = newName;

                            if (!string.Equals(newName, regName, StringComparison.OrdinalIgnoreCase))
                                customTable = oktmo.GetCustomDataTable(customTable,
                                    new SearchParams(newName, OKTMOColumns.NearCity));
                        }
                    }
                        //Try Append To VGT
                    else if (vgtWorksheet.TerritotyExists(match.Value))
                    {
                        var v = TryTemplateName(match.Value);
                        TryFillVGT(row, ref v, ref customTable, ref cellsFilled);
                    }
                        //Just Wtire to first epmty cell
                    else
                    {
                        if (streetCell.Value == null &&
                            Regex.IsMatch(match.Value, @"ая\b", RegexOptions.IgnoreCase))
                        {
                            streetCell.Value = TryTemplateName(match.Value);
                            typeOfStreetCell.Value = "улица";
                        }
                        else if (nearCityCell.Value == null)
                        {
                            nearCityCell.Value = TryTemplateName(match.Value);
                            nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            if (regionCell.Value != null)
                            {
                                regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                            else
                            {
                                subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            }
                        }
                        else
                            goto skipWordReplace;
                    }

                    value = wordWithHeadLetteRegex.Replace(value, ", ");
                }
                skipWordReplace:
                match = match.NextMatch();
            }
        }

        private bool TryFillVGT(int row, ref string value, ref DataTable customTable, ref bool cellsFilled)
        {

            var nearCityCell = worksheet.Cells[row, nearCityColumn];
            var vgtCell = worksheet.Cells[row, vgtColumn];
            var landmarkCell = worksheet.Cells[row, additionalInfoColumn];
            var regionCell = worksheet.Cells[row, regionColumn];
            var typeOfNearCity = worksheet.Cells[row, typeOfNearCityColumn];
            var res = false;
            //----Обрабатываем ВГТ-----
            if (!String.IsNullOrEmpty(value))
            {
                //Подтверждаем, что это ВГТ
                if (vgtWorksheet.TerritotyExists(value))
                {
                    var vgt = value;
                    vgtCell.Value = vgt;
                    res = true;

                    if (nearCityCell.Value != null &&
                        vgtWorksheet.CombinationExists(nearCityCell.Value.ToString(), vgt))
                        return true;

                    //Далее идут ситации если текущий насел пункт пустой, или не подходит к найденному ВГТ

                    //Пробуем определить населенный пункт
                    String city = string.Empty;
                    //Пробуем извлечь текущий насел пункт из мунОбр
                    //И тем самым подтвердить мунОбр и проставить населПункт
                    if (regionCell.Value != null &&
                        regionCell.Value.ToString().IndexOf("город", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        city = TryTemplateName(regionCell.Value.ToString().Replace("город", ""));
                        city = city.Trim();
                    }
                    if (!String.IsNullOrEmpty(city) && vgtWorksheet.CombinationExists(city, vgt))
                    {
                        nearCityCell.Value = city;
                        typeOfNearCity.Value = "город";

                        //Проверяем найденный насел пункт
                        if (oktmo.StringMatchInColumn(customTable, city, OKTMOColumns.NearCity))
                        {
                            nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            

                            customTable = oktmo.GetCustomDataTable(customTable,
                                new SearchParams(city, OKTMOColumns.NearCity));
//                            TryFillClassificator(row, customTable, ref cellsFilled);
                        }
                        else
                        {
                            nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        }
                    }
                        //В ином случае пробуем записать насел пункт через ВГТ
                    else
                    {
                        string newCity = cellsFilled
                            ? String.Empty
                            : vgtWorksheet.GetCityByTerritory(vgt);
                        if (!String.IsNullOrEmpty(newCity))
                            //Строка будет  заполнена, если существует всего один насел пункт с таким районом
                        {
                            //нужно ли нам вообще проверять найденный
                            if (nearCityCell.Value != null &&
                                string.Equals(nearCityCell.Value.ToString(), newCity,
                                    StringComparison.CurrentCultureIgnoreCase)) return res;

                            //Если текущий населенный пункт верный (он не пуст и не окрашен как неверный)
                            //мы его оставляем на месте, а найденный пишем в ориентир
                            if (nearCityCell.Value != null &&
                                nearCityCell.Style.Fill.BackgroundColor.Rgb != ExcelExtensions.BadColor.ToArgb().ToString())
                                //Пишем найденный насел пункт в ориентир
                                landmarkCell.Value += "город " + vgt + ", ";

                            //В остальных случаях найденный насел пункт попадёт в ячейку населенногоп пункта
                            else
                            {
                                //Определяем, относится ли насел пункт к выборке
                                if (oktmo.StringMatchInColumn(customTable, newCity,
                                    OKTMOColumns.TypeOfNearCity))
                                {
                                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                                        

                                    //Try to fill
                                    customTable = oktmo.GetCustomDataTable(customTable,
                                        new SearchParams(newCity, OKTMOColumns.NearCity));
                                }
                                else
                                {
                                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                }

                                //Перекидываем текущий насел пункт
                                if (nearCityCell.Value != null)
                                    landmarkCell.Value += nearCityCell + ", ";

                                nearCityCell.Value = newCity;
                            }
                        }
                    }
                }
            }
            return res;
        }

        private void TryFillStreet(int row, ref string value)
        {
            var streetCell = worksheet.Cells[row, streetColumn];
            var landmarkCell = worksheet.Cells[row, additionalInfoColumn];
            var typeOfStreetCell = worksheet.Cells[row, typeOfStreetColumn];
            var houseNumCell = worksheet.Cells[row, houseNumColumn];
            var letterCell = worksheet.Cells[row, letterColumn];
            var buildCell = worksheet.Cells[row, buildColumn];


            //Поиск улиц
            List<Regex> regs = new List<Regex> {streetToLeftRegex, streetRegex};
            foreach (Regex reg in regs)
            {
                var match = reg.Match(value);
                if (match.Success)
                {
                    //По сути если у нас уже проставлена улица, новую нужно игнорировать кроме нескольких случаев ниже

                    //Берём только первое совпадение!
                    var name = ReplaceYO(TryTemplateName(match.Groups["name"].Value));
                    var type = ReplaceYO(TryDescriptTypeOfStreet(match.Groups["type"].Value));

                    if (streetCell.Value == null || streetCell.Value.ToString() == String.Empty ||
                        streetCell.Value.ToString() != name &&
                        (typeOfStreetCell.Value == null ||
                         typeOfStreetCell.Value.ToString().ToLower() == "микрорайон".ToLower()))
                    {

                        //Backups current INFO
                        //Когда стоит микрорайон, а найдена улица, приориет у улицы
                        if ((string) typeOfStreetCell.Value == "микрорайон" &&
                            type != "микрорайон")
                            landmarkCell.Value += streetCell.Value + " " + typeOfStreetCell.Value + ", ";
                            //Когда стоит Именование, без типа
                        else if (typeOfStreetCell.Value == null && streetCell.Value != null)
                            landmarkCell.Value += streetCell.Value + ", ";

                        streetCell.Value = name;
                        typeOfStreetCell.Value = type;
                    }
                    //Отдельная логика для информации о доме
                    if (!String.IsNullOrEmpty(match.Groups["house_num"].Value))
                    {
                        if (buildCell.Value == null)
                            buildCell.Value = match.Groups["house_num"].Value;
                    }

                    value = reg.Replace(value, ", ");
                }
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="customTable"></param>
        /// <param name="cellsFilled"></param>
        /// <param name="regCenter"></param>
        /// <param name="regName"></param>
        private void TryFillClassificator(int row, ref DataTable customTable, ref bool cellsFilled, string regCenter,
            string regName)
        {
            var regionCell = worksheet.Cells[row, regionColumn];
            var settlementCell = worksheet.Cells[row, settlementColumn];
            var nearCityCell = worksheet.Cells[row, nearCityColumn];
            var typeOfNearCity = worksheet.Cells[row, typeOfNearCityColumn];
            var subjCell = worksheet.Cells[row, subjColumn];
            var landmarkCell = worksheet.Cells[row, additionalInfoColumn];

            if (customTable == null) return;
            if (cellsFilled) return;

            //
            //Записываем город если он у нас один 
            //
            var cities = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMOWorksheet.Columns.NearCity - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (cities.Count == 1)
            {
                if (nearCityCell.Value == null)
                {
                    if (nearCityCell.Value != null)
                        landmarkCell.Value += nearCityCell.Value.ToString() + ", ";
                    nearCityCell.Value = cities.First();
                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
                else if (nearCityCell.Value != null &&
                         string.Equals(nearCityCell.Value.ToString(), cities.First(), StringComparison.OrdinalIgnoreCase))
                {
                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
            }
                //Проверяем текущее значение на принадлежность к выборке
            else
            {
                if ((nearCityCell.Value != null))
                    if (cities.All(s => s != nearCityCell.Value.ToString()))
                    {
                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else
                    {
                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
            }
//
            //Записываем тип если он у один
            //
            var types = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMOWorksheet.Columns.TypeOfNearCity - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (types.Count == 1 && nearCityCell.Value != null)
            {
                if (typeOfNearCity.Value == null ||
                    !String.Equals(typeOfNearCity.Value.ToString(), types.First(), StringComparison.OrdinalIgnoreCase))
                {
                    typeOfNearCity.Value = types.First();
                    typeOfNearCity.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    typeOfNearCity.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                }
            }
                //Если в нашей выборке нельсколько типов населенных пунктов
                //И мы уже имеем какой-то тип
            else if (typeOfNearCity.Value != null)
                //Пробуем использовать тип для уточнения выборки
                if (
                    types.Any(
                        s => String.Equals(typeOfNearCity.Value.ToString(), s, StringComparison.OrdinalIgnoreCase)) &&
                    //1. Наш тип находится в пределах выборки
                    nearCityCell.Value != null && //2. у нас есть насел пункт
                    customTable.Rows.Cast<DataRow>()
                        .Any( //3. В выборке есть комбинация текущий насел пункт + текущий тип
                            r =>
                                String.Equals(r[OKTMOWorksheet.Columns.NearCity - 1].ToString(), nearCityCell.Value.ToString(),
                                    StringComparison.OrdinalIgnoreCase) &&
                                String.Equals(r[OKTMOWorksheet.Columns.TypeOfNearCity - 1].ToString(),
                                    typeOfNearCity.Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    //И тогда  мы уточняем выборку по типу населенного пункта
                    customTable = oktmo.GetCustomDataTable(customTable,
                        new SearchParams(typeOfNearCity.Value.ToString(), OKTMOColumns.TypeOfNearCity));
                    {
                        typeOfNearCity.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        typeOfNearCity.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
                }
                else
                {
                    typeOfNearCity.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    typeOfNearCity.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                }



            //
            //Записываем регион (муниципальное образование)
            //
            var regions = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMOWorksheet.Columns.Region - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (regions.Count == 1)
            {
                if (regionCell.Value == null ||
                    !String.Equals(regionCell.Value.ToString(), regions.First(), StringComparison.OrdinalIgnoreCase))
                {
                    if (regionCell.Value != null)
                        landmarkCell.Value += regionCell.Value.ToString() + ", ";
                    regionCell.Value = regions.First();
                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
                else if (regionCell.Value != null &&
                         string.Equals(regionCell.Value.ToString(), regions.First(), StringComparison.OrdinalIgnoreCase))
                {
                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;;
                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;;
                    subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
            }
                //Проверяем текущее значение на принадлежность к выборке
            else
            {
                if (regionCell.Value != null)
                    if (regions.All(s => s != regionCell.Value.ToString()))
                    {
                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else
                    {
                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
            }


            //По возможности записываем поселение
            var settlements = customTable.Rows.Cast<DataRow>()
                .Select(r => r[OKTMOWorksheet.Columns.Settlement - 1])
                .OfType<string>()
                .Distinct().ToList();
            if (settlements.Count == 1)
            {
                if (settlementCell.Value == null ||
                    !String.Equals(settlementCell.Value.ToString(), settlements.First(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    if (settlementCell.Value != null)
                        landmarkCell.Value += settlementCell.Value.ToString() + ", ";
                    settlementCell.Value = settlements.First();
                    settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }
                else if (settlementCell.Value != null &&
                         string.Equals(settlementCell.Value.ToString(), settlements.First(),
                             StringComparison.OrdinalIgnoreCase))
                {
                    settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                }

            }
            else
                //Проверяем текущее значение на принадлежность к выборке
            {
                if (settlementCell.Value != null)
                    if (settlements.All(s => s != settlementCell.Value.ToString()))
                    {
                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                    }
                    else
                    {
                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        settlementCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        settlementCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                    }
            }

            if (customTable.Rows.Count == 1)
                cellsFilled = true;
        }

        /// <summary>
        /// Метод запускается после максимального заполнения Населенного пункта, т.к. сравнивается с ним
        /// </summary>
        private void FormatDistToRegCenter()
        {
            const string code = "DIST_REG_CENTER";
            var columnIndex = GetColumnIndex(code);
            var nearCColumnIndex =GetColumnIndex("DIST_NEAR_CITY");

            //Для проверки
            var distToDeadCity =
                new Regex(
                    @"(?<dist>\d(?:\d|\s|\,|\.)+)\s?км\.?\s*(?<incity>\b(?:от|до|за)\b\s(?<cityType>[а-я]+\.?\s?)?(?<cityName>[А-Я]\w+)?)?");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    continue;
                }

                var inCityCell = worksheet.Cells[cell.Start.Row, inCityColumn];
                var nearCityCell = worksheet.Cells[cell.Start.Row, nearCityColumn];

                distValue = cell.Value.ToString();
                if (distValue == "0")
                {
                    inCityCell.Value = "да";
                    continue;
                }
                if (Regex.IsMatch(distValue, @"^(\d|\.|,)+$")) continue;
                Match match = distToRegCenteRegex.Match(distValue);
                if (match.Success)
                {
                    if (Regex.IsMatch(match.Value, @"\bв\b\s", RegexOptions.IgnoreCase))
                    {
                        inCityCell.Value = "да";
                        cell.Value = string.Empty;
                    }
                    else if (Regex.IsMatch(match.Value, @"\bза\b\s", RegexOptions.IgnoreCase))
                    {
                        inCityCell.Value = "нет";
                        cell.Value = string.Empty;
                    }
                    else
                    {
                        inCityCell.Value = "нет";
                        if (nearCityCell.Value != null &&
                            nearCityCell.Value.ToString() == match.Groups["Name"].Value)
                        {
                            worksheet.Cells[cell.Start.Row, distToNearCityColumn].Value =
                                match.Groups["num"].Value;
                            cell.Value = string.Empty;
                        }
                        else
                        {
                            cell.Value = match.Groups["num"].Value;
                        }
                    }
                }
                else
                    cell.Value = string.Empty;
            }
        }

        private void TryFillRegion(int row, ref string content, ref DataTable customTable, DataTable subjTable,
            ref bool cellsFilled, Regex reg = null)
        {
            var regionCell = worksheet.Cells[row, regionColumn];
            var nearCityCell = worksheet.Cells[row, nearCityColumn];
            var vgtCell = worksheet.Cells[row, vgtColumn];
            var typeOfNearCity = worksheet.Cells[row, typeOfNearCityColumn];
            var landmarkCell = worksheet.Cells[row, additionalInfoColumn];
            var subjCell = worksheet.Cells[row, subjColumn];

            Regex tmpRegex;
            Match match;
            if (reg == null)
            {
                tmpRegex = regionRegex;
                match = tmpRegex.Match(content);
                if (!match.Success)
                {
                    tmpRegex = regionToLeftRegex;
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
                var fullName = oktmo.GetFullName(name, OKTMOColumns.Region, type);


                //Spet 1: Подходит ли регион к субъекту
                if (!string.IsNullOrEmpty(fullName) &&
                    oktmo.StringMatchInColumn(subjTable, fullName, OKTMOColumns.Region))
                {
                    //Отлично, найденное мунОбр-е относится к текущему субъекту
                    //подтверждаем что нам надо использовать найденный текст
                    if (regionCell.Value == null ||
                        (!string.Equals(regionCell.Value.ToString(), fullName, StringComparison.OrdinalIgnoreCase) &&
                         regionCell.Style.Fill.BackgroundColor.Rgb == ExcelExtensions.BadColor.ToArgb().ToString()))
                    {
                        if (regionCell.Value != null)
                            AppendToLandMarkCell(regionCell.Value.ToString(), row);

                        regionCell.Value = fullName;

                        regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);

                        //Выборка
                        if (oktmo.StringMatchInColumn(customTable, fullName, OKTMOColumns.Region))
                        {
                            customTable = oktmo.GetCustomDataTable(customTable,
                                new SearchParams(fullName, OKTMOColumns.Region));
                            regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                            nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.Clear);
                        }
                            //Окрашиваем Регион если он подходит к субъекту но не подходит к выборке
                        else
                        {
                            regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                            nearCityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            nearCityCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                        }
                    }
                        //Запись в ориентир если текущий рег центр верный. Найденный нам просто не нужен
                    else if (regionCell.Value != null &&
                             !string.Equals(regionCell.Value.ToString(), fullName, StringComparison.OrdinalIgnoreCase))
                        AppendToLandMarkCell(fullName, row);

//                    content = tmpRegex.Replace(content, ", ");
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
                        if (!TryFillVGT(row, ref name, ref customTable, ref cellsFilled))
                        {
                            //Step 3: проверяем принадлежность региона к ОКТМО
                            //Проверялось при заполнеии fullName. Если заполнено, значи есть в ОКТМО
                            if (!string.IsNullOrEmpty(fullName))
                            {
                                //Если стоит верный, найденный кидаем в ориентир
                                if (regionCell.Value != null && regionCell.Style.Fill.BackgroundColor.Rgb == ExcelExtensions.BadColor.ToArgb().ToString())
                                {
                                    AppendToLandMarkCell(fullName, row);
                                }
                                    //Тут мы пишем неверный в пустую ячейку, либо заменяем один неверный на другой
                                else
                                {
                                    if (regionCell.Value != null)
                                        AppendToLandMarkCell(regionCell.Value.ToString(), row);
                                    regionCell.Value = fullName;
                                    regionCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    regionCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                    subjCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    subjCell.Style.Fill.BackgroundColor.SetColor(ExcelExtensions.BadColor);
                                    
                                }
                            }
                                //Запись в оринетир если найденный регион не существует в ОКТМО
                            else
                                AppendToLandMarkCell(name + " " + type, row);
                        }
                    }
                }
                content = tmpRegex.Replace(content, ", ");
            }
        }

        [Obsolete("Не доделано",true)]
        private void TryFillCommunications(long row,ref string val)
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
            Regex sentenceReg = new Regex(@"(?n)(?<=(^|\b))(?!\.)[^\!\?$^]{5,}?(?=((?<!\s\w{1,2})\.|\!|\?|$))", RegexOptions.Multiline);
            MatchCollection sentencesMatchCollection = sentenceReg.Matches(val);


            //Общие регулярки
            const string wordUnions = @"\s*(и|\,|;|:)\s*"; //Пунктуационанные знаки в предложении, объединяющие части предложения
            const string sentenceEnds = @"\s*((<!\\s\w{1,4})\.|!|?)"; //Пунктационные знаки, обозначающие конец предложения
            
            const string justWords = @"\(w(\w)*|\s)+"; //Паттерн для выялвения просто слов

            const string startCollocation = @"(?<=^|\""|(?<!\s\w{1,2})\.|\,|\)|\()"; //Символы, обозначающие начало предложения
            const string endCollocation = @"(?=$|\""|(?<!\s\w{1,3})\.(\s|$|\,)|\,|\)|\()"; //Символы, обозначающие конец предложения

            const string orDel = @"|"; //Символ Или
            const string spacesNRq = @"\s*"; //Наличие пробела в кол-ве от нуля до бесконечности

            //==========
            //Статус
            //TODO дефолтный значения при точномм, не точном наличии или отсутствии
            //Перечень фраз для подтверждения наличия коммуникации
            const string comValid = @"(?<valid>круглый\sгод|всегда|подведен(о|ы)|централизирован(а|о)|(?!в\sобществе\s)проводят|провед(ё|е)но?(?!\sк\sгранице)|на\s(участке|территории)|есть(?!\s*возможность)|име(е|ю)тся|(?<kvt>\d(\d|\.|\,)*)\s*квт)";
            const string comCanConnectAlwaysLeft = @"";
            const string comCanConnectAlwaysRight = @"";
            //Перечень фраз для подтверждения возможности провести коммуникацию
            const string comCanConnect = @"(?<canconn>в\sперспективах|\bТУ\b|проводится|будет|проведут|в\sобществе\sпроводят|легко\sпровести|оплачивается\sотдельно|(проведен\s(к|по))?границе|подключение(\sту)?|рядом\sпроходит|(есть\s)?возможно(сть)?|в\s\d+\sм(\.|етрах|\s)|актуально(\sпровести)?|разешени(е|я)|около|техусловия|соласовано|(на|по)\sулице|не\sдалеко)";
            const string comNo = @"(?<no>нет|отсутству(е|ю)т)"; //Фразы, подтверждающие отсутствие коммуникации
            const string comTemp = @"(?<temp>летний|зимний)"; //Наличие сезонной коммуникации


            const string delimCom = @"(\s*(\,|\.)\s*)"; //Символы разделители между преречисленными коммуникациями

            const string commonCommunicatuionNames = @"\b(?<all>коммуникации|удобства)\b"; //Паттерн с перечнем всех коммуникация
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
            const string anyCom = "(?<anyCom>" + commonCommunicatuionNames + orDel + electrNames + orDel + waterNames + orDel + gasNames +
                         orDel + severageNames + ")";

            //Строка-паттерн-перечень всех возможных коммуникаций для выявления их в прямой последновательности
            const string stringOfAnyCom = "(" +delimCom   + anyCom + @"|\s*\,\s*";




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
                startCollocation +  anyCom + spacesNRq + comValid+ spacesNRq + endCollocation,

                //                  "свет, вода          Недоступен"
                startCollocation +  anyCom + spacesNRq + comNo+ spacesNRq + endCollocation,

                //                  "свет, вода          Возможно проведение"
                startCollocation +  anyCom + spacesNRq + comCanConnect+ spacesNRq + endCollocation,

            };

            //Цикл для обработки всех вариантов, представленных выше
            foreach (Regex reg in patterns.Select(funcS => new Regex(funcS)))
            {

            }
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
        /// Метод возвращает переданную строку в формате Первая буква заглавная, остальные прописные
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

            string result = s;
            foreach (Match match in words)
            {
                var newWord = Regex.Replace(match.Value, justWordPattern,
                    m => String.Format("{0}{1}", m.Groups[1].Value.ToUpper(), m.Groups[2].Value.ToLower()));
                result = result.Replace(match.Value, newWord);
            }
            return result;
        }

        private static string TryChangeRegionEndness(string s)
        {
            var reg = new Regex(@"(ого|ом|ем)\b");

            var match = reg.Match(s);
            if (! match.Success) return s;

            string newString = reg.Replace(s, "ий");

            return newString;
        }

        private static string TryChangeSubjectEndness(string s)
        {
            var reg = new Regex(@"(ой)\b");

            var match = reg.Match(s);
            if (!match.Success) return s;

            string newString = reg.Replace(s, "ая");

            return newString;
        }

        public static string ReplaceYO(string s)
        {
            var s2 = s.Replace("ё", "е");
            var s3 = s2.Replace("Ё", "Е");
            return s3;
        }
    }
}

// ReSharper restore UnusedMember.Local