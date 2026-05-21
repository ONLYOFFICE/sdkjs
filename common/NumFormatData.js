/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

(function(window) {

	//Excel uses DateSeparator with 2 letters only in date patterns
	//todo fi-FI locale use a dot as a time separator, but Excel still wouldn’t display or recognize it
var g_aCultureInfos = {
	1: {LCID: 1, Name: "ar", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.س.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], AbbreviatedMonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "134", LongDatePattern: "dd/mmmm/yyyy"},
	4: {LCID: 4, Name: "zh-Hans", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	5: {LCID: 5, Name: "cs", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Kč", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["neděle", "pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota"], AbbreviatedDayNames: ["ne", "po", "út", "st", "čt", "pá", "so"], MonthNames: ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec", ""], AbbreviatedMonthNames: ["led", "úno", "bře", "dub", "kvě", "čvn", "čvc", "srp", "zář", "říj", "lis", "pro", ""], MonthGenitiveNames: ["ledna", "února", "března", "dubna", "května", "června", "července", "srpna", "září", "října", "listopadu", "prosince", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "odp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	6: {LCID: 6, Name: "da", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["søndag", "mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag"], AbbreviatedDayNames: ["sø", "ma", "ti", "on", "to", "fr", "lø"], MonthNames: ["januar", "februar", "marts", "april", "maj", "juni", "juli", "august", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy"},
	7: {LCID: 7, Name: "de", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	8: {LCID: 8, Name: "el", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Κυριακή", "Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο"], AbbreviatedDayNames: ["Κυρ", "Δευ", "Τρι", "Τετ", "Πεμ", "Παρ", "Σαβ"], MonthNames: ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος", "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος", ""], AbbreviatedMonthNames: ["Ιαν", "Φεβ", "Μαρ", "Απρ", "Μαϊ", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ", ""], MonthGenitiveNames: ["Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου", "Ιουνίου", "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου", "Νοεμβρίου", "Δεκεμβρίου", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "πμ", PMDesignator: "μμ", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	9: {LCID: 9, Name: "en", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ d\\,\\ yyyy"},
	10: {LCID: 10, Name: "es", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["do.", "lu.", "ma.", "mi.", "ju.", "vi.", "sá."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	11: {LCID: 11, Name: "fi", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sunnuntai", "maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai"], AbbreviatedDayNames: ["su", "ma", "ti", "ke", "to", "pe", "la"], MonthNames: ["tammikuu", "helmikuu", "maaliskuu", "huhtikuu", "toukokuu", "kesäkuu", "heinäkuu", "elokuu", "syyskuu", "lokakuu", "marraskuu", "joulukuu", ""], AbbreviatedMonthNames: ["tammi", "helmi", "maalis", "huhti", "touko", "kesä", "heinä", "elo", "syys", "loka", "marras", "joulu", ""], MonthGenitiveNames: ["tammikuuta", "helmikuuta", "maaliskuuta", "huhtikuuta", "toukokuuta", "kesäkuuta", "heinäkuuta", "elokuuta", "syyskuuta", "lokakuuta", "marraskuuta", "joulukuuta", ""], AbbreviatedMonthGenitiveNames: ["tammik.", "helmik.", "maalisk.", "huhtik.", "toukok.", "kesäk.", "heinäk.", "elok.", "syysk.", "lokak.", "marrask.", "jouluk.", ""], AMDesignator: "ap.", PMDesignator: "ip.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	12: {LCID: 12, Name: "fr", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	14: {LCID: 14, Name: "hu", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Ft", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat"], AbbreviatedDayNames: ["V", "H", "K", "Sze", "Cs", "P", "Szo"], MonthNames: ["január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "febr.", "márc.", "ápr.", "máj.", "jún.", "júl.", "aug.", "szept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "de.", PMDesignator: "du.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.\\ mmmm\\ d\\.\\,\\ dddd"},
	16: {LCID: 16, Name: "it", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"], AbbreviatedDayNames: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"], MonthNames: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre", ""], AbbreviatedMonthNames: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	17: {LCID: 17, Name: "ja", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"], AbbreviatedDayNames: ["日", "月", "火", "水", "木", "金", "土"], MonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "午前", PMDesignator: "午後", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	18: {LCID: 18, Name: "ko", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₩", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"], AbbreviatedDayNames: ["일", "월", "화", "수", "목", "금", "토"], MonthNames: ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "오전", PMDesignator: "오후", UseAMPM: 1, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"년\"\\ m\"월\"\\ d\"일\"\\ dddd"},
	21: {LCID: 21, Name: "pl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "zł", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"], AbbreviatedDayNames: ["niedz.", "pon.", "wt.", "śr.", "czw.", "pt.", "sob."], MonthNames: ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień", ""], AbbreviatedMonthNames: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru", ""], MonthGenitiveNames: ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września", "października", "listopada", "grudnia", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	22: {LCID: 22, Name: "pt", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "R$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"], AbbreviatedDayNames: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"], MonthNames: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro", ""], AbbreviatedMonthNames: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	25: {LCID: 25, Name: "ru", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₽", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", ""], AbbreviatedMonthNames: ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\ \"г.\""},
	29: {LCID: 29, Name: "sv", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"], AbbreviatedDayNames: ["sön", "mån", "tis", "ons", "tor", "fre", "lör"], MonthNames: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "\"den \"d\\ mmmm\\ yyyy"},
	31: {LCID: 31, Name: "tr", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₺", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"], AbbreviatedDayNames: ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"], MonthNames: ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık", ""], AbbreviatedMonthNames: ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ÖÖ", PMDesignator: "ÖS", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "d\\ mmmm\\ yyyy\\ dddd"},
	33: {LCID: 33, Name: "id", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Rp", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"], AbbreviatedDayNames: ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"], MonthNames: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	34: {LCID: 34, Name: "uk", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₴", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["неділя", "понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота"], AbbreviatedDayNames: ["Нд", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["січень", "лютий", "березень", "квітень", "травень", "червень", "липень", "серпень", "вересень", "жовтень", "листопад", "грудень", ""], AbbreviatedMonthNames: ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру", ""], MonthGenitiveNames: ["січня", "лютого", "березня", "квітня", "травня", "червня", "липня", "серпня", "вересня", "жовтня", "листопада", "грудня", ""], AbbreviatedMonthGenitiveNames: ["січ", "лют", "бер", "кві", "тра", "чер", "лип", "сер", "вер", "жов", "лис", "гру", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\" р.\""},
	36: {LCID: 36, Name: "sl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljek", "torek", "sreda", "četrtek", "petek", "sobota"], AbbreviatedDayNames: ["ned.", "pon.", "tor.", "sre.", "čet.", "pet.", "sob."], MonthNames: ["januar", "februar", "marec", "april", "maj", "junij", "julij", "avgust", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "feb.", "mar.", "apr.", "maj", "jun.", "jul.", "avg.", "sep.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "pop.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy"},
	38: {LCID: 38, Name: "lv", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["svētdiena", "pirmdiena", "otrdiena", "trešdiena", "ceturtdiena", "piektdiena", "sestdiena"], AbbreviatedDayNames: ["svētd.", "pirmd.", "otrd.", "trešd.", "ceturtd.", "piektd.", "sestd."], MonthNames: ["janvāris", "februāris", "marts", "aprīlis", "maijs", "jūnijs", "jūlijs", "augusts", "septembris", "oktobris", "novembris", "decembris", ""], AbbreviatedMonthNames: ["janv.", "febr.", "marts", "apr.", "maijs", "jūn.", "jūl.", "aug.", "sept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priekšp.", PMDesignator: "pēcp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ yyyy\\.\\ \"gada\"\\ d\\.\\ mmmm"},
	39: {LCID: 39, Name: "lt", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sekmadienis", "pirmadienis", "antradienis", "trečiadienis", "ketvirtadienis", "penktadienis", "šeštadienis"], AbbreviatedDayNames: ["sk", "pr", "an", "tr", "kt", "pn", "št"], MonthNames: ["sausis", "vasaris", "kovas", "balandis", "gegužė", "birželis", "liepa", "rugpjūtis", "rugsėjis", "spalis", "lapkritis", "gruodis", ""], AbbreviatedMonthNames: ["saus.", "vas.", "kov.", "bal.", "geg.", "birž.", "liep.", "rugp.", "rugs.", "spal.", "lapkr.", "gruod.", ""], MonthGenitiveNames: ["sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio", "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priešpiet", PMDesignator: "popiet", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\ \"m\"\\.\\ mmmm\\ d\\ \"d\"\\.\\,\\ dddd"},
	42: {LCID: 42, Name: "vi", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₫", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Chủ Nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"], AbbreviatedDayNames: ["CN", "T2", "T3", "T4", "T5", "T6", "T7"], MonthNames: ["Tháng Giêng", "Tháng Hai", "Tháng Ba", "Tháng Tư", "Tháng Năm", "Tháng Sáu", "Tháng Bảy", "Tháng Tám", "Tháng Chín", "Tháng Mười", "Tháng Mười Một", "Tháng Mười Hai", ""], AbbreviatedMonthNames: ["Thg1", "Thg2", "Thg3", "Thg4", "Thg5", "Thg6", "Thg7", "Thg8", "Thg9", "Thg10", "Thg11", "Thg12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "SA", PMDesignator: "CH", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	44: {LCID: 44, Name: "az", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["bazar", "bazar ertəsi", "çərşənbə axşamı", "çərşənbə", "cümə axşamı", "cümə", "şənbə"], AbbreviatedDayNames: ["B.", "B.E.", "Ç.A.", "Ç.", "C.A.", "C.", "Ş."], MonthNames: ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul", "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr", ""], AbbreviatedMonthNames: ["yan", "fev", "mar", "apr", "may", "iyn", "iyl", "avq", "sen", "okt", "noy", "dek", ""], MonthGenitiveNames: ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\,\\ dddd"},
	63: {LCID: 63, Name: "kk", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₸", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["жексенбі", "дүйсенбі", "сейсенбі", "сәрсенбі", "бейсенбі", "жұма", "сенбі"], AbbreviatedDayNames: ["жс", "дс", "сс", "ср", "бс", "жм", "сб"], MonthNames: ["Қаңтар", "Ақпан", "Наурыз", "Сәуір", "Мамыр", "Маусым", "Шілде", "Тамыз", "Қыркүйек", "Қазан", "Қараша", "Желтоқсан", ""], AbbreviatedMonthNames: ["қаң.", "ақп.", "нау.", "сәу.", "мам.", "мау.", "шіл.", "там.", "қыр.", "қаз.", "қар.", "жел.", ""], MonthGenitiveNames: ["қаңтар", "ақпан", "наурыз", "сәуір", "мамыр", "маусым", "шілде", "тамыз", "қыркүйек", "қазан", "қараша", "желтоқсан", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "yyyy\\ \"ж\"\\.\\ d\\ mmmm\\,\\ dddd"},
	80: {LCID: 80, Name: "mn", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["ням", "даваа", "мягмар", "лхагва", "пүрэв", "баасан", "бямба"], AbbreviatedDayNames: ["Ня", "Да", "Мя", "Лх", "Пү", "Ба", "Бя"], MonthNames: ["Нэгдүгээр сар", "Хоёрдугаар сар", "Гуравдугаар сар", "Дөрөвдүгээр сар", "Тавдугаар сар", "Зургаадугаар сар", "Долоодугаар сар", "Наймдугаар сар", "Есдүгээр сар", "Аравдугаар сар", "Арван нэгдүгээр сар", "Арван хоёрдугаар сар", ""], AbbreviatedMonthNames: ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар", "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар", ""], MonthGenitiveNames: ["нэгдүгээр сар", "хоёрдугаар сар", "гуравдугаар сар", "дөрөвдүгээр сар", "тавдугаар сар", "зургаадугаар сар", "долоодугаар сар", "наймдугаар сар", "есдүгээр сар", "аравдугаар сар", "арван нэгдүгээр сар", "арван хоёрдугаар сар", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ү.ө.", PMDesignator: "ү.х.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.mm\\.dd\\,\\ dddd"},
	1025: {LCID: 1025, Name: "ar-SA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.س.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], AbbreviatedMonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "134", LongDatePattern: "dd/mmmm/yyyy"},
	1026: {LCID: 1026, Name: "bg-BG", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "лв.", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["неделя", "понеделник", "вторник", "сряда", "четвъртък", "петък", "събота"], AbbreviatedDayNames: ["нед", "пон", "вт", "ср", "четв", "пет", "съб"], MonthNames: ["януари", "февруари", "март", "април", "май", "юни", "юли", "август", "септември", "октомври", "ноември", "декември", ""], AbbreviatedMonthNames: ["яну", "фев", "мар", "апр", "май", "юни", "юли", "авг", "сеп", "окт", "ное", "дек", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dd\\ mmmm\\ yyyy\\ \"г.\""},
	1028: {LCID: 1028, Name: "zh-TW", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "NT$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	1029: {LCID: 1029, Name: "cs-CZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Kč", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["neděle", "pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota"], AbbreviatedDayNames: ["ne", "po", "út", "st", "čt", "pá", "so"], MonthNames: ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec", ""], AbbreviatedMonthNames: ["led", "úno", "bře", "dub", "kvě", "čvn", "čvc", "srp", "zář", "říj", "lis", "pro", ""], MonthGenitiveNames: ["ledna", "února", "března", "dubna", "května", "června", "července", "srpna", "září", "října", "listopadu", "prosince", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "odp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	1030: {LCID: 1030, Name: "da-DK", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["søndag", "mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag"], AbbreviatedDayNames: ["sø", "ma", "ti", "on", "to", "fr", "lø"], MonthNames: ["januar", "februar", "marts", "april", "maj", "juni", "juli", "august", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy"},
	1031: {LCID: 1031, Name: "de-DE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	1032: {LCID: 1032, Name: "el-GR", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Κυριακή", "Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο"], AbbreviatedDayNames: ["Κυρ", "Δευ", "Τρι", "Τετ", "Πεμ", "Παρ", "Σαβ"], MonthNames: ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος", "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος", ""], AbbreviatedMonthNames: ["Ιαν", "Φεβ", "Μαρ", "Απρ", "Μαϊ", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ", ""], MonthGenitiveNames: ["Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου", "Ιουνίου", "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου", "Νοεμβρίου", "Δεκεμβρίου", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "πμ", PMDesignator: "μμ", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	1033: {LCID: 1033, Name: "en-US", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ d\\,\\ yyyy"},
	1035: {LCID: 1035, Name: "fi-FI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sunnuntai", "maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai"], AbbreviatedDayNames: ["su", "ma", "ti", "ke", "to", "pe", "la"], MonthNames: ["tammikuu", "helmikuu", "maaliskuu", "huhtikuu", "toukokuu", "kesäkuu", "heinäkuu", "elokuu", "syyskuu", "lokakuu", "marraskuu", "joulukuu", ""], AbbreviatedMonthNames: ["tammi", "helmi", "maalis", "huhti", "touko", "kesä", "heinä", "elo", "syys", "loka", "marras", "joulu", ""], MonthGenitiveNames: ["tammikuuta", "helmikuuta", "maaliskuuta", "huhtikuuta", "toukokuuta", "kesäkuuta", "heinäkuuta", "elokuuta", "syyskuuta", "lokakuuta", "marraskuuta", "joulukuuta", ""], AbbreviatedMonthGenitiveNames: ["tammik.", "helmik.", "maalisk.", "huhtik.", "toukok.", "kesäk.", "heinäk.", "elok.", "syysk.", "lokak.", "marrask.", "jouluk.", ""], AMDesignator: "ap.", PMDesignator: "ip.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	1036: {LCID: 1036, Name: "fr-FR", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	1038: {LCID: 1038, Name: "hu-HU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Ft", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat"], AbbreviatedDayNames: ["V", "H", "K", "Sze", "Cs", "P", "Szo"], MonthNames: ["január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "febr.", "márc.", "ápr.", "máj.", "jún.", "júl.", "aug.", "szept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "de.", PMDesignator: "du.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.\\ mmmm\\ d\\.\\,\\ dddd"},
	1040: {LCID: 1040, Name: "it-IT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"], AbbreviatedDayNames: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"], MonthNames: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre", ""], AbbreviatedMonthNames: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	1041: {LCID: 1041, Name: "ja-JP", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"], AbbreviatedDayNames: ["日", "月", "火", "水", "木", "金", "土"], MonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "午前", PMDesignator: "午後", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	1042: {LCID: 1042, Name: "ko-KR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₩", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"], AbbreviatedDayNames: ["일", "월", "화", "수", "목", "금", "토"], MonthNames: ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "오전", PMDesignator: "오후", UseAMPM: 1, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"년\"\\ m\"월\"\\ d\"일\"\\ dddd"},
	1043: {LCID: 1043, Name: "nl-NL", CurrencyPositivePattern: 2, CurrencyNegativePattern: 12, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"], AbbreviatedDayNames: ["zo", "ma", "di", "wo", "do", "vr", "za"], MonthNames: ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	1045: {LCID: 1045, Name: "pl-PL", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "zł", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"], AbbreviatedDayNames: ["niedz.", "pon.", "wt.", "śr.", "czw.", "pt.", "sob."], MonthNames: ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień", ""], AbbreviatedMonthNames: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru", ""], MonthGenitiveNames: ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września", "października", "listopada", "grudnia", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	1046: {LCID: 1046, Name: "pt-BR", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "R$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"], AbbreviatedDayNames: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"], MonthNames: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro", ""], AbbreviatedMonthNames: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	1049: {LCID: 1049, Name: "ru-RU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₽", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", ""], AbbreviatedMonthNames: ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\ \"г.\""},
	1050: {LCID: 1050, Name: "hr-HR", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedjelja", "ponedjeljak", "utorak", "srijeda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sri", "čet", "pet", "sub"], MonthNames: ["siječanj", "veljača", "ožujak", "travanj", "svibanj", "lipanj", "srpanj", "kolovoz", "rujan", "listopad", "studeni", "prosinac", ""], AbbreviatedMonthNames: ["sij", "vlj", "ožu", "tra", "svi", "lip", "srp", "kol", "ruj", "lis", "stu", "pro", ""], MonthGenitiveNames: ["siječnja", "veljače", "ožujka", "travnja", "svibnja", "lipnja", "srpnja", "kolovoza", "rujna", "listopada", "studenog", "prosinca", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "d\\.\\ mmmm\\ yyyy\\."},
	1051: {LCID: 1051, Name: "sk-SK", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["nedeľa", "pondelok", "utorok", "streda", "štvrtok", "piatok", "sobota"], AbbreviatedDayNames: ["ne", "po", "ut", "st", "št", "pi", "so"], MonthNames: ["január", "február", "marec", "apríl", "máj", "jún", "júl", "august", "september", "október", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "máj", "jún", "júl", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: ["januára", "februára", "marca", "apríla", "mája", "júna", "júla", "augusta", "septembra", "októbra", "novembra", "decembra", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	1053: {LCID: 1053, Name: "sv-SE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"], AbbreviatedDayNames: ["sön", "mån", "tis", "ons", "tor", "fre", "lör"], MonthNames: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "\"den \"d\\ mmmm\\ yyyy"},
	1055: {LCID: 1055, Name: "tr-TR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₺", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"], AbbreviatedDayNames: ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"], MonthNames: ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık", ""], AbbreviatedMonthNames: ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ÖÖ", PMDesignator: "ÖS", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "d\\ mmmm\\ yyyy\\ dddd"},
	1057: {LCID: 1057, Name: "id-ID", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Rp", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"], AbbreviatedDayNames: ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"], MonthNames: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	1058: {LCID: 1058, Name: "uk-UA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₴", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["неділя", "понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота"], AbbreviatedDayNames: ["Нд", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["січень", "лютий", "березень", "квітень", "травень", "червень", "липень", "серпень", "вересень", "жовтень", "листопад", "грудень", ""], AbbreviatedMonthNames: ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру", ""], MonthGenitiveNames: ["січня", "лютого", "березня", "квітня", "травня", "червня", "липня", "серпня", "вересня", "жовтня", "листопада", "грудня", ""], AbbreviatedMonthGenitiveNames: ["січ", "лют", "бер", "кві", "тра", "чер", "лип", "сер", "вер", "жов", "лис", "гру", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\" р.\""},
	1060: {LCID: 1060, Name: "sl-SI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljek", "torek", "sreda", "četrtek", "petek", "sobota"], AbbreviatedDayNames: ["ned.", "pon.", "tor.", "sre.", "čet.", "pet.", "sob."], MonthNames: ["januar", "februar", "marec", "april", "maj", "junij", "julij", "avgust", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "feb.", "mar.", "apr.", "maj", "jun.", "jul.", "avg.", "sep.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "pop.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy"},
	1062: {LCID: 1062, Name: "lv-LV", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["svētdiena", "pirmdiena", "otrdiena", "trešdiena", "ceturtdiena", "piektdiena", "sestdiena"], AbbreviatedDayNames: ["svētd.", "pirmd.", "otrd.", "trešd.", "ceturtd.", "piektd.", "sestd."], MonthNames: ["janvāris", "februāris", "marts", "aprīlis", "maijs", "jūnijs", "jūlijs", "augusts", "septembris", "oktobris", "novembris", "decembris", ""], AbbreviatedMonthNames: ["janv.", "febr.", "marts", "apr.", "maijs", "jūn.", "jūl.", "aug.", "sept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priekšp.", PMDesignator: "pēcp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ yyyy\\.\\ \"gada\"\\ d\\.\\ mmmm"},
	1063: {LCID: 1063, Name: "lt-LT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sekmadienis", "pirmadienis", "antradienis", "trečiadienis", "ketvirtadienis", "penktadienis", "šeštadienis"], AbbreviatedDayNames: ["sk", "pr", "an", "tr", "kt", "pn", "št"], MonthNames: ["sausis", "vasaris", "kovas", "balandis", "gegužė", "birželis", "liepa", "rugpjūtis", "rugsėjis", "spalis", "lapkritis", "gruodis", ""], AbbreviatedMonthNames: ["saus.", "vas.", "kov.", "bal.", "geg.", "birž.", "liep.", "rugp.", "rugs.", "spal.", "lapkr.", "gruod.", ""], MonthGenitiveNames: ["sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio", "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priešpiet", PMDesignator: "popiet", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\ \"m\"\\.\\ mmmm\\ d\\ \"d\"\\.\\,\\ dddd"},
	1066: {LCID: 1066, Name: "vi-VN", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₫", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Chủ Nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"], AbbreviatedDayNames: ["CN", "T2", "T3", "T4", "T5", "T6", "T7"], MonthNames: ["Tháng Giêng", "Tháng Hai", "Tháng Ba", "Tháng Tư", "Tháng Năm", "Tháng Sáu", "Tháng Bảy", "Tháng Tám", "Tháng Chín", "Tháng Mười", "Tháng Mười Một", "Tháng Mười Hai", ""], AbbreviatedMonthNames: ["Thg1", "Thg2", "Thg3", "Thg4", "Thg5", "Thg6", "Thg7", "Thg8", "Thg9", "Thg10", "Thg11", "Thg12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "SA", PMDesignator: "CH", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	1068: {LCID: 1068, Name: "az-Latn-AZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["bazar", "bazar ertəsi", "çərşənbə axşamı", "çərşənbə", "cümə axşamı", "cümə", "şənbə"], AbbreviatedDayNames: ["B.", "B.E.", "Ç.A.", "Ç.", "C.A.", "C.", "Ş."], MonthNames: ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul", "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr", ""], AbbreviatedMonthNames: ["yan", "fev", "mar", "apr", "may", "iyn", "iyl", "avq", "sen", "okt", "noy", "dek", ""], MonthGenitiveNames: ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\,\\ dddd"},
	1087: {LCID: 1087, Name: "kk-KZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₸", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["жексенбі", "дүйсенбі", "сейсенбі", "сәрсенбі", "бейсенбі", "жұма", "сенбі"], AbbreviatedDayNames: ["жс", "дс", "сс", "ср", "бс", "жм", "сб"], MonthNames: ["Қаңтар", "Ақпан", "Наурыз", "Сәуір", "Мамыр", "Маусым", "Шілде", "Тамыз", "Қыркүйек", "Қазан", "Қараша", "Желтоқсан", ""], AbbreviatedMonthNames: ["қаң.", "ақп.", "нау.", "сәу.", "мам.", "мау.", "шіл.", "там.", "қыр.", "қаз.", "қар.", "жел.", ""], MonthGenitiveNames: ["қаңтар", "ақпан", "наурыз", "сәуір", "мамыр", "маусым", "шілде", "тамыз", "қыркүйек", "қазан", "қараша", "желтоқсан", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "yyyy\\ \"ж\"\\.\\ d\\ mmmm\\,\\ dddd"},
	1104: {LCID: 1104, Name: "mn-MN", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["ням", "даваа", "мягмар", "лхагва", "пүрэв", "баасан", "бямба"], AbbreviatedDayNames: ["Ня", "Да", "Мя", "Лх", "Пү", "Ба", "Бя"], MonthNames: ["Нэгдүгээр сар", "Хоёрдугаар сар", "Гуравдугаар сар", "Дөрөвдүгээр сар", "Тавдугаар сар", "Зургаадугаар сар", "Долоодугаар сар", "Наймдугаар сар", "Есдүгээр сар", "Аравдугаар сар", "Арван нэгдүгээр сар", "Арван хоёрдугаар сар", ""], AbbreviatedMonthNames: ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар", "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар", ""], MonthGenitiveNames: ["нэгдүгээр сар", "хоёрдугаар сар", "гуравдугаар сар", "дөрөвдүгээр сар", "тавдугаар сар", "зургаадугаар сар", "долоодугаар сар", "наймдугаар сар", "есдүгээр сар", "аравдугаар сар", "арван нэгдүгээр сар", "арван хоёрдугаар сар", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ү.ө.", PMDesignator: "ү.х.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.mm\\.dd\\,\\ dddd"},
	2049: {LCID: 2049, Name: "ar-IQ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ع.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	2052: {LCID: 2052, Name: "zh-CN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	2055: {LCID: 2055, Name: "de-CH", CurrencyPositivePattern: 2, CurrencyNegativePattern: 2, CurrencySymbol: "CHF", NumberDecimalSeparator: ".", NumberGroupSeparator: "’", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	2057: {LCID: 2057, Name: "en-GB", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "£", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	2058: {LCID: 2058, Name: "es-MX", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	2060: {LCID: 2060, Name: "fr-BE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "134", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	2064: {LCID: 2064, Name: "it-CH", CurrencyPositivePattern: 2, CurrencyNegativePattern: 2, CurrencySymbol: "CHF", NumberDecimalSeparator: ".", NumberGroupSeparator: "’", NumberGroupSizes: [3], DayNames: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"], AbbreviatedDayNames: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"], MonthNames: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre", ""], AbbreviatedMonthNames: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	2070: {LCID: 2070, Name: "pt-PT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"], AbbreviatedDayNames: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"], MonthNames: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro", ""], AbbreviatedMonthNames: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\" de \"mmmm\" de \"yyyy"},
	2073: {LCID: 2073, Name: "ru-MD", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "L", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["вс", "пн", "вт", "ср", "чт", "пт", "сб"], MonthNames: ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь", ""], AbbreviatedMonthNames: ["янв.", "февр.", "март", "апр.", "май", "июнь", "июль", "авг.", "сент.", "окт.", "нояб.", "дек.", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв.", "февр.", "мар.", "апр.", "мая", "июн.", "июл.", "авг.", "сент.", "окт.", "нояб.", "дек.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy\\ \"г\"\\."},
	2077: {LCID: 2077, Name: "sv-FI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"], AbbreviatedDayNames: ["sön", "mån", "tis", "ons", "tors", "fre", "lör"], MonthNames: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "feb.", "mars", "apr.", "maj", "juni", "juli", "aug.", "sep.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "fm", PMDesignator: "em", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	2092: {LCID: 2092, Name: "az-Cyrl-AZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["базар", "базар ертәси", "чәршәнбә ахшамы", "чәршәнбә", "ҹүмә ахшамы", "ҹүмә", "шәнбә"], AbbreviatedDayNames: ["Б", "Бе", "Ча", "Ч", "Ҹа", "Ҹ", "Ш"], MonthNames: ["jанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthNames: ["Јан", "Фев", "Мар", "Апр", "Мај", "Ијун", "Ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], MonthGenitiveNames: ["јанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthGenitiveNames: ["Јан", "Фев", "Мар", "Апр", "мая", "ијун", "ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy"},
	2128: {LCID: 2128, Name: "mn-Mong-CN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 0], DayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], AbbreviatedDayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], MonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], AbbreviatedMonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd"},
	3073: {LCID: 3073, Name: "ar-EG", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ج.م.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	3076: {LCID: 3076, Name: "zh-HK", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "HK$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	3079: {LCID: 3079, Name: "de-AT", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Jänner", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jän", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jän.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sep.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	3081: {LCID: 3081, Name: "en-AU", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	3082: {LCID: 3082, Name: "es-ES", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["do.", "lu.", "ma.", "mi.", "ju.", "vi.", "sá."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	3084: {LCID: 3084, Name: "fr-CA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 15, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "d\\ mmmm\\ yyyy"},
	3152: {LCID: 3152, Name: "mn-Mong-MN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 0], DayNames: ["ᠨᠢᠮ᠎ᠠ", "ᠳᠠᠸᠠ", "ᠮᠢᠭᠮᠠᠷ", "ᡀᠠᠭᠪᠠ", "ᠫᠦᠷᠪᠦ", "ᠪᠠᠰᠠᠩ", "ᠪᠢᠮᠪᠠ"], AbbreviatedDayNames: ["ᠨᠢᠮ᠎ᠠ", "ᠳᠠᠸᠠ", "ᠮᠢᠭᠮᠠᠷ", "ᡀᠠᠭᠪᠠ", "ᠫᠦᠷᠪᠦ", "ᠪᠠᠰᠠᠩ", "ᠪᠢᠮᠪᠠ"], MonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], AbbreviatedMonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd"},
	4097: {LCID: 4097, Name: "ar-LY", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ل.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	4100: {LCID: 4100, Name: "zh-SG", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	4103: {LCID: 4103, Name: "de-LU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	4105: {LCID: 4105, Name: "en-CA", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "mmmm\\ d\\,\\ yyyy"},
	4106: {LCID: 4106, Name: "es-GT", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Q", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	4108: {LCID: 4108, Name: "fr-CH", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CHF", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	4122: {LCID: 4122, Name: "hr-BA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "KM", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedjelja", "ponedjeljak", "utorak", "srijeda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sri", "čet", "pet", "sub"], MonthNames: ["siječanj", "veljača", "ožujak", "travanj", "svibanj", "lipanj", "srpanj", "kolovoz", "rujan", "listopad", "studeni", "prosinac", ""], AbbreviatedMonthNames: ["sij", "velj", "ožu", "tra", "svi", "lip", "srp", "kol", "ruj", "lis", "stu", "pro", ""], MonthGenitiveNames: ["siječnja", "veljače", "ožujka", "travnja", "svibnja", "lipnja", "srpnja", "kolovoza", "rujna", "listopada", "studenoga", "prosinca", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy\\."},
	5121: {LCID: 5121, Name: "ar-DZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ج.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	5124: {LCID: 5124, Name: "zh-MO", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "MOP", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	5127: {LCID: 5127, Name: "de-LI", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "CHF", NumberDecimalSeparator: ".", NumberGroupSeparator: "’", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	5129: {LCID: 5129, Name: "en-NZ", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	5130: {LCID: 5130, Name: "es-CR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₡", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	5132: {LCID: 5132, Name: "fr-LU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	6153: {LCID: 6153, Name: "en-IE", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "€", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	6154: {LCID: 6154, Name: "es-PA", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "B/.", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "315", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	6156: {LCID: 6156, Name: "fr-MC", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	7169: {LCID: 7169, Name: "ar-TN", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ت.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	7177: {LCID: 7177, Name: "en-ZA", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "R", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	7178: {LCID: 7178, Name: "es-DO", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	7180: {LCID: 7180, Name: "fr-029", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "EC$", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre", ""], AbbreviatedMonthNames: ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc.", ""], MonthGenitiveNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthGenitiveNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	8193: {LCID: 8193, Name: "ar-OM", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.ع.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	8201: {LCID: 8201, Name: "en-JM", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	8202: {LCID: 8202, Name: "es-VE", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "Bs.S", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	8204: {LCID: 8204, Name: "fr-RE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	9217: {LCID: 9217, Name: "ar-YE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.ي.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	9225: {LCID: 9225, Name: "en-029", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "EC$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	9226: {LCID: 9226, Name: "es-CO", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	9228: {LCID: 9228, Name: "fr-CD", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "FC", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	9242: {LCID: 9242, Name: "sr-Latn-RS", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "RSD", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljak", "utorak", "sreda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sre", "čet", "pet", "sub"], MonthNames: ["januar", "februar", "mart", "april", "maj", "jun", "jul", "avgust", "septembar", "oktobar", "novembar", "decembar", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "avg", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "pre podne", PMDesignator: "po podne", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy\\."},
	10241: {LCID: 10241, Name: "ar-SY", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ل.س.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	10249: {LCID: 10249, Name: "en-BZ", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	10250: {LCID: 10250, Name: "es-PE", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "S/", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", ""], AbbreviatedMonthNames: ["Ene.", "Feb.", "Mar.", "Abr.", "May.", "Jun.", "Jul.", "Ago.", "Set.", "Oct.", "Nov.", "Dic.", ""], MonthGenitiveNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "set.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	10252: {LCID: 10252, Name: "fr-SN", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	10266: {LCID: 10266, Name: "sr-Cyrl-RS", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "дин.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["недеља", "понедељак", "уторак", "среда", "четвртак", "петак", "субота"], AbbreviatedDayNames: ["нед.", "пон.", "ут.", "ср.", "чет.", "пет.", "суб."], MonthNames: ["јануар", "фебруар", "март", "април", "мај", "јун", "јул", "август", "септембар", "октобар", "новембар", "децембар", ""], AbbreviatedMonthNames: ["јан.", "феб.", "март", "апр.", "мај", "јун", "јул", "авг.", "септ.", "окт.", "нов.", "дец.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy\\."},
	11265: {LCID: 11265, Name: "ar-JO", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ا.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	11273: {LCID: 11273, Name: "en-TT", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	11274: {LCID: 11274, Name: "es-AR", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	11276: {LCID: 11276, Name: "fr-CM", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "FCFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "mat.", PMDesignator: "soir", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	12289: {LCID: 12289, Name: "ar-LB", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ل.ل.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	12297: {LCID: 12297, Name: "en-ZW", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "US$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	12298: {LCID: 12298, Name: "es-EC", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	12300: {LCID: 12300, Name: "fr-CI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	13313: {LCID: 13313, Name: "ar-KW", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ك.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	13321: {LCID: 13321, Name: "en-PH", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₱", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	13322: {LCID: 13322, Name: "es-CL", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	13324: {LCID: 13324, Name: "fr-ML", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	14337: {LCID: 14337, Name: "ar-AE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.إ.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	14345: {LCID: 14345, Name: "en-ID", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Rp", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	14346: {LCID: 14346, Name: "es-UY", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", ""], AbbreviatedMonthNames: ["Ene.", "Feb.", "Mar.", "Abr.", "May.", "Jun.", "Jul.", "Ago.", "Set.", "Oct.", "Nov.", "Dic.", ""], MonthGenitiveNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "set.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	14348: {LCID: 14348, Name: "fr-MA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "DH", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["jan.", "fév.", "mar.", "avr.", "mai", "jui.", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	15361: {LCID: 15361, Name: "ar-BH", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ب.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "ابريل", "مايو", "يونيو", "يوليو", "اغسطس", "سبتمبر", "اكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	15369: {LCID: 15369, Name: "en-HK", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	15370: {LCID: 15370, Name: "es-PY", CurrencyPositivePattern: 2, CurrencyNegativePattern: 12, CurrencySymbol: "₲", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	15372: {LCID: 15372, Name: "fr-HT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "G", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	16385: {LCID: 16385, Name: "ar-QA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.ق.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	16393: {LCID: 16393, Name: "en-IN", CurrencyPositivePattern: 2, CurrencyNegativePattern: 12, CurrencySymbol: "₹", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 2], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	16394: {LCID: 16394, Name: "es-BO", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Bs", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	17417: {LCID: 17417, Name: "en-MY", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "RM", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\,\\ yyyy"},
	17418: {LCID: 17418, Name: "es-SV", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	18441: {LCID: 18441, Name: "en-SG", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	18442: {LCID: 18442, Name: "es-HN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "L", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ dd\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	19466: {LCID: 19466, Name: "es-NI", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "C$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	20490: {LCID: 20490, Name: "es-PR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "315", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	21514: {LCID: 21514, Name: "es-US", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ dd\\,\\ yyyy"},
	22538: {LCID: 22538, Name: "es-419", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "XDR", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a.m.", PMDesignator: "p.m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	23562: {LCID: 23562, Name: "es-CU", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a.m.", PMDesignator: "p.m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	27674: {LCID: 27674, Name: "sr-Cyrl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "дин.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["недеља", "понедељак", "уторак", "среда", "четвртак", "петак", "субота"], AbbreviatedDayNames: ["нед.", "пон.", "ут.", "ср.", "чет.", "пет.", "суб."], MonthNames: ["јануар", "фебруар", "март", "април", "мај", "јун", "јул", "август", "септембар", "октобар", "новембар", "децембар", ""], AbbreviatedMonthNames: ["јан.", "феб.", "март", "апр.", "мај", "јун", "јул", "авг.", "септ.", "окт.", "нов.", "дец.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy\\."},
	28698: {LCID: 28698, Name: "sr-Latn", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "RSD", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljak", "utorak", "sreda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sre", "čet", "pet", "sub"], MonthNames: ["januar", "februar", "mart", "april", "maj", "jun", "jul", "avgust", "septembar", "oktobar", "novembar", "decembar", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "avg", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "pre podne", PMDesignator: "po podne", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy\\."},
	29740: {LCID: 29740, Name: "az-Cyrl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["базар", "базар ертәси", "чәршәнбә ахшамы", "чәршәнбә", "ҹүмә ахшамы", "ҹүмә", "шәнбә"], AbbreviatedDayNames: ["Б", "Бе", "Ча", "Ч", "Ҹа", "Ҹ", "Ш"], MonthNames: ["jанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthNames: ["Јан", "Фев", "Мар", "Апр", "Мај", "Ијун", "Ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], MonthGenitiveNames: ["јанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthGenitiveNames: ["Јан", "Фев", "Мар", "Апр", "мая", "ијун", "ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy"},
	30724: {LCID: 30724, Name: "zh", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	30764: {LCID: 30764, Name: "az-Latn", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["bazar", "bazar ertəsi", "çərşənbə axşamı", "çərşənbə", "cümə axşamı", "cümə", "şənbə"], AbbreviatedDayNames: ["B.", "B.E.", "Ç.A.", "Ç.", "C.A.", "C.", "Ş."], MonthNames: ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul", "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr", ""], AbbreviatedMonthNames: ["yan", "fev", "mar", "apr", "may", "iyn", "iyl", "avq", "sen", "okt", "noy", "dek", ""], MonthGenitiveNames: ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\,\\ dddd"},
	30800: {LCID: 30800, Name: "mn-Cyrl", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["ням", "даваа", "мягмар", "лхагва", "пүрэв", "баасан", "бямба"], AbbreviatedDayNames: ["Ня", "Да", "Мя", "Лх", "Пү", "Ба", "Бя"], MonthNames: ["Нэгдүгээр сар", "Хоёрдугаар сар", "Гуравдугаар сар", "Дөрөвдүгээр сар", "Тавдугаар сар", "Зургаадугаар сар", "Долоодугаар сар", "Наймдугаар сар", "Есдүгээр сар", "Аравдугаар сар", "Арван нэгдүгээр сар", "Арван хоёрдугаар сар", ""], AbbreviatedMonthNames: ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар", "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар", ""], MonthGenitiveNames: ["нэгдүгээр сар", "хоёрдугаар сар", "гуравдугаар сар", "дөрөвдүгээр сар", "тавдугаар сар", "зургаадугаар сар", "долоодугаар сар", "наймдугаар сар", "есдүгээр сар", "аравдугаар сар", "арван нэгдүгээр сар", "арван хоёрдугаар сар", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ү.ө.", PMDesignator: "ү.х.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.mm\\.dd\\,\\ dddd"},
	31748: {LCID: 31748, Name: "zh-Hant", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "HK$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	31824: {LCID: 31824, Name: "mn-Mong", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 0], DayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], AbbreviatedDayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], MonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], AbbreviatedMonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd"},
};
	var g_aAdditionalCurrencySymbols = ["ADP", "AED", "AFA", "AFN", "ALL", "AMD", "ANG", "AOA", "ARS", "ATS", "AUD",
		"AWG", "AZM", "AZN", "BAM", "BBD", "BDT", "BEF", "BGL", "BGN", "BHD", "BIF", "BMD", "BND", "BOB", "BOV", "BRL",
		"BSD", "BTN", "BWP", "BYB", "BYN", "BYR", "BZD", "CAD", "CDF", "CHE", "CHF", "CHW", "CLF", "CLP", "CNY", "COP",
		"COU", "CRC", "CSD", "CUC", "CUP", "CVE", "CYP", "CZK", "DEM", "DJF", "DKK", "DOP", "DZD", "ECS", "ECV", "EEK",
		"EGP", "ERN", "ESP", "ETB", "EUR", "FIM", "FJD", "FKP", "FRF", "GBP", "GEL", "GHC", "GHS", "GIP", "GMD", "GNF",
		"GRD", "GTQ", "GYD", "HKD", "HNL", "HRK", "HTG", "HUF", "IDR", "IEP", "ILS", "INR", "IQD", "IRR", "ISK", "ITL",
		"JMD", "JOD", "JPY", "KAF", "KES", "KGS", "KHR", "KMF", "KPW", "KRW", "KWD", "KYD", "KZT", "LAK", "LBP", "LKR",
		"LRD", "LSL", "LTL", "LUF", "LVL", "LYD", "MAD", "MDL", "MGA", "MGF", "MKD", "MMK", "MNT", "MOP", "MRO", "MRU",
		"MTL", "MUR", "MVR", "MWK", "MXN", "MXV", "MYR", "MZM", "MZN", "NAD", "NGN", "NIO", "NLG", "NOK", "NPR", "NTD",
		"NZD", "OMR", "PAB", "PEN", "PGK", "PHP", "PKR", "PLN", "PTE", "PYG", "QAR", "ROL", "RON", "RSD", "RUB", "RUR",
		"RWF", "SAR", "SBD", "SCR", "SDD", "SDG", "SDP", "SEK", "SGD", "SHP", "SIT", "SKK", "SLL", "SOS", "SPL", "SRD",
		"SRG", "STD", "SVC", "SYP", "SZL", "THB", "TJR", "TJS", "TMM", "TMT", "TND", "TOP", "TRL", "TRY", "TTD", "TWD",
		"TZS", "UAH", "UGX", "USD", "USN", "USS", "UYI", "UYU", "UZS", "VEB", "VEF", "VES", "VND", "VUV", "WST", "XAF",
		"XAG", "XAU", "XB5", "XBA", "XBB", "XBC", "XBD", "XCD", "XDR", "XFO", "XFU", "XOF", "XPD", "XPF", "XPT", "XTS",
		"XXX", "YER", "YUM", "ZAR", "ZMK", "ZMW", "ZWD", "ZWL", "ZWN", "ZWR"
	];

	let c_oAscDateFormatExcel = {
		"1025": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"1026": [
			"dd\\.m\\.yyyy\\ \"г.\";@",
			"d\\.m\\.yyyy\\ \"г.\";@",
			"dd\\.mm\\.yyyy\\ \"г.\";@",
			"yyyy\\-mm\\-dd;@",
			"[$-402]dd\\ mmmm\\ yyyy\\ \"г.\";@",
			"[$-402]dddd\\,\\ dd\\ mmmm\\ yyyy\\ \"г.\";@"
		],
		"1027": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-403]dddd\\,\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-403]d\" de \"mmmm\" de \"yyyy;@",
			"[$-403]d\" \"mmmm\" \"yyyy;@"
		],
		"1028": [
			"yyyy\\-mm\\-dd;@",
			"yyyy\"年\"m\"月\"d\"日\";@",
			"m\"月\"d\"日\";@",
			"[DBNum1][$-404]yyyy\"年\"m\"月\"d\"日\";@",
			"[DBNum1][$-404]m\"月\"d\"日\";@",
			"[$-404]aaaa;@",
			"[$-404]aaa;@",
			"yyyy/m/d;@",
			"yyyy/m/d\\ h:mm;@",
			"[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@",
			"m/d;@",
			"m/d/yy;@",
			"mm/dd/yy;@",
			"[$-409]d\\-mmm;@",
			"[$-409]d\\-mmm\\-yy;@",
			"[$-409]mmmmm;@",
			"[$-409]mmmmm\\-yy;@"
		],
		"1029": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-405]d\\-mmm\\.;@",
			"[$-405]d/mmm/yy;@",
			"[$-405]dd\\-mmm\\-yy;@",
			"[$-405]mmm\\-yy;@",
			"[$-405]mmmm\\ yy;@",
			"[$-405]d\\.\\ mmmm\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@",
			"[$-405]mmmmm;@",
			"[$-405]mmmmm\\-yy;@",
			"d/m/yyyy;@",
			"[$-405]d\\-mmm\\-yyyy;@"
		],
		"1030": [
			"dd\\-mm\\-yy;@",
			"[$-406]d\\.\\ mmmm\\ yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"yyyy\\.mm\\.dd;@",
			"yy\\-mm\\-dd;@",
			"[$-406]mmmm\\ yyyy;@",
			"d\\.m\\.yy;@",
			"d/m\\ yyyy;@",
			"dd\\-mm\\-yy\\ hh:mm;@",
			"dd\\-mm\\-yy\\ hh:mm:ss;@",
			"yyyy\\-mm\\-dd\\ hh:mm;@",
			"[$-406]mmmm\\ yy;@",
			"dd\\.mm\\.yyyy;@",
			"d\\.m\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"dd/mm\\ yyyy;@",
			"dd/mm\\ yy;@",
			"d/m\\ yy;@",
			"[$-406]mmmmm;@",
			"[$-406]mmmmm\\-yy;@"
		],
		"1031": [
			"yyyy\\-mm\\-dd;@",
			"d\\.m;@",
			"d\\.m\\.yy;@",
			"dd\\.mm\\.yy;@",
			"[$-407]d\\.\\ mmm\\.;@",
			"[$-407]d\\.\\ mmm\\.\\ yy;@",
			"[$-407]d\\.\\ mmm\\ yy;@",
			"[$-407]mmm\\.\\ yy;@",
			"[$-407]mmmm\\ yy;@",
			"[$-407]d\\.\\ mmmm\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d\\.m\\.yy\\ h:mm;@",
			"[$-407]mmmmm;@",
			"[$-407]mmmmm\\ yy;@",
			"d\\.m\\.yyyy;@",
			"[$-407]d\\.\\ mmm\\.\\ yyyy;@"
		],
		"1032": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"d/m/yyyy;@",
			"[$-408]d\\-mmm;@",
			"[$-408]d\\-mmm\\-yy;@",
			"[$-408]dd\\-mmm\\-yy;@",
			"[$-408]mmm\\-yy;@",
			"[$-408]d\\ mmmm\\ yyyy;@",
			"[$-408]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@",
			"[$-408]mmmmm;@",
			"[$-408]mmmmm\\-yy;@",
			"[$-408]d\\-mmm\\-yyyy;@"
		],
		"1033": [
			"[$-1070000]d/m/yy;@",
			"[$-1070000]d/mm/yyyy;@",
			"[$-1070000]d/mm/yyyy\\ h:mm\\ \"น.\";@",
			"[$-1070409]d/mm/yyyy\\ h:mm\\ AM/PM;@",
			"[$-D070000]d/m/yy;@",
			"[$-D070000]d/mm/yyyy;@",
			"[$-D070000]d/mm/yyyy\\ h:mm\\ \"น.\";@",
			"[$-D07041E]d\\ mmm\\ yy;@",
			"[$-D07041E]d\\ mmmm\\ yyyy;@",
			"[$-107041E]d\\ mmm\\ yy;@",
			"[$-107041E]d\\ mmmm\\ yyyy;@"
		],
		"1034": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"dd\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-40A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-40A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-40A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"1035": [
			"d\\.m\\.;@",
			"d\\.m\\.yy;@",
			"d\\.m\\.yyyy;@",
			"[$-40B]d\\.\\ mmmm\\t\\a;@",
			"[$-40B]d\\.\\ mmmm\\t\\a\\ yy;@",
			"[$-40B]d\\.\\ mmmm\\t\\a\\ yyyy;@",
			"[$-40B]mmmm\\ yy;@",
			"[$-40B]mmmm\\ yyyy;@",
			"[$-40B]d\\.\\ mmmm\\t\\a\\ yyyy\\ h:mm;@",
			"d\\.m\\.yyyy\\ h:mm;@",
			"d\\.m\\.yy\\ h:mm;@",
			"[$-40B]mmmmm;@",
			"[$-40B]mmmmm\\ yy;@",
			"yyyy\\-mm\\-dd;@",
			"yyyy\\-mm\\-dd\\ hh:mm;@"
		],
		"1036": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-40C]d\\-mmm;@",
			"[$-40C]d\\-mmm\\-yy;@",
			"[$-40C]dd\\-mmm\\-yy;@",
			"[$-40C]mmm\\-yy;@",
			"[$-40C]mmmm\\-yy;@",
			"[$-40C]d\\ mmmm\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@",
			"[$-40C]mmmmm;@",
			"[$-40C]mmmmm\\-yy;@",
			"m/d/yyyy;@",
			"[$-40C]d\\-mmm\\-yyyy;@"
		],
		"1038": [
			"m\\.\\ d\\.;@",
			"yyyy/\\ m/\\ d\\.;@",
			"yyyy/mm/dd;@",
			"[$-40E]yyyy/\\ mmm/\\ d\\.;@",
			"[$-40E]yy/\\ mmmm\\ d\\.;@",
			"[$-40E]mmmm\\ d\\.;@",
			"[$-40E]yyyy/\\ mmm\\.;@",
			"[$-40E]yyyy/\\ mmmm;@",
			"[$-40E]yyyy/\\ mmmm\\ d\\.;@",
			"[$-40E]yyyy/\\ m/\\ d\\.\\ h:mm\\ AM/PM;@",
			"yyyy/\\ m/\\ d\\.\\ h:mm;@",
			"[$-40E]mmm/\\ d\\.;@",
			"yyyy\\-mm\\-dd;@",
			"yyyy\\ mm\\ dd;@",
			"yyyy\\.mm\\.dd;@",
			"[$-40E]mmmmm\\.;@",
			"[$-40E]yy\\-mmmmm\\.;@",
			"[$-40E]yy/\\ mmmm;@"
		],
		"1039": [
			"d\\.m\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.\\ m\\.\\ yyyy\\.;@",
			"d\\.\\ m\\.\\ \"'\"yy\\.;@",
			"yyyy\\-mm\\-dd;@",
			"yy\\ mm\\ dd;@",
			"[$-40F]d\\.\\ mmmm\\ yyyy;@",
			"[$-40F]dd\\.\\ mmmm\\ yyyy;@"
		],
		"1040": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-410]d\\-mmm;@",
			"[$-410]d\\-mmm\\-yy;@",
			"[$-410]dd\\-mmm\\-yy;@",
			"[$-410]mmm\\-yy;@",
			"[$-410]mmmm\\-yy;@",
			"[$-410]d\\ mmmm\\ yyyy;@",
			"[$-409]d/m/yy\\ h\\.mm\\ AM/PM;@",
			"d/m/yy\\ h\\.mm;@",
			"[$-410]mmmmm;@",
			"[$-410]mmmmm\\-yy;@",
			"d/m/yyyy;@",
			"[$-410]d\\-mmm\\-yyyy;@"
		],
		"1042": [
			"yyyy\\-mm\\-dd;@",
			"yyyy\"년\"\\ m\"월\"\\ d\"일\";@",
			"yy\"年\"\\ m\"月\"\\ d\"日\";@",
			"yyyy\"년\"\\ m\"월\";@",
			"m\"월\"\\ d\"일\";@",
			"yy\"-\"m\"-\"d;@",
			"yy\"-\"m\"-\"d\\ h:mm;@",
			"[$-412]yy\"-\"m\"-\"d\\ AM/PM\\ h:mm;@",
			"[$-409]yy\"-\"m\"-\"d\\ h:mm\\ AM/PM;@",
			"yy\"/\"m\"/\"d;@",
			"yyyy\"-\"m\"-\"d;@",
			"yyyy\"/\"m\"/\"d;@",
			"m\"/\"d;@",
			"m\"/\"d\"/\"yy;@",
			"mm\"/\"dd\"/\"yy;@",
			"[$-409]d\"-\"mmm;@",
			"[$-409]d\"-\"mmm\"-\"yy;@",
			"[$-409]mmm\"-\"yy;@",
			"[$-409]mmmm\"-\"yy;@",
			"[$-409]mmmmm;@",
			"[$-409]mmmmm\\-yy;@"
		],
		"1043": [
			"yyyy\\-mm\\-dd;@",
			"d\\-m;@",
			"d\\-mm\\-yy;@",
			"dd\\-mm\\-yy;@",
			"[$-413]d\\-mmm;@",
			"[$-413]d\\-mmm\\-yy;@",
			"[$-413]dd\\-mmm\\-yy;@",
			"[$-413]mmm\\-yy;@",
			"[$-413]mmmm\\-yy;@",
			"[$-413]d\\ mmmm\\ yyyy;@",
			"[$-409]d\\-mm\\-yy\\ h:mm\\ AM/PM;@",
			"d\\-mm\\-yy\\ h:mm;@",
			"[$-413]mmmmm;@",
			"[$-413]mmmmm\\-yy;@",
			"m/d/yyyy;@",
			"[$-413]d\\-mmm\\-yyyy;@"
		],
		"1044": [
			"d/m/;@",
			"d/m/yy;@",
			"d/m/yyyy;@",
			"dd/mm/yy;@",
			"dd/mm/yyyy;@",
			"[$-414]d/\\ mmm\\.;@",
			"[$-414]d/\\ mmmm;@",
			"[$-414]d/\\ mmm\\.\\ yyyy;@",
			"[$-414]d/\\ mmmm\\ yyyy;@",
			"[$-414]mmm\\.\\ yy;@",
			"[$-414]mmmm\\ yy;@",
			"[$-414]mmmm\\ yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"dd/mm/yy\\ h:mm;@",
			"[$-409]m/d/yy\\ h:mm\\ AM/PM;@",
			"m/d/yy\\ h:mm;@"
		],
		"1045": [
			"d\\-mm;@",
			"yyyy\\-mm\\-dd;@",
			"yy\\-mm\\-dd;@",
			"[$-415]d\\ mmm;@",
			"[$-415]d\\ mmm\\ yy;@",
			"[$-415]dd\\ mmm\\ yy;@",
			"[$-415]mmm\\ yy;@",
			"[$-415]mmmm\\ yy;@",
			"[$-415]d\\ mmmm\\ yyyy;@",
			"[$-409]dd\\-mm\\-yy\\ h:mm\\ AM/PM;@",
			"dd\\-mm\\-yy\\ h:mm;@",
			"[$-415]mmmmm;@",
			"[$-415]mmmmm\\.yy;@",
			"d\\-m\\-yyyy;@",
			"[$-415]d\\-mmm\\-yyyy;@"
		],
		"1046": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-416]d\\-mmm;@",
			"[$-416]d\\-mmm\\-yy;@",
			"[$-416]dd\\-mmm\\-yy;@",
			"[$-416]mmm\\-yy;@",
			"[$-416]mmmm\\-yy;@",
			"[$-416]d;@",
			"mmmm\\,\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@"
		],
		"1047": [
			"yyyy\\-mm\\-dd;@",
			"[$-10417]dd\\-mm\\-yyyy;@",
			"[$-10417]dd\\-mm\\-yy;@",
			"[$-10417]dddd\\,\\ \"ils’\"\\ d\\.\\ mmmm\\,\\ yyyy;@",
			"[$-10417]dddd\\,\\ \"ils\"\\ d\\ mmmm\\ yyyy;@",
			"[$-10417]d\\ mmmm\\ yyyy;@"
		],
		"1048": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-418]d\\-mmm;@",
			"[$-418]d\\-mmm\\-yy;@",
			"[$-418]dd\\-mmm\\-yy;@",
			"[$-418]mmm\\-yy;@",
			"[$-418]mmmm\\-yy;@",
			"[$-418]d\\ mmmm\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@",
			"[$-418]mmmmm;@",
			"[$-418]mmmmm\\-yy;@",
			"d/m/yyyy;@",
			"[$-418]d\\-mmm\\-yyyy;@"
		],
		"1049": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-419]d\\ mmm;@",
			"[$-419]d\\ mmm\\ yy;@",
			"[$-419]dd\\ mmm\\ yy;@",
			"[$-F419]yyyy\\,\\ mmmm;@",
			"[$-419]mmmm\\ yyyy;@",
			"[$-FC19]dd\\ mmmm\\ yyyy\\ \\г\\.;@",
			"[$-409]dd/mm/yy\\ h:mm\\ AM/PM;@",
			"dd/mm/yy\\ h:mm;@",
			"[$-419]mmmm;@",
			"[$-FC19]yyyy\\,\\ dd\\ mmmm;@",
			"d/m/yyyy;@",
			"[$-419]d\\-mmm\\-yyyy;@"
		],
		"1050": [
			"yyyy\\-mm\\-dd;@",
			"d\\.m\\.;@",
			"d\\.m\\.yy\\.;@",
			"dd\\.mm\\.yy\\.;@",
			"[$-41A]d\\-mmm;@",
			"[$-41A]d\\-mmm\\-yy;@",
			"[$-41A]dd\\-mmm\\-yy;@",
			"[$-41A]mmm\\-yy;@",
			"[$-41A]mmmm\\-yy;@",
			"[$-41A]d\\.\\ mmmm\\ yyyy\\.;@",
			"[$-409]d\\.m\\.yy\\.\\ h:mm\\ AM/PM;@",
			"d\\.m\\.yy\\.\\ h:mm;@",
			"[$-41A]mmmmm;@",
			"[$-41A]mmmmm\\-yy\\.;@",
			"d\\.m\\.yyyy\\.;@",
			"[$-41A]d\\-mmm\\-yyyy\\.;@"
		],
		"1051": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-41B]d\\-mmm\\.;@",
			"[$-41B]d/mmm/yy;@",
			"[$-41B]dd\\-mmm\\-yy;@",
			"[$-41B]mmm\\-yy;@",
			"[$-41B]mmmm\\ yy;@",
			"[$-41B]d\\.\\ mmmm\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@",
			"[$-41B]mmmmm;@",
			"[$-41B]mmmmm\\-yy;@",
			"d/m/yyyy;@",
			"[$-41B]d/mmm/yyyy;@"
		],
		"1053": [
			"yyyy\\-mm\\-dd;@",
			"yyyy\\-mm\\-dd\\ hh:mm;@",
			"yy\\-mm\\-dd;@",
			"yy\\-mm\\-dd\\ hh:mm;@",
			"d/m\\ yyyy;@",
			"d/m\\ \\-yy;@",
			"d/m\\ yy;@",
			"d/m/yy;@",
			"[$-41D]\"den \"\\ d\\ mmmm\\ yyyy;@",
			"[$-41D]d\\ mmmm\\ yyyy;@",
			"[$-41D]d\\ mmmm\\ \\-yy;@",
			"[$-41D]mmmmm;@",
			"[$-41D]mmmmm\\-yy;@",
			"yyyy\\ mm\\ dd;@",
			"[$-41D]mmmm;@",
			"[$-41D]dd\\-mmm;@",
			"[$-41D]mmmm\\ yyyy;@",
			"[$-41D]mmmm\\ \\-yy;@",
			"[$-41D]mmm\\-yy;@",
			"yyyy;@"
		],
		"1055": [
			"yyyy\\-mm\\-dd;@",
			"d/m;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-41F]d\\ mmmm;@",
			"[$-41F]d\\ mmmm\\ yy;@",
			"[$-41F]dd\\ mmmm\\ yy;@",
			"dd/mm/yyyy;@",
			"[$-41F]mmmm\\ yy;@",
			"[$-41F]d\\ mmmm\\ yyyy;@",
			"d/m/yy\\ h:mm;@",
			"[$-41F]d\\ mmmm\\ yyyy\\ h:mm;@",
			"[$-41F]mmmmm;@",
			"[$-41F]mmmmm\\ yy;@",
			"m/d/yyyy;@",
			"[$-41F]d\\ mmm\\ yyyy;@"
		],
		"1056": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"m/d/yyyy;@",
			"m/d/yy;@",
			"mm/dd/yy;@",
			"mm/dd/yyyy;@",
			"[$-420]dd\\ mmmm\\,\\ yyyy;@",
			"[$-420]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-420]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-420]mmmm\\ dd\\,\\ yyyy;@",
			"[$-420]dd/mmmm/yyyy;@"
		],
		"1057": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-421]dd\\ mmmm\\ yyyy;@"
		],
		"1060": [
			"d\\.m\\.yyyy;@",
			"d\\.m\\.yy;@",
			"d\\.\\ m\\.\\ yyyy;@",
			"dd\\.mm\\.yyyy;@",
			"d\\.\\ m\\.\\ yy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.\\ mm\\.\\ yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-424]d\\.\\ mmmm\\ yyyy;@",
			"[$-424]dd\\.\\ mmmm\\ yyyy;@",
			"[$-424]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@"
		],
		"1061": [
			"d\\.mm\\.yyyy;@",
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-425]d\\.\\ mmmm\\ yyyy\". a.\";@",
			"[$-425]dd\\.\\ mmmm\\ yyyy\". a.\";@",
			"[$-425]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@"
		],
		"1062": [
			"yyyy\\.mm\\.dd\\.;@",
			"yy\\.mm\\.dd\\.;@",
			"yyyy\\-mm\\-dd;@",
			"[$-426]dddd\\,\\ yyyy\". gada \"d\\.\\ mmmm;@"
		],
		"1064": [
			"yyyy\\-mm\\-dd;@",
			"[$-10428]dd\\.mm\\.yyyy;@",
			"[$-10428]dd\\.mm\\.yy;@",
			"[$-10428]d\\.m\\.yy;@",
			"[$-10428]dd\\-mm\\-yyyy;@",
			"[$-10428]dd/mm/yy;@",
			"[$-10428]d\\ mmmm\\ yyyy\" с.\";@",
			"[$-10428]dd\\ mmmm\\ yyyy\" с.\";@",
			"[$-10428]dddd\\,\\ dd\\ mmmm\\ yyyy;@"
		],
		"1065": [
			"[$-160429]dd/mm/yyyy;@",
			"[$-160429]dd/mm/yy;@",
			"[$-160429]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-160429]d\\ mmmm\\ yyyy;@"
		],
		"1066": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yy;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-101042A]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-101042A]d\\ mmm\\ yy;@",
			"[$-101042A]d\\ mmmm\\ yyyy;@",
			"[$-101040C]d\\ mmm\\ yy;@",
			"[$-101040C]d\\ mmmm\\ yyyy;@",
			"[$-1010409]d\\ mmm\\ yy;@",
			"[$-1010409]d\\ mmmm\\ yyyy;@"
		],
		"1067": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d/mm/yyyy;@",
			"dd/mm/yyyy;@",
			"[$-42B]d/mmm/yyyy;@",
			"[$-42B]dd/mmm/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-42B]d\\ mmmm\\,\\ yyyy;@",
			"[$-42B]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-42B]dddd\\,\\ dd\\ mmmm\\ yyyy;@",
			"[$-42B]dd\\ mmmm\\ yyyy;@",
			"[$-42B]d\\-mmm\\-yyyy;@",
			"[$-42B]dd\\-mmm\\-yyyy;@",
			"[$-42B]ddd\\,\\ d\\-mmmm\\-yyyy;@",
			"[$-42B]ddd\\,\\ dd\\-mmmm\\-yyyy;@"
		],
		"1068": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-42C]d\\ mmmm\\ yyyy;@",
			"[$-42C]dd\\ mmmm\\ yyyy;@"
		],
		"1069": [
			"yyyy\\-mm\\-dd;@",
			"m/d;@",
			"yy/m/d;@",
			"yy/mm/dd;@",
			"[$-42D]mmm\\-d;@",
			"[$-42D]yy\\-mmm\\-d;@",
			"[$-42D]yy\\-mmm\\-dd;@",
			"[$-42D]yy\\-mmm;@",
			"[$-42D]yy\\-mmmm;@",
			"[$-42D]yyyy\"(e)ko\"\\ mmmm\"ren\"\\ d\"(a)\";@",
			"[$-42D]yy/mm/dd/\\ h:mm\\ AM/PM;@",
			"yy/m/d/\\ h:mm;@",
			"[$-42D]mmmmm;@",
			"[$-42D]yy\\-mmmmm;@",
			"yyyy/m/d;@",
			"[$-42D]yyyy\\-mmm\\-d;@",
			"yyyy\\.mm\\.dd;@",
			"[$-42D]yyyy\"(e)ko\"\\ mmmm\"ren\"\\ d\"(a)\";@",
			"[$-42D]yyyy\"(e)ko\"\\ mmmm\"k\"\\ d\"(a)\";@",
			"[$-42D]yyyy\"(e)ko\"\\ mmmm;@"
		],
		"1070": [
			"yyyy\\-mm\\-dd;@",
			"[$-1042E]d\\.m\\.yyyy;@",
			"[$-1042E]d\\.m\\.yy;@",
			"[$-1042E]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-1042E]d\\.\\ mmmm\\ yyyy;@"
		],
		"1071": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-42F]dddd\\,\\ dd\\ mmmm\\ yyyy;@"
		],
		"1072": [
			"[$-10430]yyyy\\-mm\\-dd;@",
			"[$-10430]yyyy\\ mmm\\ d;@",
			"[$-10430]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10430]yyyy\\ mmmm\\ d;@"
		],
		"1073": [
			"[$-10431]yyyy\\-mm\\-dd;@",
			"[$-10431]yyyy\\ mmm\\ d;@",
			"[$-10431]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10431]yyyy\\ mmmm\\ d;@"
		],
		"1074": [
			"[$-10432]yyyy\\-mm\\-dd;@",
			"[$-10432]yyyy\\ mmm\\ d;@",
			"[$-10432]dd\\ mmmm\\ yyyy;@",
			"[$-10432]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10432]yyyy\\ mmmm\\ d;@"
		],
		"1075": [
			"[$-10433]yyyy\\-mm\\-dd;@",
			"[$-10433]yyyy\\ mmm\\ d;@",
			"[$-10433]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10433]yyyy\\ mmmm\\ d;@"
		],
		"1076": [
			"yyyy\\-mm\\-dd;@",
			"[$-10434]m/d/yyyy;@",
			"[$-10434]m/d/yy;@",
			"[$-10434]mmm\\ d\\,\\ yyyy;@",
			"[$-10434]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10434]mmmm\\ d\\,\\ yyyy;@"
		],
		"1077": [
			"yyyy\\-mm\\-dd;@",
			"[$-10435]m/d/yyyy;@",
			"[$-10435]m/d/yy;@",
			"[$-10435]mmm\\ d\\,\\ yyyy;@",
			"[$-10435]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10435]mmmm\\ d\\,\\ yyyy;@"
		],
		"1078": [
			"yyyy/mm/dd;@",
			"yy/mm/dd;@",
			"yyyy\\-mm\\-dd;@",
			"[$-436]dd\\ mmmm\\ yyyy;@"
		],
		"1079": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-437]dddd\\,\\ d\\ mmmm\\,\\ yyyy\\ \"წელი\";@"
		],
		"1080": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-438]d\\.\\ mmmm\\ yyyy;@"
		],
		"1081": [
			"yyyy\\-mm\\-dd;@",
			"[$-4010000]d/m/yy;@",
			"[$-4010000]d/m/yyyy;@",
			"[$-4010439]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-4010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010000]d/m/yy;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010439]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010439]d\\ mmm\\ yy;@",
			"[$-1010439]d\\ mmmm\\ yyyy;@",
			"[$-4010439]d\\ mmm\\ yy;@",
			"[$-4010439]d\\ mmmm\\ yyyy;@",
			"[$-1010409]d\\ mmm\\ yy;@",
			"[$-1010409]d\\ mmmm\\ yyyy;@"
		],
		"1082": [
			"yyyy\\-mm\\-dd;@",
			"[$-1043A]dd/mm/yyyy;@",
			"[$-1043A]dd\\ mmm\\ yyyy;@",
			"[$-1043A]dddd\\,\\ d\\ \"ta\"\\’\\ mmmm\\ yyyy;@",
			"[$-1043A]d\\ \"ta\"\\’\\ mmmm\\ yyyy;@"
		],
		"1083": [
			"[$-1043B]yyyy\\-mm\\-dd;@",
			"[$-1043B]yyyy\\ mmm\\ d;@",
			"[$-1043B]dddd\\,\\ mmmm\\ d\". b. \"yyyy;@",
			"[$-1043B]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-1043B]yyyy\\ mmmm\\ d;@"
		],
		"1085": [
			"yyyy\\-mm\\-dd;@",
			"[$-1043D]dd/mm/yyyy;@",
			"[$-1043D]dd/mm/yy;@",
			"[$-1043D]d\\ט\\ן\\ mmm\\ yyyy;@",
			"[$-1043D]dddd\\,\\ d\\ט\\ן\\ mmmm\\ yyyy;@",
			"[$-1043D]d\\ט\\ן\\ mmmm\\ yyyy;@"
		],
		"1086": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-43E]dd\\ mmmm\\ yyyy;@"
		],
		"1087": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-43F]d\\ mmmm\\ yyyy\\ \"ж.\";@",
			"[$-43F]dd\\ mmmm\\ yyyy\\ \"ж.\";@"
		],
		"1088": [
			"yyyy\\-mm\\-dd;@",
			"dd\\.mm\\.yy;@",
			"[$-440]d\"-\"mmmm\\ yyyy\"-ж.\";@"
		],
		"1089": [
			"m/d/yyyy;@",
			"m/d/yy;@",
			"mm/dd/yy;@",
			"mm/dd/yyyy;@",
			"yy/mm/dd;@",
			"yyyy\\-mm\\-dd;@",
			"[$-441]dd\\-mmm\\-yy;@",
			"[$-441]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-441]mmmm\\ dd\\,\\ yyyy;@",
			"[$-441]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-441]dd\\ mmmm\\,\\ yyyy;@"
		],
		"1090": [
			"yyyy\\-mm\\-dd;@",
			"[$-10442]dd\\.mm\\.yy\\ \"ý.\";@",
			"[$-10442]dd\\.mm\\.yyyy;@",
			"[$-10442]yyyy\"-nji ýylyň \"d\"-nji \"mmmm;@",
			"[$-10442]d\\ mmmm\\ yyyy\\ dddd;@"
		],
		"1091": [
			"dd/mm\\ yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-443]yyyy\\ \"yil\"\\ d\\-mmmm;@",
			"[$-443]d\\ mmmm\\ yyyy;@",
			"[$-443]dd\\ mmmm\\ yyyy;@"
		],
		"1092": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-444]d\\ mmmm\\ yyyy;@",
			"[$-444]dd\\ mmmm\\ yyyy;@"
		],
		"1093": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-445]dd\\ mmmm\\ yyyy;@",
			"[$-445]d\\ mmmm\\ yyyy;@",
			"[$-5000445]dd\\-mm\\-yyyy;@",
			"[$-5000445]dd\\-mm\\-yy;@",
			"[$-5000445]d\\-m\\-yy;@",
			"[$-5000445]d\\.m\\.yy;@",
			"[$-5000445]yyyy\\-mm\\-dd;@",
			"[$-5000445]dd\\ mmmm\\ yyyy;@",
			"[$-5000445]d\\ mmmm\\ yyyy;@"
		],
		"1094": [
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"dd\\-mm\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-446]dd\\ mmmm\\ yyyy\\ dddd;@",
			"[$-446]d\\ mmmm\\ yyyy;@",
			"[$-6000446]dd\\-mm\\-yy;@",
			"[$-6000446]d\\-m\\-yy;@",
			"[$-6000446]d\\.m\\.yy;@",
			"[$-6000446]dd\\-mm\\-yyyy;@",
			"[$-6000446]yyyy\\-mm\\-dd;@",
			"[$-6000446]dd\\ mmmm\\ yyyy\\ dddd;@",
			"[$-6000446]d\\ mmmm\\ yyyy;@"
		],
		"1095": [
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"dd\\-mm\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-447]dd\\ mmmm\\ yyyy;@",
			"[$-447]d\\ mmmm\\ yyyy;@",
			"[$-7000447]dd\\-mm\\-yy;@",
			"[$-7000447]d\\-m\\-yy;@",
			"[$-7000447]d\\.m\\.yy;@",
			"[$-7000447]dd\\-mm\\-yyyy;@",
			"[$-7000447]yyyy\\-mm\\-dd;@",
			"[$-7000447]dd\\ mmmm\\ yyyy;@",
			"[$-7000447]d\\ mmmm\\ yyyy;@"
		],
		"1096": [
			"[$-10448]dd\\-mm\\-yy;@",
			"[$-10448]d\\-m\\-yy;@",
			"[$-10448]d\\.m\\.yy;@",
			"[$-10448]dd\\-mm\\-yyyy;@",
			"[$-10448]yyyy\\-mm\\-dd;@",
			"[$-10448]dd\\ mmmm\\ yyyy;@",
			"[$-10448]d\\ mmmm\\ yyyy;@",
			"[$-10448]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@"
		],
		"1097": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-449]dd\\ mmmm\\ yyyy;@",
			"[$-449]d\\ mmmm\\ yyyy;@",
			"[$-9000449]dd\\-mm\\-yyyy;@",
			"[$-9000449]dd\\-mm\\-yy;@",
			"[$-9000449]d\\-m\\-yy;@",
			"[$-9000449]d\\.m\\.yy;@",
			"[$-9000449]yyyy\\-mm\\-dd;@",
			"[$-9000449]dd\\ mmmm\\ yyyy;@",
			"[$-9000449]d\\ mmmm\\ yyyy;@"
		],
		"1098": [
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"dd\\-mm\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-44A]dd\\ mmmm\\ yyyy;@",
			"[$-44A]d\\ mmmm\\ yyyy;@",
			"[$-A00044A]dd\\-mm\\-yy;@",
			"[$-A00044A]d\\-m\\-yy;@",
			"[$-A00044A]d\\.m\\.yy;@",
			"[$-A00044A]dd\\-mm\\-yyyy;@",
			"[$-A00044A]yyyy\\-mm\\-dd;@",
			"[$-A00044A]dd\\ mmmm\\ yyyy;@",
			"[$-A00044A]d\\ mmmm\\ yyyy;@"
		],
		"1099": [
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"dd\\-mm\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-44B]dd\\ mmmm\\ yyyy;@",
			"[$-44B]d\\ mmmm\\ yyyy;@"
		],
		"1100": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-44C]dd\\ mmmm\\ yyyy;@",
			"[$-44C]d\\ mmmm\\ yyyy;@",
			"[$-C00044C]dd\\-mm\\-yyyy;@",
			"[$-C00044C]dd\\-mm\\-yy;@",
			"[$-C00044C]d\\-m\\-yy;@",
			"[$-C00044C]d\\.m\\.yy;@",
			"[$-C00044C]yyyy\\-mm\\-dd;@",
			"[$-C00044C]dd\\ mmmm\\ yyyy;@",
			"[$-C00044C]d\\ mmmm\\ yyyy;@"
		],
		"1101": [
			"yyyy\\-mm\\-dd;@",
			"[$-1044D]dd\\-mm\\-yyyy;@",
			"[$-1044D]yyyy\\,mmmm\\ dd\\,\\ dddd;@"
		],
		"1102": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-44E]dd\\ mmmm\\ yyyy;@",
			"[$-44E]d\\ mmmm\\ yyyy;@"
		],
		"1103": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-44F]dd\\ mmmm\\ yyyy\\ dddd;@",
			"[$-44F]d\\ mmmm\\ yyyy;@"
		],
		"1104": [
			"yyyy\\-mm\\-dd;@",
			"yy\\.mm\\.dd;@",
			"[$-450]yyyy\\ \"оны\"\\ mmmm\\ d;@"
		],
		"1105": [
			"[$-10451]yyyy/m/d;@",
			"[$-10451]yyyy\\-m\\-d;@",
			"[$-10451]yyyy\\.m\\.d;@",
			"[$-10451]yyyy\\.mm\\.dd;@",
			"[$-10451]yyyy\\-mm\\-dd;@",
			"[$-10451]yyyy/mm/dd;@",
			"[$-10451]yy\\-m\\-d;@",
			"[$-10451]yy/m/d;@",
			"[$-10451]yy\\.m\\.d;@",
			"[$-10451]yyyy\"ལོའི་ཟླ\"\\ m\"ཚེས\"\\ d;@",
			"[$-10451]yyyy\"ལོའི་ཟླ\"\\ m\"ཚེས\"\\ d\\ dddd;@",
			"[$-10451]yyyy\\ལ\\ོ\\འ\\ི\\་\\ཟ\\ླ\\ mmm\\ d;@",
			"[$-10451]yyyy\\ལ\\ོ\\འ\\ི\\་\\ཟ\\ླ\\ mmm\\ d\\ dddd;@"
		],
		"1106": [
			"yyyy\\-mm\\-dd;@",
			"[$-10452]dd/mm/yyyy;@",
			"[$-10452]dd/mm/yy;@",
			"[$-10452]d\\ mmm\\ yyyy;@",
			"[$-10452]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-10452]d\\ mmmm\\ yyyy;@"
		],
		"1107": [
			"[$-10453]dd/mm/yy;@",
			"[$-10453]yyyy\\-mm\\-dd;@",
			"[$-10453]d\\ mmmm\\ yyyy;@",
			"[$-10453]ddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10453]dddd\\ d\\ mmmm\\ yyyy;@"
		],
		"1108": [
			"yyyy\\-mm\\-dd;@",
			"[$-10454]d/m/yyyy;@",
			"[$-10454]d\\ mmm\\ yyyy;@",
			"[$-10454]dddd\\ \\ທ\\ີ\\ d\\ mmmm\\ gg\\ yyyy;@",
			"[$-10454]d\\ mmmm\\ yyyy;@"
		],
		"1109": [
			"yyyy\\-mm\\-dd;@",
			"[$-10455]d/m/yyyy;@",
			"[$-10455]d/m/yy;@",
			"[$-10455]yyyy\\၊\\ mmm\\ d;@",
			"[$-10455]yyyy\\၊\\ mmmm\\ d\\၊\\ dddd;@",
			"[$-10455]yyyy\\၊\\ mmmm\\ d;@"
		],
		"1110": [
			"yyyy\\-mm\\-dd;@",
			"dd/mm/yy;@",
			"d/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"dd\\.mm\\.yy;@",
			"dd/mm/yyyy;@",
			"[$-456]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-456]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-456]d\" de \"mmmm\" de \"yyyy;@"
		],
		"1111": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-457]dd\\ mmmm\\ yyyy;@",
			"[$-457]d\\ mmmm\\ yyyy;@"
		],
		"1112": [
			"yyyy\\-mm\\-dd;@",
			"[$-10458]d/m/yyyy;@",
			"[$-10458]d/m/yy;@",
			"[$-10458]mmm\\ d\\,\\ yyyy;@",
			"[$-10458]mmmm\\ d\\,\\ yyyy\\,\\ dddd;@",
			"[$-10458]mmmm\\ d\\,\\ yyyy;@"
		],
		"1113": [
			"yyyy\\-mm\\-dd;@",
			"[$-10459]m/d/yyyy;@",
			"[$-10459]m/d/yy;@",
			"[$-10459]mmm\\ d\\,\\ yyyy;@",
			"[$-10459]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10459]mmmm\\ d\\,\\ yyyy;@"
		],
		"1114": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-45A]dd\\ mmmm\\,\\ yyyy;@",
			"[$-45A]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@"
		],
		"1115": [
			"[$-1045B]yyyy\\-mm\\-dd;@",
			"[$-1045B]yyyy\\ mmm\\ d;@",
			"[$-1045B]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-1045B]yyyy\\ mmmm\\ d;@"
		],
		"1116": [
			"[$-1045C]m/d/yyyy;@",
			"[$-1045C]m/d/yy;@",
			"[$-1045C]mm/dd/yy;@",
			"[$-1045C]mm/dd/yyyy;@",
			"[$-1045C]yy/mm/dd;@",
			"[$-1045C]yyyy\\-mm\\-dd;@",
			"[$-1045C]dd\\-mmm\\-yy;@",
			"[$-1045C]dddd\\,\\ mmmm\\ dd\\,yyyy;@",
			"[$-1045C]mmmm\\ dd\\,yyyy;@",
			"[$-1045C]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-1045C]dd\\ mmmm\\,\\ yyyy;@"
		],
		"1117": [
			"[$-1045D]d/m/yyyy;@",
			"[$-1045D]d/m/yy;@",
			"[$-1045D]dd/mm/yy;@",
			"[$-1045D]yy/mm/dd;@",
			"[$-1045D]yyyy\\-mm\\-dd;@",
			"[$-1045D]dd\\-mmm\\-yy;@",
			"[$-1045D]dddd\\,mmmm\\ dd\\,yyyy;@",
			"[$-1045D]mmmm\\ dd\\,yyyy;@",
			"[$-1045D]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-1045D]dd\\ mmmm\\,\\ yyyy;@"
		],
		"1118": [
			"yyyy\\-mm\\-dd;@",
			"[$-1045E]dd/mm/yyyy;@",
			"[$-1045E]d\\ mmm\\ yyyy;@",
			"[$-1045E]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-1045E]d\\ mmmm\\ yyyy;@"
		],
		"1119": [
			"yyyy\\-mm\\-dd;@",
			"[$-1045F]d/m/yyyy;@",
			"[$-1045F]dd/mm/yyyy;@",
			"[$-1045F]dddd\\،\\ d\\ mmmm\\ yyyy;@",
			"[$-1045F]d\\ mmmm\\ yyyy;@"
		],
		"1120": [
			"yyyy\\-mm\\-dd;@",
			"[$-10460]m/d/yyyy;@",
			"[$-10460]m/d/yy;@",
			"[$-10460]mmm\\ d\\,\\ yyyy;@",
			"[$-10460]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10460]mmmm\\ d\\,\\ yyyy;@"
		],
		"1121": [
			"[$-10461]m/d/yyyy;@",
			"[$-10461]m/d/yy;@",
			"[$-10461]mm/dd/yy;@",
			"[$-10461]mm/dd/yyyy;@",
			"[$-10461]yy/mm/dd;@",
			"[$-10461]yyyy\\-mm\\-dd;@",
			"[$-10461]dd\\-mmm\\-yy;@",
			"[$-10461]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-10461]mmmm\\ dd\\,\\ yyyy;@",
			"[$-10461]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-10461]dd\\ mmmm\\,\\ yyyy;@"
		],
		"1122": [
			"yyyy\\-mm\\-dd;@",
			"[$-10462]dd\\-mm\\-yyyy;@",
			"[$-10462]dd\\-mm\\-yy;@",
			"[$-10462]d\\ mmm\\ yyyy;@",
			"[$-10462]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10462]d\\ mmmm\\ yyyy;@"
		],
		"1123": [
			"[$-160463]yyyy/m/d;@",
			"[$-160463]yyyy\\-mm\\-dd;@",
			"[$-160463]d\\ mmmm\\ yyyy;@",
			"[$-160463]dddd\\ d\\ mmmm\\ yyyy;@"
		],
		"1124": [
			"yyyy\\-mm\\-dd;@",
			"[$-10464]m/d/yyyy;@",
			"[$-10464]m/d/yy;@",
			"[$-10464]mmm\\ d\\,\\ yyyy;@",
			"[$-10464]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10464]mmmm\\ d\\,\\ yyyy;@"
		],
		"1125": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]dd/mm/yy;@",
			"[$-1010000]dd/mm/yyyy;@",
			"[$-1010000]dd\\ mm\\ yyyy;@",
			"[$-1010465]dd\\ mmm\\ yyyy;@",
			"[$-1010465]dd\\ mmmm\\ yyyy;@",
			"[$-1010465]ddd\\,\\ dd\\ mmmm\\ yyyy;@",
			"[$-1010465]ddd\\,\\ dd\\ mmm\\ yyyy;@",
			"[$-1010465]dddd\\,\\ dd\\ mmm\\ yyyy;@",
			"[$-1010465]dddd\\,\\ dd\\ mmmm\\ yyyy;@"
		],
		"1126": [
			"yyyy\\-mm\\-dd;@",
			"[$-10466]d/m/yyyy;@",
			"[$-10466]d\\ mmm\\ yyyy;@",
			"[$-10466]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-10466]mmmm\\ dd\\,\\ yyyy;@"
		],
		"1127": [
			"yyyy\\-mm\\-dd;@",
			"[$-10467]d/m/yyyy;@",
			"[$-10467]d\\ mmm\\,\\ yyyy;@",
			"[$-10467]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10467]d\\ mmmm\\ yyyy;@"
		],
		"1128": [
			"yyyy\\-mm\\-dd;@",
			"[$-10468]d/m/yyyy;@",
			"[$-10468]d/m/yy;@",
			"[$-10468]d\\ mmm\\,\\ yyyy;@",
			"[$-10468]dddd\\ d\\ mmmm\\,\\ yyyy;@",
			"[$-10468]d\\ mmmm\\,\\ yyyy;@"
		],
		"1129": [
			"yyyy\\-mm\\-dd;@",
			"[$-10469]d/m/yyyy;@",
			"[$-10469]d\\ mmm\\ yyyy;@",
			"[$-10469]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-10469]mmmm\\ dd\\,\\ yyyy;@"
		],
		"1130": [
			"yyyy\\-mm\\-dd;@",
			"[$-1046A]d/m/yyyy;@",
			"[$-1046A]d\\ mm\\ yyyy;@",
			"[$-1046A]dddd\\,\\ d\\ mmm\\ yyyy;@",
			"[$-1046A]d\\ mmm\\ yyyy;@"
		],
		"1131": [
			"[$-1046B]dd/mm/yyyy;@",
			"[$-1046B]dd/mm/yy;@",
			"[$-1046B]d/m/yy;@",
			"[$-1046B]dd\\-mm\\-yy;@",
			"[$-1046B]yyyy\\-mm\\-dd;@",
			"[$-1046B]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-1046B]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-1046B]d\" de \"mmmm\" de \"yyyy;@"
		],
		"1132": [
			"[$-1046C]yyyy\\-mm\\-dd;@",
			"[$-1046C]yyyy\\ mmm\\ d;@",
			"[$-1046C]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-1046C]yyyy\\ mmmm\\ d;@"
		],
		"1133": [
			"[$-1046D]dd\\.mm\\.yy;@",
			"[$-1046D]yyyy\\-mm\\-dd;@",
			"[$-1046D]d\\ mmmm\\ yyyy\\ \"й\";@",
			"[$-1046D]dddd\\ mmmm\\ yyyy\\ \"й\";@"
		],
		"1134": [
			"yyyy\\-mm\\-dd;@",
			"[$-1046E]dd\\.mm\\.yy;@",
			"[$-1046E]dd/mm/yy;@",
			"[$-1046E]dd\\-mm\\-yy;@",
			"[$-1046E]d\\.\\ mmmm\\ yyyy;@",
			"[$-1046E]dd\\.\\ mmmmyyyy;@",
			"[$-1046E]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-1046E]dddd\\,\\ dd\\.\\ mmmm\\ yyyy;@",
			"[$-1046E]dddd\\,\" den \"d\\.\\ mmmm\\ yyyy;@",
			"[$-1046E]dddd\\,\" den \"dd\\.\\ mmmm\\ yyyy;@"
		],
		"1135": [
			"[$-1046F]dd\\-mm\\-yyyy;@",
			"[$-1046F]dd\\-mm\\-yy;@",
			"[$-1046F]yyyy\\-mm\\-dd;@",
			"[$-1046F]yyyy\\ mm\\ dd;@",
			"[$-1046F]mmmm\\ d\".-at, \"yyyy;@",
			"[$-1046F]d\\.\\ mmmm\\ yyyy;@",
			"[$-1046F]dd\\.\\ mmmm\\ yyyy;@",
			"[$-1046F]dddd\\ dd\\ mmmm\\ yyyy;@"
		],
		"1136": [
			"yyyy\\-mm\\-dd;@",
			"[$-10470]d/m/yyyy;@",
			"[$-10470]d/m/yy;@",
			"[$-10470]d\\ mmm\\ yyyy;@",
			"[$-10470]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-10470]d\\ mmmm\\ yyyy;@"
		],
		"1137": [
			"yyyy\\-mm\\-dd;@",
			"[$-10471]d/m/yyyy;@",
			"[$-10471]mmm\\ d\\,\\ yyyy;@",
			"[$-10471]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-10471]mmmm\\ dd\\,\\ yyyy;@"
		],
		"1138": [
			"yyyy\\-mm\\-dd;@",
			"[$-10472]dd/mm/yyyy;@",
			"[$-10472]dd/mm/yy;@",
			"[$-10472]dd\\-mmm\\-yyyy;@",
			"[$-10472]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10472]dd\\ mmmm\\ yyyy;@"
		],
		"1139": [
			"yyyy\\-mm\\-dd;@",
			"[$-10473]d/m/yyyy;@",
			"[$-10473]dd/mm/yyyy;@",
			"[$-10473]dd/mm/yy;@",
			"[$-10473]d\\ mmm\\ yyyy;@",
			"[$-10473]dddd\\ \"፣\"\\ mmmm\\ d\\ \"መዓልቲ\"\\ yyyy;@",
			"[$-10473]dddd\\፣\\ d\\ mmmm\\ yyyy;@",
			"[$-10473]d\\ mmmm\\ yyyy;@"
		],
		"1140": [
			"yyyy\\-mm\\-dd;@",
			"[$-10474]dd/mm/yyyy;@",
			"[$-10474]dd/mm/yy;@",
			"[$-10474]dd\\-mm\\-yyyy;@",
			"[$-10474]dd\\-mm\\-yy;@",
			"[$-10474]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-10474]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@",
			"[$-10474]dd/mmmm/yyyy;@",
			"[$-10474]d/mmmm/yyyy;@",
			"[$-10474]dd\\ mmmm\\,\\ yyyy;@",
			"[$-10474]d\\ mmmm\\,\\ yyyy;@"
		],
		"1141": [
			"yyyy\\-mm\\-dd;@",
			"[$-10475]d/m/yyyy;@",
			"[$-10475]d/m/yy;@",
			"[$-10475]d\\ mmm\\ yyyy;@",
			"[$-10475]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-10475]d\\ mmmm\\ yyyy;@"
		],
		"1142": [
			"yyyy\\-mm\\-dd;@",
			"[$-10476]d\\ m\\ yyyy\\ gg;@",
			"[$-10476]\"die\"\\ d\\ mmm\\ yyyy\\ gg;@",
			"[$-10476]dddd\\,\\ \"die\"\\ d\\ mmmm\\ yyyy\\ gg;@",
			"[$-10476]\"die\"\\ d\\ mmmm\\ yyyy\\ gg;@"
		],
		"1143": [
			"yyyy\\-mm\\-dd;@",
			"[$-10477]dd/mm/yyyy;@",
			"[$-10477]dd/mm/yy;@",
			"[$-10477]dd\\-mmm\\-yyyy;@",
			"[$-10477]dddd\\,\\ mmmm\\ d\\,\\ yyyy;@",
			"[$-10477]mmmm\\ d\\,\\ yyyy;@"
		],
		"1144": [
			"[$-10478]yyyy/m/d;@",
			"[$-10478]yyyy\\-m\\-d;@",
			"[$-10478]yyyy\\.m\\.d;@",
			"[$-10478]yyyy\\.mm\\.dd;@",
			"[$-10478]yyyy\\-mm\\-dd;@",
			"[$-10478]yyyy/mm/dd;@",
			"[$-10478]yyyy\"ꈎ\"\\ m\"ꆪ\"\\ d\"ꑍ\";@",
			"[$-10478]dddd\\,\\ yyyy\"ꈎ\"\\ m\"ꆪ\"\\ d\"ꑍ\";@",
			"[$-10478]yyyy\"ꈎ\"\\ m\"ꆪ\"\\ d\"ꑍ\"\\,\\ dddd;@",
			"[$-10478]yyyy\\ꈎ\\ mmm\\ d\\ꑍ;@",
			"[$-10478]dddd\\,\\ yyyy\\ꈎ\\ mmm\\ d\\ꑍ;@"
		],
		"1145": [
			"yyyy\\-mm\\-dd;@",
			"[$-10479]d\\-m\\-yyyy;@",
			"[$-10479]d\\ mmm\\ yyyy;@",
			"[$-10479]d\\ mmm\\ yy;@",
			"[$-10479]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10479]d\\ mmmm\\ yyyy;@"
		],
		"1146": [
			"[$-1047A]dd\\-mm\\-yyyy;@",
			"[$-1047A]dd\\-mm\\-yy;@",
			"[$-1047A]dd/mm/yy;@",
			"[$-1047A]d/m/yy;@",
			"[$-1047A]yyyy\\-mm\\-dd;@",
			"[$-1047A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-1047A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-1047A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"1148": [
			"[$-1047C]m/d/yyyy;@",
			"[$-1047C]m/d/yy;@",
			"[$-1047C]mm/dd/yy;@",
			"[$-1047C]mm/dd/yyyy;@",
			"[$-1047C]yy/mm/dd;@",
			"[$-1047C]yyyy\\-mm\\-dd;@",
			"[$-1047C]dd\\-mmm\\-yy;@",
			"[$-1047C]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-1047C]mmmm\\ dd\\,\\ yyyy;@",
			"[$-1047C]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-1047C]dd\\ mmmm\\,\\ yyyy;@"
		],
		"1150": [
			"yyyy\\-mm\\-dd;@",
			"[$-1047E]dd/mm/yyyy;@",
			"[$-1047E]d\\ mmm\\ yyyy;@",
			"[$-1047E]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1047E]d\\ mmmm\\ yyyy;@"
		],
		"1152": [
			"[$-10480]yyyy\\-m\\-d;@",
			"[$-10480]yyyy\\.m\\.d;@",
			"[$-10480]yyyy\\-mm\\-dd;@",
			"[$-10480]yyyy\\.mm\\.dd;@",
			"[$-10480]yyyy\\-\"يىل\"\\ d\\-mmmm;@",
			"[$-10480]yyyy\\-\"يىل\"\\ d\\-mmmm\\ dddd;@",
			"[$-10480]yyyy\\-\"يىلى\"\\ mmm\"نىڭ\"\\ d\"-كۈنى\";@",
			"[$-10480]yyyy\\-\"يىلى\"\\ mmm\"نىڭ\"\\ d\"-كۈنى\"\\ dddd;@",
			"[$-10480]yyyy\\-m\\-d\\ dddd;@"
		],
		"1153": [
			"yyyy\\-mm\\-dd;@",
			"[$-10481]dd\\-mm\\-yyyy;@",
			"[$-10481]d\\ mmm\\ yyyy;@",
			"[$-10481]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-10481]d\\ mmmm\\ yyyy;@"
		],
		"1154": [
			"yyyy\\-mm\\-dd;@",
			"[$-10482]d/mm/yyyy;@",
			"[$-10482]d/mm/yy;@",
			"[$-10482]d\\ mmm\\ yyyy;@",
			"[$-10482]dddd\\ d\\ mmmm\" de \"yyyy;@",
			"[$-10482]dddd\\ d\\ mmmm\\ \"de\"\\ yyyy;@",
			"[$-10482]d\\ mmmm\\ \"de\"\\ yyyy;@"
		],
		"1155": [
			"yyyy\\-mm\\-dd;@",
			"[$-10483]dd/mm/yyyy;@",
			"[$-10483]dd/mm/yy;@",
			"[$-10483]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10483]d\\ mmm\\ yy;@",
			"[$-10483]d\\ mmmm\\ yyyy;@"
		],
		"1156": [
			"[$-10484]dd/mm/yyyy;@",
			"[$-10484]dd/mm/yy;@",
			"[$-10484]dd\\.mm\\.yy;@",
			"[$-10484]dd\\-mm\\-yy;@",
			"[$-10484]yyyy\\-mm\\-dd;@",
			"[$-10484]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10484]d\\ mmm\\ yy;@",
			"[$-10484]d\\ mmmm\\ yyyy;@"
		],
		"1157": [
			"[$-10485]dd\\.mm\\.yyyy;@",
			"[$-10485]d\\.m\\.yyyy;@",
			"[$-10485]yyyy\\-mm\\-dd;@",
			"[$-10485]yyyy\\ mm\\ d;@",
			"[$-10485]dd\\ yyyy\\ mm\\ d;@",
			"[$-10485]dddd\\,\\ yyyy\\ \"с.\"\\ mmmm\\ d\\ \"күнэ\";@",
			"[$-10485]yyyy\\ \"с.\"\\ mmmm\\ d\\ \"күнэ\";@",
			"[$-10485]dddd\\,\\ mmmm\\ d\\ \"күнэ\"\\ yyyy\\ \"с.\";@"
		],
		"1158": [
			"yyyy\\-mm\\-dd;@",
			"[$-10486]dd/mm/yyyy;@",
			"[$-10486]d/mm/yyyy;@",
			"[$-10486]dddd\\,\\ dd\" rech \"mmmm\" rech \"yyyy;@",
			"[$-10486]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-10486]d\" de \"mmmm\" de \"yyyy;@"
		],
		"1159": [
			"[$-10487]yyyy\\-mm\\-dd;@",
			"[$-10487]yyyy\\ mmm\\ d;@",
			"[$-10487]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10487]yyyy\\ mmmm\\ d;@"
		],
		"1160": [
			"yyyy\\-mm\\-dd;@",
			"[$-10488]dd\\-mm\\-yyyy;@",
			"[$-10488]d\\ mmm\\,\\ yyyy;@",
			"[$-10488]dddd\\,\\ d\\ mmm\\,\\ yyyy;@",
			"[$-10488]d\\ mmmm\\,\\ yyyy;@"
		],
		"1164": [
			"[$-16048C]yyyy/m/d;@",
			"[$-16048C]yyyy\\-mm\\-dd;@",
			"[$-16048C]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-16048C]d\\ mmmm\\ yyyy;@"
		],
		"1169": [
			"yyyy\\-mm\\-dd;@",
			"[$-10491]dd/mm/yyyy;@",
			"[$-10491]d\\ mmm\\ yyyy;@",
			"[$-10491]dd\\ mmmm\\ yyyy;@",
			"[$-10491]dddd\\,\\ d\"mh\"\\ mmmm\\ yyyy;@",
			"[$-10491]d\"mh\"\\ mmmm\\ yyyy;@"
		],
		"2052": [
			"yyyy\\-mm\\-dd;@",
			"[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@",
			"[DBNum1][$-804]yyyy\"年\"m\"月\";@",
			"[DBNum1][$-804]m\"月\"d\"日\";@",
			"yyyy\"年\"m\"月\"d\"日\";@",
			"yyyy\"年\"m\"月\";@",
			"m\"月\"d\"日\";@",
			"[$-804]aaaa;@",
			"[$-804]aaa;@",
			"yyyy/m/d;@",
			"[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@",
			"yyyy/m/d\\ h:mm;@",
			"yy/m/d;@",
			"m/d;@",
			"m/d/yy;@",
			"mm/dd/yy;@",
			"[$-409]d\\-mmm;@",
			"[$-409]d\\-mmm\\-yy;@",
			"[$-409]dd\\-mmm\\-yy;@",
			"[$-409]mmm\\-yy;@",
			"[$-409]mmmm\\-yy;@",
			"[$-409]mmmmm;@",
			"[$-409]mmmmm\\-yy;@"
		],
		"2055": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.mm\\.yy;@",
			"dd\\.\\ m\\.\\ yy;@",
			"d\\.m\\.yy;@",
			"dd\\.mm\\.yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-807]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-807]d\\.\\ mmmm\\ yyyy;@",
			"[$-807]d\\.\\ mmm\\ yy;@"
		],
		"2057": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-809]dd\\ mmmm\\ yyyy;@",
			"[$-809]d\\ mmmm\\ yyyy;@"
		],
		"2058": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-80A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-80A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-80A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"2060": [
			"d/mm/yyyy;@",
			"d/mm/yy;@",
			"dd\\.mm\\.yy;@",
			"yy/mm/dd;@",
			"dd\\-mm\\-yy;@",
			"dd/mm/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-80C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-80C]d\\ mmmm\\ yyyy;@",
			"[$-80C]dd\\-mmm\\-yy;@"
		],
		"2064": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.\\ mm\\.\\ yy;@",
			"d/m/yy;@",
			"dd\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-810]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-810]d\\-mmm\\-yy;@",
			"[$-810]d\\ mmmm\\ yyyy;@"
		],
		"2067": [
			"d/mm/yyyy;@",
			"d/mm/yy;@",
			"dd\\-mm\\-yy;@",
			"dd\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-813]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-813]dd\\-mmm\\-yy;@",
			"[$-813]d\\ mmmm\\ yyyy;@",
			"[$-813]dd\\ mmm\\ yy;@"
		],
		"2068": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-814]d\\.\\ mmmm\\ yyyy;@",
			"[$-814]dd\\.\\ mmmm\\ yyyy;@"
		],
		"2070": [
			"yyyy\\-mm\\-dd;@",
			"dd\\-mm\\-yyyy;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"[$-816]d/mmm;@",
			"[$-816]d\\-mmm\\-yy;@",
			"[$-816]dd\\-mmm\\-yy;@",
			"[$-816]mmm/yy;@",
			"[$-816]mmmm\\ yy;@",
			"[$-816]d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy;@",
			"[$-409]d/m/yy\\ h:mm\\ AM/PM;@",
			"d/m/yy\\ h:mm;@",
			"[$-816]mmmmm;@",
			"[$-816]mmmmm\\-yy;@",
			"d/m/yyyy;@",
			"[$-816]d\\-mmm\\-yyyy;@"
		],
		"2072": [
			"yyyy\\-mm\\-dd;@",
			"[$-10818]dd\\.mm\\.yyyy;@",
			"[$-10818]d\\ mmm\\ yyyy;@",
			"[$-10818]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-10818]d\\ mmmm\\ yyyy;@"
		],
		"2073": [
			"yyyy\\-mm\\-dd;@",
			"[$-10819]dd\\.mm\\.yyyy;@",
			"[$-10819]d\\ mmm\\ yyyy\\ \"г\"\\.;@",
			"[$]dddd\\,\\ d\\ mmmm\\ yyyy\\ \"г\"\\.;@",
			"[$]d\\ mmmm\\ yyyy\\ \"г\"\\.;@"
		],
		"2074": [
			"d\\.m\\.yyyy;@",
			"d\\.m\\.yy;@",
			"d\\.\\ m\\.\\ yyyy;@",
			"dd\\.mm\\.yyyy;@",
			"d\\.\\ m\\.\\ yy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.\\ mm\\.\\ yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-81A]d\\.\\ mmmm\\ yyyy;@",
			"[$-81A]dd\\.\\ mmmm\\ yyyy;@",
			"[$-81A]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@"
		],
		"2077": [
			"d\\.m\\.yyyy;@",
			"dd\\.mm\\.yyyy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-81D]\"den \"d\\ mmmm\\ yyyy;@"
		],
		"2092": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-82C]d\\ mmmm\\ yyyy;@",
			"[$-82C]dd\\ mmmm\\ yyyy;@"
		],
		"2094": [
			"[$-1082E]d\\.\\ m\\.\\ yyyy;@",
			"[$-1082E]d\\.\\ m\\.\\ yy;@",
			"[$-1082E]dd\\.mm\\.yyyy;@",
			"[$-1082E]dd\\.mm\\.yy;@",
			"[$-1082E]yyyy\\-mm\\-dd;@",
			"[$-1082E]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-1082E]d\\.\\ mmmm\\ yyyy;@"
		],
		"2098": [
			"[$-10832]yyyy\\-mm\\-dd;@",
			"[$-10832]yyyy\\ mmm\\ d;@",
			"[$-10832]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10832]yyyy\\ mmmm\\ d;@"
		],
		"2107": [
			"[$-1083B]yyyy\\-mm\\-dd;@",
			"[$-1083B]yy\\-mm\\-dd;@",
			"[$-1083B]dddd\\,\\ mmmm\\ d\". b. \"yyyy;@",
			"[$-1083B]mmmm\\ d\". b. \"yyyy;@"
		],
		"2108": [
			"yyyy\\-mm\\-dd;@",
			"[$-1083C]dd/mm/yyyy;@",
			"[$-1083C]d\\ mmm\\ yyyy;@",
			"[$-1083C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1083C]d\\ mmmm\\ yyyy;@"
		],
		"2110": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-83E]dd\\ mmmm\\ yyyy;@"
		],
		"2115": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"dd/mm\\ yyyy;@",
			"d\\.m\\.yy;@",
			"dd/mm/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-843]yyyy\\ \"йил\"\\ d\\-mmmm;@",
			"[$-843]d\\ mmmm\\ yyyy;@",
			"[$-843]dd\\ mmmm\\ yyyy;@"
		],
		"2117": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"d\\-m\\-yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-845]dd\\ mmmm\\ yyyy;@",
			"[$-845]d\\ mmmm\\ yyyy;@",
			"[$-5000845]dd\\-mm\\-yyyy;@",
			"[$-5000845]dd\\-mm\\-yy;@",
			"[$-5000845]d\\-m\\-yy;@",
			"[$-5000845]d\\.m\\.yy;@",
			"[$-5000845]yyyy\\-mm\\-dd;@",
			"[$-5000845]d/m/yyyy;@",
			"[$-5000845]dd\\ mmmm\\ yyyy;@",
			"[$-5000845]d\\ mmmm\\ yyyy;@",
			"[$-5000845]d\\ mmmm\\,\\ yyyy;@",
			"[$-5000845]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@"
		],
		"2118": [
			"[$-10846]dd\\-mm\\-yy;@",
			"[$-10846]d\\-m\\-yy;@",
			"[$-10846]d\\.m\\.yy;@",
			"[$-10846]dd\\-mm\\-yyyy;@",
			"[$-10846]yyyy\\-mm\\-dd;@",
			"[$-10846]dd\\ mmmm\\ yyyy\\ dddd;@",
			"[$-10846]d\\ mmmm\\ yyyy;@"
		],
		"2128": [
			"[$-10850]yyyy/m/d;@",
			"[$-10850]yyyy\\-m\\-d;@",
			"[$-10850]yyyy\\.m\\.d;@",
			"[$-10850]yyyy\\.mm\\.dd;@",
			"[$-10850]yyyy\\-mm\\-dd;@",
			"[$-10850]yyyy/mm/dd;@",
			"[$-10850]yy\\-m\\-d;@",
			"[$-10850]yy/m/d;@",
			"[$-10850]yy\\.m\\.d;@",
			"[$-10850]yy/mm/dd;@",
			"[$-10850]yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd;@",
			"[$-10850]yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ;@"
		],
		"2137": [
			"[$-10859]dd/mm/yyyy;@",
			"[$-10859]dd/mm/yy;@",
			"[$-10859]yyyy\\-mm\\-dd;@",
			"[$-10859]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-10859]dd\\ mmmm\\ yyyy;@"
		],
		"2141": [
			"[$-1085D]d/mm/yyyy;@",
			"[$-1085D]d/m/yy;@",
			"[$-1085D]dd/mm/yyyy;@",
			"[$-1085D]yy\\-mm\\-dd;@",
			"[$-1085D]yyyy\\-mm\\-dd;@",
			"[$-1085D]dd\\-mmm\\-yy;@",
			"[$-1085D]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-1085D]ddd\\,\\ mmmm\\ dd\\,yyyy;@",
			"[$-1085D]mmmm\\ dd\\,yyyy;@",
			"[$-1085D]dd\\ mmmm\\,\\ yyyy;@"
		],
		"2143": [
			"[$-1085F]dd\\-mm\\-yyyy;@",
			"[$-1085F]dd\\-mm\\-yy;@",
			"[$-1085F]yyyy\\-mm\\-dd;@",
			"[$-1085F]dd\\ mmmm\\,\\ yyyy;@",
			"[$-1085F]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@"
		],
		"2144": [
			"yyyy\\-mm\\-dd;@",
			"[$-10860]d/m/yyyy;@",
			"[$-10860]d/m/yy;@",
			"[$-10860]d\\ mmm\\ yyyy;@",
			"[$-10860]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-10860]d\\ mmmm\\ yyyy;@"
		],
		"2145": [
			"yyyy\\-mm\\-dd;@",
			"[$-10861]yyyy/m/d;@",
			"[$-10861]yy/m/d;@",
			"[$-10861]yyyy\\ mmm\\ d;@",
			"[$-10861]yyyy\\ mmmm\\ d\\,\\ dddd;@",
			"[$-10861]yyyy\\ mmmm\\ d;@"
		],
		"2151": [
			"[$-10867]dd/mm/yyyy;@",
			"[$-10867]dd/mm/yy;@",
			"[$-10867]dd\\.mm\\.yy;@",
			"[$-10867]dd\\-mm\\-yy;@",
			"[$-10867]yyyy\\-mm\\-dd;@",
			"[$-10867]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-10867]d\\ mmm\\ yy;@",
			"[$-10867]d\\ mmmm\\ yyyy;@"
		],
		"2155": [
			"[$-1086B]dd/mm/yyyy;@",
			"[$-1086B]dd/mm/yy;@",
			"[$-1086B]d/m/yy;@",
			"[$-1086B]dd\\-mm\\-yy;@",
			"[$-1086B]yyyy\\-mm\\-dd;@",
			"[$-1086B]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-1086B]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-1086B]d\" de \"mmmm\" de \"yyyy;@"
		],
		"2163": [
			"yyyy\\-mm\\-dd;@",
			"[$-10873]dd/mm/yyyy;@",
			"[$-10873]dd/mm/yy;@",
			"[$-10873]d\\ mmm\\ yyyy;@",
			"[$-10873]dddd\\፣\\ d\\ mmmm\\ yyyy;@",
			"[$-10873]d\\ mmmm\\ yyyy;@"
		],
		"3073": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"3076": [
			"d/m/yyyy;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"yy/m/d;@",
			"yy/mm/dd;@",
			"yyyy/m/d;@",
			"yyyy/mm/dd;@",
			"yyyy\\-mm\\-dd;@",
			"[$-C04]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@",
			"[$-C04]d\\ mmmm\\,\\ yyyy;@",
			"[$-C04]dddd\\ yyyy\\ mm\\ dd;@",
			"yyyy\\ mm\\ dd;@"
		],
		"3079": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.m\\.yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-C07]dddd\\,\\ dd\\.\\ mmmm\\ yyyy;@",
			"[$-C07]d\\.mmmm\\ yyyy;@",
			"[$-C07]d\\.mmmyyyy;@",
			"[$-C07]d\\ mmm\\ yyyy;@"
		],
		"3081": [
			"d/mm/yyyy;@",
			"d/mm/yy;@",
			"d/m/yy;@",
			"d/m/yyyy;@",
			"dd/mm/yy;@",
			"dd/mm/yyyy;@",
			"[$-C09]dd\\-mmm\\-yy;@",
			"[$-C09]dd\\-mmmm\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"yy/mm/dd;@",
			"yyyy/mm/dd;@",
			"[$-C09]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-C09]d\\ mmmm\\ yyyy;@"
		],
		"3082": [
			"yyyy\\-mm\\-dd;@",
			"d\\-m;@",
			"d\\-m\\-yy;@",
			"dd\\-mm\\-yy;@",
			"[$-C0A]d\\-mmm;@",
			"[$-C0A]d\\-mmm\\-yy;@",
			"[$-C0A]dd\\-mmm\\-yy;@",
			"[$-C0A]mmm\\-yy;@",
			"[$-C0A]mmmm\\-yy;@",
			"[$-C0A]d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy;@",
			"[$-409]d\\-m\\-yy\\ h:mm\\ AM/PM;@",
			"d\\-m\\-yy\\ h:mm;@",
			"[$-C0A]mmmmm;@",
			"[$-C0A]mmmmm\\-yy;@",
			"d\\-m\\-yyyy;@",
			"[$-C0A]d\\-mmm\\-yyyy;@"
		],
		"3084": [
			"yyyy\\-mm\\-dd;@",
			"yy\\-mm\\-dd;@",
			"dd\\-mm\\-yy;@",
			"yy\\ mm\\ dd;@",
			"dd/mm/yy;@",
			"[$-C0C]d\\ mmmm\\,\\ yyyy;@",
			"[$-C0C]d\\ mmm\\ yyyy;@"
		],
		"3098": [
			"d\\.m\\.yyyy\\.;@",
			"d\\.m\\.yy\\.;@",
			"d\\.\\ m\\.\\ yyyy\\.;@",
			"dd\\.mm\\.yyyy\\.;@",
			"d\\.\\ m\\.\\ yy\\.;@",
			"dd\\.mm\\.yy\\.;@",
			"dd\\.\\ mm\\.\\ yy\\.;@",
			"yyyy\\-mm\\-dd;@",
			"[$-C1A]d\\.\\ mmmm\\ yyyy\\.;@",
			"[$-C1A]dd\\.\\ mmmm\\ yyyy\\.;@",
			"[$-C1A]dddd\\,\\ d\\.\\ mmmm\\ yyyy\\.;@"
		],
		"3131": [
			"[$-10C3B]d\\.m\\.yyyy;@",
			"[$-10C3B]dd\\.mm\\.yyyy;@",
			"[$-10C3B]d\\.m\\.yy;@",
			"[$-10C3B]yyyy\\-mm\\-dd;@",
			"[$-10C3B]dddd\", \"mmmm\\ d\". b. \"yyyy;@",
			"[$-10C3B]mmmm\\ d\". b. \"yyyy;@"
		],
		"3152": [
			"[$-10C50]yyyy/m/d;@",
			"[$-10C50]yyyy\\-m\\-d;@",
			"[$-10C50]yyyy\\.m\\.d;@",
			"[$-10C50]yyyy\\.mm\\.dd;@",
			"[$-10C50]yyyy\\-mm\\-dd;@",
			"[$-10C50]yyyy/mm/dd;@",
			"[$-10C50]yy\\-m\\-d;@",
			"[$-10C50]yy/m/d;@",
			"[$-10C50]yy\\.m\\.d;@",
			"[$-10C50]yy/mm/dd;@",
			"[$-10C50]yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd;@",
			"[$-10C50]yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ;@"
		],
		"3179": [
			"[$-10C6B]dd/mm/yyyy;@",
			"[$-10C6B]dd/mm/yy;@",
			"[$-10C6B]d/m/yy;@",
			"[$-10C6B]dd\\-mm\\-yy;@",
			"[$-10C6B]yyyy\\-mm\\-dd;@",
			"[$-10C6B]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@"
		],
		"4097": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"4100": [
			"d/m/yyyy;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"yy/m/d;@",
			"yy/mm/dd;@",
			"yyyy/m/d;@",
			"yyyy/mm/dd;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@",
			"[$-1004]d\\ mmmm\\,\\ yyyy;@",
			"[$-1004]dddd\\ yyyy\\ mm\\ dd;@",
			"yyyy\\ mm\\ dd;@"
		],
		"4103": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.mm\\.yy;@",
			"d\\.m\\.yy;@",
			"d\\.m\\.yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1007]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-1007]d\\.\\ mmmm\\ yyyy;@",
			"[$-1007]d\\.\\ mmm\\ yyyy;@"
		],
		"4105": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"yyyy\\-mm\\-dd;@",
			"yy\\-mm\\-dd;@",
			"m/dd/yy;@",
			"[$-1009]mmmm\\ d\\,\\ yyyy;@",
			"[$-1009]d\\-mmm\\-yy;@"
		],
		"4106": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/mm/yyyy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-100A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-100A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-100A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"4108": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.\\ m\\.\\ yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-100C]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-100C]d\\.\\ mmmm\\ yyyy;@",
			"[$-100C]d\\ mmm\\ yy;@"
		],
		"4122": [
			"yyyy\\-mm\\-dd;@",
			"[$-1101A]d\\.\\ m\\.\\ yyyy\\.;@",
			"[$-1101A]d\\.\\ m\\.\\ yy\\.;@",
			"[$-1101A]d\\.\\ mmm\\ yyyy\\.;@",
			"[$-1101A]dddd\\,\\ d\\.\\ mmmm\\ yyyy\\.;@",
			"[$-1101A]d\\.\\ mmmm\\ yyyy\\.;@"
		],
		"4155": [
			"[$-1103B]dd\\.mm\\.yyyy;@",
			"[$-1103B]dd\\.mm\\.yy;@",
			"[$-1103B]d\\.m\\.yy;@",
			"[$-1103B]yyyy\\-mm\\-dd;@",
			"[$-1103B]dddd\\,\\ mmmm\\ d\". b. \"yyyy;@",
			"[$-1103B]mmmm\\ d\". b. \"yyyy;@"
		],
		"4191": [
			"yyyy\\-mm\\-dd;@",
			"dd\\-mm;@",
			"dd\\-mm\\-yyyy;@",
			"dd\\.mmm\\.yyyy;@",
			"[$-105F]d\\-mmm;@",
			"[$-105F]d\\-mmm\\-yy;@",
			"[$-105F]dd\\-mmm\\-yy;@",
			"[$-105F]mmm\\-yy;@",
			"[$-105F]mmmm\\-yy;@",
			"[$-105F]dd\\ mmmm\\,\\ yyyy;@",
			"[$-105F]dd\\-mm\\-yy\\ h:mm;@",
			"dd\\-mm\\-yy\\ h:mm;@",
			"[$-105F]mmmmm;@",
			"[$-105F]mmmmm\\,\\ yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-105F]dd\\.mmm\\.yyyy;@"
		],
		"5121": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"5124": [
			"d/m/yyyy;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"yy/m/d;@",
			"yy/mm/dd;@",
			"yyyy/m/d;@",
			"yyyy/mm/dd;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1404]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@",
			"[$-1404]d\\ mmmm\\,\\ yyyy;@",
			"[$-1404]dddd\\ yyyy\\ mm\\ dd;@",
			"yyyy\\ mm\\ dd;@"
		],
		"5127": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"d\\.mm\\.yy;@",
			"dd\\.\\ m\\.\\ yy;@",
			"d\\.m\\.yy;@",
			"dd\\.mm\\.yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1407]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@",
			"[$-1407]d\\.\\ mmmm\\ yyyy;@",
			"[$-1407]d\\.\\ mmm\\ yy;@"
		],
		"5129": [
			"d/mm/yyyy;@",
			"d/mm/yy;@",
			"dd/mm/yy;@",
			"d\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1409]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-1409]d\\ mmmm\\ yyyy;@"
		],
		"5130": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-140A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-140A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-140A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"5132": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"dd\\.mm\\.yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-140C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-140C]d\\ mmm\\ yy;@",
			"[$-140C]d\\ mmmm\\ yyyy;@"
		],
		"5146": [
			"yyyy\\-mm\\-dd;@",
			"[$-1141A]d\\.\\ m\\.\\ yyyy\\.;@",
			"[$-1141A]d\\.\\ mmm\\ yyyy\\.;@",
			"[$-1141A]dddd\\,\\ d\\.\\ mmmm\\ yyyy\\.;@",
			"[$-1141A]d\\.\\ mmmm\\ yyyy\\.;@"
		],
		"5179": [
			"[$-1143B]yyyy\\-mm\\-dd;@",
			"[$-1143B]yy\\-mm\\-dd;@",
			"[$-1143B]dddd\\,\\ mmmm\\ d\". b. \"yyyy;@",
			"[$-1143B]mmmm\\ d\". b. \"yyyy;@"
		],
		"6153": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"d\\.m\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1809]dd\\ mmmm\\ yyyy;@",
			"[$-1809]d\\ mmmm\\ yyyy;@"
		],
		"6154": [
			"mm/dd/yyyy;@",
			"mm/dd/yy;@",
			"d/m/yy;@",
			"dd/mm/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-180A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-180A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-180A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"6156": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"dd\\.mm\\.yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-180C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-180C]d\\ mmm\\ yy;@",
			"[$-180C]d\\ mmmm\\ yyyy;@"
		],
		"6170": [
			"yyyy\\-mm\\-dd;@",
			"[$-1181A]d\\.m\\.yyyy\\.;@",
			"[$-1181A]d\\.m\\.yy\\.;@",
			"[$-1181A]d\\.\\ m\\.\\ yyyy\\.;@",
			"[$-1181A]d\\.\\ mmmm\\ yyyy\\.;@",
			"[$-1181A]dddd\\,\\ d\\.\\ mmmm\\ yyyy\\.;@"
		],
		"6203": [
			"[$-1183B]dd\\.mm\\.yyyy;@",
			"[$-1183B]dd\\.mm\\.yy;@",
			"[$-1183B]d\\.m\\.yy;@",
			"[$-1183B]yyyy\\-mm\\-dd;@",
			"[$-1183B]dddd\\,\\ mmmm\\ d\". b. \"yyyy;@",
			"[$-1183B]mmmm\\ d\". b. \"yyyy;@"
		],
		"7169": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"7177": [
			"yyyy/mm/dd;@",
			"yy/mm/dd;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1C09]dd\\ mmmm\\ yyyy;@"
		],
		"7178": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"mm/dd/yyyy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1C0A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-1C0A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-1C0A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"7180": [
			"yyyy\\-mm\\-dd;@",
			"[$-11C0C]dd/mm/yyyy;@",
			"[$-11C0C]d\\ mmm\\ yyyy;@",
			"[$-11C0C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-11C0C]d\\ mmmm\\ yyyy;@"
		],
		"7194": [
			"d\\.m\\.yyyy;@",
			"d\\.m\\.yy;@",
			"d\\.\\ m\\.\\ yyyy;@",
			"dd\\.mm\\.yyyy;@",
			"d\\.\\ m\\.\\ yy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.\\ mm\\.\\ yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-1C1A]d\\.\\ mmmm\\ yyyy;@",
			"[$-1C1A]dd\\.\\ mmmm\\ yyyy;@",
			"[$-1C1A]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@"
		],
		"7227": [
			"[$-11C3B]yyyy\\-mm\\-dd;@",
			"[$-11C3B]yy\\-mm\\-dd;@",
			"[$-11C3B]dddd\\,\\ mmmm\\ d\". b. \"yyyy;@",
			"[$-11C3B]mmmm\\ d\". b. \"yyyy;@"
		],
		"8193": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"8201": [
			"dd/mm/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-2009]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-2009]mmmm\\ dd\\,\\ yyyy;@",
			"[$-2009]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-2009]dd\\ mmmm\\,\\ yyyy;@"
		],
		"8202": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-200A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-200A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-200A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"8204": [
			"yyyy\\-mm\\-dd;@",
			"[$-1200C]dd/mm/yyyy;@",
			"[$-1200C]d\\ mmm\\ yyyy;@",
			"[$-1200C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1200C]d\\ mmmm\\ yyyy;@"
		],
		"8218": [
			"[$-1201A]d\\.m\\.yyyy;@",
			"[$-1201A]d\\.m\\.yy;@",
			"[$-1201A]d\\.\\ m\\.\\ yyyy;@",
			"[$-1201A]dd\\.mm\\.yyyy;@",
			"[$-1201A]d\\.\\ m\\.\\ yy;@",
			"[$-1201A]dd\\.mm\\.yy;@",
			"[$-1201A]dd\\.\\ mm\\.\\ yy;@",
			"[$-1201A]yyyy\\-mm\\-dd;@",
			"[$-1201A]d\\.\\ mmmm\\ yyyy;@",
			"[$-1201A]dd\\.\\ mmmm\\ yyyy;@",
			"[$-1201A]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@"
		],
		"8251": [
			"[$-1203B]d\\.m\\.yyyy;@",
			"[$-1203B]dd\\.mm\\.yyyy;@",
			"[$-1203B]d\\.m\\.yy;@",
			"[$-1203B]yyyy\\-mm\\-dd;@",
			"[$-1203B]mmmm\\ d\". p. \"yyyy;@",
			"[$-1203B]dddd\\,\\ mmmm\\ d\". p. \"yyyy;@"
		],
		"9217": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"9225": [
			"mm/dd/yyyy;@",
			"mm/dd/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-2409]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-2409]mmmm\\ dd\\,\\ yyyy;@",
			"[$-2409]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-2409]dd\\ mmmm\\,\\ yyyy;@"
		],
		"9226": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/mm/yyyy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-240A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-240A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-240A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"9228": [
			"yyyy\\-mm\\-dd;@",
			"[$-1240C]dd/mm/yyyy;@",
			"[$-1240C]d\\ mmm\\ yyyy;@",
			"[$-1240C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1240C]d\\ mmmm\\ yyyy;@"
		],
		"9242": [
			"d\\.m\\.yyyy;@",
			"d\\.m\\.yy;@",
			"d\\.\\ m\\.\\ yyyy;@",
			"dd\\.mm\\.yyyy;@",
			"d\\.\\ m\\.\\ yy;@",
			"dd\\.mm\\.yy;@",
			"dd\\.\\ mm\\.\\ yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-241A]d\\.\\ mmmm\\ yyyy;@",
			"[$-241A]dd\\.\\ mmmm\\ yyyy;@",
			"[$-241A]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@"
		],
		"9275": [
			"[$-1243B]d\\.m\\.yyyy;@",
			"[$-1243B]dd\\.mm\\.yyyy;@",
			"[$-1243B]d\\.m\\.yy;@",
			"[$-1243B]yyyy\\-mm\\-dd;@",
			"[$-1243B]mmmm\\ d\". p. \"yyyy;@",
			"[$-1243B]dddd\\,\\ mmmm\\ d\\.\\ yyyy;@"
		],
		"10241": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"10249": [
			"dd/mm/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-2809]dddd\\,\\ dd\\ mmmm\\ yyyy;@"
		],
		"10250": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-280A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-280A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-280A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"10252": [
			"yyyy\\-mm\\-dd;@",
			"[$-1280C]dd/mm/yyyy;@",
			"[$-1280C]d\\ mmm\\ yyyy;@",
			"[$-1280C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1280C]d\\ mmmm\\ yyyy;@"
		],
		"10266": [
			"d\\.m\\.yyyy\\.;@",
			"d\\.m\\.yy\\.;@",
			"d\\.\\ m\\.\\ yyyy\\.;@",
			"dd\\.mm\\.yyyy\\.;@",
			"d\\.\\ m\\.\\ yy\\.;@",
			"dd\\.mm\\.yy\\.;@",
			"dd\\.\\ mm\\.\\ yy\\.;@",
			"yyyy\\-mm\\-dd;@",
			"[$-281A]d\\.\\ mmmm\\ yyyy\\.;@",
			"[$-281A]dd\\.\\ mmmm\\ yyyy\\.;@",
			"[$-281A]dddd\\,\\ d\\.\\ mmmm\\ yyyy\\.;@"
		],
		"11265": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"11273": [
			"dd/mm/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-2C09]dddd\\,\\ dd\\ mmmm\\ yyyy;@"
		],
		"11274": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-2C0A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-2C0A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-2C0A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"11276": [
			"yyyy\\-mm\\-dd;@",
			"[$-12C0C]dd/mm/yyyy;@",
			"[$-12C0C]d\\ mmm\\ yyyy;@",
			"[$-12C0C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-12C0C]d\\ mmmm\\ yyyy;@"
		],
		"12289": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"12297": [
			"m/d/yyyy;@",
			"m/d/yy;@",
			"mm/dd/yy;@",
			"mm/dd/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"yy/mm/dd;@",
			"[$-3009]dd\\-mmm\\-yy;@",
			"[$-3009]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-3009]mmmm\\ dd\\,\\ yyyy;@",
			"[$-3009]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-3009]dd\\ mmmm\\,\\ yyyy;@"
		],
		"12298": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-300A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-300A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-300A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"12300": [
			"yyyy\\-mm\\-dd;@",
			"[$-1300C]dd/mm/yyyy;@",
			"[$-1300C]d\\ mmm\\ yyyy;@",
			"[$-1300C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1300C]d\\ mmmm\\ yyyy;@"
		],
		"13313": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"13321": [
			"m/d/yyyy;@",
			"m/d/yy;@",
			"mm/dd/yy;@",
			"mm/dd/yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"yy/mm/dd;@",
			"[$-3409]dd\\-mmm\\-yy;@",
			"[$-3409]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-3409]mmmm\\ dd\\,\\ yyyy;@",
			"[$-3409]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-3409]dd\\ mmmm\\,\\ yyyy;@"
		],
		"13322": [
			"dd\\-mm\\-yyyy;@",
			"dd\\-mm\\-yy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-340A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-340A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-340A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"13324": [
			"yyyy\\-mm\\-dd;@",
			"[$-1340C]dd/mm/yyyy;@",
			"[$-1340C]d\\ mmm\\ yyyy;@",
			"[$-1340C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1340C]d\\ mmmm\\ yyyy;@"
		],
		"14337": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"14345": [
			"yyyy\\-mm\\-dd;@",
			"[$-13809]dd/mm/yyyy;@",
			"[$-13809]dd/mm/yy;@",
			"[$-13809]d\\ mmm\\ yyyy;@",
			"[$-13809]dddd\\,\\ dd\\ mmmm\\ yyyy;@",
			"[$-13809]dd\\ mmmm\\ yyyy;@"
		],
		"14346": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-380A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-380A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"14348": [
			"yyyy\\-mm\\-dd;@",
			"[$-1380C]dd/mm/yyyy;@",
			"[$-1380C]d\\ mmm\\ yyyy;@",
			"[$-1380C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-1380C]d\\ mmmm\\ yyyy;@"
		],
		"15361": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"15369": [
			"yyyy\\-mm\\-dd;@",
			"[$-13C09]d/m/yyyy;@",
			"[$-13C09]d\\ mmm\\ yyyy;@",
			"[$-13C09]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-13C09]d\\ mmmm\\ yyyy;@"
		],
		"15370": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-3C0A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-3C0A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-3C0A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"15372": [
			"yyyy\\-mm\\-dd;@",
			"[$-13C0C]dd/mm/yyyy;@",
			"[$-13C0C]d\\ mmm\\ yyyy;@",
			"[$-13C0C]dddd\\ d\\ mmmm\\ yyyy;@",
			"[$-13C0C]d\\ mmmm\\ yyyy;@"
		],
		"16385": [
			"yyyy\\-mm\\-dd;@",
			"[$-1010000]d/m/yyyy;@",
			"[$-1010000]yyyy/mm/dd;@",
			"[$-1010401]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@",
			"[$-2010000]d/mm/yyyy;@",
			"[$-2010000]yyyy/mm/dd;@",
			"[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@"
		],
		"16393": [
			"[$-14009]dd\\-mm\\-yyyy;@",
			"[$-14009]dd\\-mm\\-yy;@",
			"[$-14009]d\\-m\\-yy;@",
			"[$-14009]d\\.m\\.yy;@",
			"[$-14009]yyyy\\-mm\\-dd;@",
			"[$-14009]dd\\ mmmm\\ yyyy;@",
			"[$-14009]d\\ mmmm\\ yyyy;@",
			"[$-14009]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@"
		],
		"16394": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-400A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-400A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-400A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"17417": [
			"[$-14409]d/m/yyyy;@",
			"[$-14409]d/m/yy;@",
			"[$-14409]dd/mm/yyyy;@",
			"[$-14409]dd/mm/yy;@",
			"[$-14409]yyyy\\-mm\\-dd;@",
			"[$-14409]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@",
			"[$-14409]d\\ mmmm\\,\\ yyyy;@"
		],
		"17418": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"mm\\-dd\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-440A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"
		],
		"18441": [
			"yyyy\\-mm\\-dd;@",
			"[$-14809]d/m/yyyy;@",
			"[$-14809]d/m/yy;@",
			"[$-14809]d\\ mmm\\ yyyy;@",
			"[$-14809]dddd\\,\\ d\\ mmmm\\ yyyy;@",
			"[$-14809]d\\ mmmm\\ yyyy;@"
		],
		"18442": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"mm\\-dd\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-480A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"
		],
		"19466": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"mm\\-dd\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-4C0A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"
		],
		"20490": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"mm\\-dd\\-yyyy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-500A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"
		],
		"21514": [
			"[$-1540A]m/d/yyyy;@",
			"[$-1540A]m/d/yy;@",
			"[$-1540A]mm/dd/yy;@",
			"[$-1540A]yy/mm/dd;@",
			"[$-1540A]yyyy\\-mm\\-dd;@",
			"[$-1540A]dd\\-mmm\\-yy;@",
			"[$-1540A]dddd\\,\\ mmmm\\ dd\\,\\ yyyy;@",
			"[$-1540A]mmmm\\ dd\\,\\ yyyy;@",
			"[$-1540A]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@",
			"[$-1540A]dd\\ mmmm\\,\\ yyyy;@"
		],
		"22538": [
			"dd/mm/yyyy;@",
			"dd/mm/yy;@",
			"d/mm/yy;@",
			"d/m/yy;@",
			"dd\\-mm\\-yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-580A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@",
			"[$-580A]dddd\\ d\" de \"mmmm\" de \"yyyy;@",
			"[$-580A]d\" de \"mmmm\" de \"yyyy;@"
		],
		"63488": [
			"mm-dd-yy",
			"[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
			"d\\.m\\.;@",
			"d\\.m\\.yy;@",
			"d\\.m\\.yyyy;@",
			"[$-40B]d\\.\\ mmmm\\t\\a;@",
			"[$-40B]d\\.\\ mmmm\\t\\a\\ yy;@",
			"[$-40B]d\\.\\ mmmm\\t\\a\\ yyyy;@",
			"[$-40B]mmmm\\ yy;@",
			"[$-40B]mmmm\\ yyyy;@",
			"[$-40B]d\\.\\ mmmm\\t\\a\\ yyyy\\ h:mm;@",
			"d\\.m\\.yyyy\\ h:mm;@",
			"d\\.m\\.yy\\ h:mm;@",
			"[$-40B]mmmmm;@",
			"[$-40B]mmmmm\\ yy;@",
			"yyyy\\-mm\\-dd;@",
			"yyyy\\-mm\\-dd\\ hh:mm;@"
		],
		"64546": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-FC22]d\\ mmmm\\ yyyy\" р.\";@"
		],
		"64547": [
			"dd\\.mm\\.yyyy;@",
			"dd\\.mm\\.yy;@",
			"yyyy\\-mm\\-dd;@",
			"[$-FC23]d\\ mmmm\\ yyyy;@"
		],
		"64551": [
			"yyyy\\-mm\\-dd;@",
			"yyyy\\.mm\\.dd;@",
			"[$-FC27]yyyy\\ \"m.\"\\ mmmm\\ d\\ \"d.\";@",
			"[$-427]yyyy\\ \"m.\"\\ mmmm\\ d\\ \"d.\";@"
		]
	};

	let c_oAscTimeFormatExcel = {
		"1025": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"1026": [
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"1027": [
			"h:mm:ss;@",
			"h:mm;@"
		],
		"1028": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]yyyy/m/d h:mm AM/PM;@",
			"yyyy/m/d h:mm;@",
			"h\"時\"mm\"分\";@",
			"h\"時\"mm\"分\"ss\"秒\";@",
			"上午/下午h\"時\"mm\"分\";@",
			"上午/下午h\"時\"mm\"分\"ss\"秒\";@"
		],
		"1029": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1030": [
			"hh:mm;@",
			"hh:mm:ss;@",
			"[$-409]hh:mm AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"dd-mm-yy hh:mm;@",
			"mm-dd-yy hh:mm:ss;@",
			"dd-mm-yy hh:mm:ss;@",
			"yyyy-mm-dd hh:mm;@"
		],
		"1031": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d.m.yy h:mm;@"
		],
		"1032": [
			"h:mm;@",
			"[$-408]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-408]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-408]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1033": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]m/d/yy h:mm AM/PM;@",
			"m/d/yy h:mm;@"
		],
		"1034": [
			"h:mm:ss;@",
			"hh:mm:ss;@",
			"hh:mm;@",
			"hh\"H\"mm\"'\";@"
		],
		"1035": [
			"[$-409]h\\.mm\\.ss AM/PM",
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d.m.yyyy h:mm AM/PM;@",
			"d.m.yyyy h:mm;@"
		],
		"1036": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1038": [
			"h:mm;@",
			"[$-40E]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-40E]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-40E]yyyy. m. d. h:mm AM/PM;@",
			"yyyy. m. d. h:mm;@",
			"[$-40E]h \"óra\" m \"perc\" AM/PM;@",
			"h \"óra\" m \"perc\";@",
			"[$-40E]h \"óra\" m \"perckor\" AM/PM;@"
		],
		"1039": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"hh:mm;@"
		],
		"1040": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1042": [
			"h:mm;@",
			"h:mm:ss;@",
			"[$-412]AM/PM h:mm;@",
			"[$-412]AM/PM h:mm:ss;@",
			"[$-409]h:mm AM/PM;@",
			"[$-409]h:mm:ss AM/PM;@",
			"yyyy\"-\"m\"-\"d h:mm;@",
			"[$-412]yyyy\"-\"m\"-\"d AM/PM h:mm;@",
			"[$-409]yyyy\"-\"m\"-\"d h:mm AM/PM;@",
			"h\"시\" mm\"분\";@",
			"h\"시\" mm\"분\" ss\"초\";@",
			"[$-412]AM/PM h\"시\" mm\"분\";@",
			"[$-412]AM/PM h\"시\" mm\"분\" ss\"초\";@"
		],
		"1043": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d-mm-yy h:mm AM/PM;@",
			"d-mm-yy h:mm;@"
		],
		"1044": [
			"hh:mm;@",
			"[$-409]h:mm AM/PM;@",
			"hh:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"kl 'hh.mm;@",
			"[h]:mm:ss;@",
			"[$-409]m/d/yy h:mm AM/PM;@",
			"m/d/yy hh:mm;@"
		],
		"1045": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]yy-mm-dd h:mm AM/PM;@",
			"yy-mm-dd h:mm;@"
		],
		"1046": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1047": [
			"[$-10417]hh:mm:ss;@",
			"[$-10417]hh:mm;@"
		],
		"1048": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1049": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]dd/mm/yy h:mm AM/PM;@",
			"dd/mm/yy h:mm;@"
		],
		"1050": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d.m.yy. h:mm AM/PM;@",
			"d.m.yy. h:mm;@"
		],
		"1051": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"1053": [
			"hh:mm;@",
			"hh:mm:ss;@",
			"\"kl \"hh:mm;@",
			"\"kl \"hh:mm:ss;@",
			"[$-409]yyyy-mm-dd h:mm AM/PM;@",
			"[h]:mm:ss;@",
			"yyyy-mm-dd hh:mm;@",
			"[$-409]yyyy-mm-dd h:mm AM/PM;@"
		],
		"1055": [
			"hh:mm;@",
			"hh:mm:ss;@",
			"mm:ss.0;@",
			"d/m/yy hh:mm;@",
			"dd/mm/yy hh:mm;@",
			"d/m/yyyy hh:mm;@",
			"m/d/yy hh:mm;@"
		],
		"1056": [
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1057": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1060": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1061": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1062": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1064": [
			"[$-10428]hh:mm:ss;@",
			"[$-10428]h:mm:ss;@",
			"[$-10428]hh:mm;@",
			"[$-10428]h:mm;@"
		],
		"1065": [
			"[$-1000000]hh:mm:ss;@",
			"[$-1000429]hh:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-3000000]h:mm:ss;@",
			"[$-3000429]h:mm AM/PM;@",
			"[$-3000409]h:mm AM/PM;@"
		],
		"1066": [
			"[$-1000000]h:mm;@",
			"[$-100042A]h:mm:ss AM/PM;@",
			"[$-1000409]h:mm:ss AM/PM;@",
			"[$-1000000]h:mm:ss;@"
		],
		"1067": [
			"h:mm:ss;@",
			"hh:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@"
		],
		"1068": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1069": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]yy/mm/dd h:mm AM/PM;@",
			"yy/mm/dd h:mm;@"
		],
		"1070": [
			"[$-1042E]h:mm:ss;@",
			"[$-1042E]h:mm \"hodź\".;@"
		],
		"1071": [
			"hh:mm:ss;@"
		],
		"1072": [
			"[$-10430]hh:mm:ss;@",
			"[$-10430]hh:mm;@"
		],
		"1073": [
			"[$-10431]hh:mm:ss;@",
			"[$-10431]hh:mm;@"
		],
		"1074": [
			"[$-10432]hh:mm:ss;@",
			"[$-10432]hh:mm;@"
		],
		"1075": [
			"[$-10433]hh:mm:ss;@",
			"[$-10433]hh:mm;@"
		],
		"1076": [
			"[$-10434]hh:mm:ss;@",
			"[$-10434]hh:mm;@"
		],
		"1077": [
			"[$-10435]hh:mm:ss;@",
			"[$-10435]hh:mm;@"
		],
		"1078": [
			"[$-409]hh:mm:ss AM/PM;@",
			"[$-409]h:mm:ss AM/PM;@"
		],
		"1079": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1080": [
			"hh.mm.ss;@",
			"hh:mm:ss;@"
		],
		"1081": [
			"[$-1000000]h:mm;@",
			"[$-4000000]h:mm;@",
			"[$-1000439]h:mm AM/PM;@",
			"[$-4000439]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-1000439]h:mm:ss AM/PM;@",
			"[$-4000439]h:mm:ss AM/PM;@",
			"[$-1000409]h:mm:ss AM/PM;@"
		],
		"1082": [
			"[$-1043A]hh:mm:ss;@",
			"[$-1043A]hh:mm;@"
		],
		"1083": [
			"[$-1043B]hh:mm:ss;@",
			"[$-1043B]hh:mm;@"
		],
		"1085": [
			"[$-1043D]hh:mm:ss;@",
			"[$-1043D]hh:mm;@"
		],
		"1086": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1087": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1088": [
			"h:mm:ss;@"
		],
		"1089": [
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1090": [
			"[$-10442]hh:mm:ss;@",
			"[$-10442]h:mm:ss;@",
			"[$-10442]hh:mm;@",
			"[$-10442]h:mm;@"
		],
		"1091": [
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"1092": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1093": [
			"[$-10445]hh.mm.ss;@",
			"[$-10445]h.mm.ss;@",
			"[$-10409]AM/PM hh.mm.ss;@",
			"[$-10409]AM/PM h.mm.ss;@",
			"[$-10445]hh.mm;@",
			"[$-10445]h.mm;@",
			"[$-10409]AM/PM hh.mm;@",
			"[$-10409]AM/PM h.mm;@"
		],
		"1094": [
			"[$-446]AM/PM hh:mm:ss;@",
			"[$-446]AM/PM h:mm:ss;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"1095": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-447]AM/PM hh:mm:ss;@",
			"[$-447]AM/PM h:mm:ss;@"
		],
		"1096": [
			"[$-10448]hh:mm:ss;@",
			"[$-10448]h:mm:ss;@",
			"[$-10409]AM/PM hh:mm:ss;@",
			"[$-10409]AM/PM h:mm:ss;@",
			"[$-10448]hh:mm;@",
			"[$-10448]h:mm;@",
			"[$-10409]AM/PM hh:mm;@",
			"[$-10409]AM/PM h:mm;@"
		],
		"1097": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-449]hh:mm:ss AM/PM;@",
			"[$-449]h:mm:ss AM/PM;@"
		],
		"1098": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-44A]AM/PM hh:mm:ss;@",
			"[$-44A]AM/PM h:mm:ss;@"
		],
		"1099": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-44B]AM/PM hh:mm:ss;@",
			"[$-44B]AM/PM h:mm:ss;@"
		],
		"1100": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-1044C]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-1044C]hh:mm;@"
		],
		"1101": [
			"[$-1044D]AM/PM h:mm:ss;@",
			"[$-1044D]AM/PM hh:mm:ss;@",
			"[$-1044D]h:mm:ss;@",
			"[$-1044D]AM/PM h:mm;@",
			"[$-1044D]AM/PM hh:mm;@",
			"[$-1044D]h:mm;@"
		],
		"1102": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-44E]hh:mm:ss AM/PM;@",
			"[$-44E]h:mm:ss AM/PM;@"
		],
		"1103": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-44F]hh:mm:ss AM/PM;@",
			"[$-44F]h:mm:ss AM/PM;@"
		],
		"1104": [
			"h:mm:ss;@"
		],
		"1105": [
			"[$-10451]hh:mm:ss;@",
			"[$-10451]hh:mm;@"
		],
		"1106": [
			"[$-10452]hh:mm:ss;@",
			"[$-10452]hh:mm;@"
		],
		"1107": [
			"[$-10453]hh:mm:ss;@",
			"[$-10453]h:mm;@"
		],
		"1108": [
			"[$-10454]h:mm:ss;@",
			"[$-10454]hh:mm:ss;@",
			"[$-10454]h:mm;@",
			"[$-10454]hh:mm;@"
		],
		"1109": [
			"[$-10455]hh:mm:ss;@",
			"[$-10455]h:mm;@",
			"[$-10455]hh:mm;@"
		],
		"1110": [
			"h:mm:ss;@",
			"hh:mm:ss;@",
			"hh:mm;@",
			"[$-456]hh:mm:ss AM/PM;@"
		],
		"1111": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-457]hh:mm:ss AM/PM;@",
			"[$-457]h:mm:ss AM/PM;@"
		],
		"1112": [
			"[$-10458]h:mm:ss AM/PM;@",
			"[$-10458]hh:mm:ss;@",
			"[$-10458]h:mm AM/PM;@",
			"[$-10458]hh:mm;@"
		],
		"1113": [
			"[$-10409]AM/PM h:mm:ss;@",
			"[$-10459]hh:mm:ss;@",
			"[$-10409]AM/PM h:mm;@",
			"[$-10459]hh:mm;@"
		],
		"1114": [
			"[$-45A]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"1115": [
			"[$-1045B]hh.mm.ss;@",
			"[$-1045B]hh.mm;@"
		],
		"1116": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-1045C]h:mm:ss;@",
			"[$-1045C]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-1045C]h:mm;@",
			"[$-1045C]hh:mm;@"
		],
		"1117": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-1045D]h:mm:ss;@",
			"[$-1045D]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-1045D]h:mm;@",
			"[$-1045D]hh:mm;@"
		],
		"1118": [
			"[$-1045E]h:mm:ss AM/PM;@",
			"[$-1045E]hh:mm:ss;@",
			"[$-1045E]h:mm AM/PM;@",
			"[$-1045E]hh:mm;@"
		],
		"1119": [
			"[$-1045F]hh:mm:ss;@",
			"[$-1045F]hh:mm;@"
		],
		"1120": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10460]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10460]hh:mm;@"
		],
		"1121": [
			"[$-10461]h:mm:ss AM/PM;@",
			"[$-10461]hh:mm:ss AM/PM;@",
			"[$-10461]h:mm:ss;@",
			"[$-10461]hh:mm:ss;@",
			"[$-10461]h:mm AM/PM;@",
			"[$-10461]hh:mm AM/PM;@",
			"[$-10461]h:mm;@",
			"[$-10461]hh:mm;@"
		],
		"1122": [
			"[$-10462]hh:mm:ss;@",
			"[$-10462]hh:mm;@"
		],
		"1123": [
			"[$-160463]h:mm:ss;@",
			"[$-160463]hh:mm:ss;@",
			"[$-160463]h:mm;@",
			"[$-160463]hh:mm;@"
		],
		"1124": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10464]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10464]hh:mm;@"
		],
		"1125": [
			"hh:mm;@",
			"hh:mm:ss;@",
			"hh:mm;@",
			"[$-465]hh:mm AM/PM;@",
			"hh:mm:ss;@",
			"[$-465]hh:mm:ss AM/PM;@"
		],
		"1126": [
			"[$-10466]hh:mm:ss;@",
			"[$-10466]hh:mm;@"
		],
		"1127": [
			"[$-10467]hh:mm:ss;@",
			"[$-10467]hh:mm;@"
		],
		"1128": [
			"[$-10468]hh:mm:ss;@",
			"[$-10468]hh:mm;@"
		],
		"1129": [
			"[$-10469]hh:mm:ss;@",
			"[$-10469]hh:mm;@"
		],
		"1130": [
			"[$-1046A]h:m:s;@",
			"[$-1046A]hh:mm:ss;@",
			"[$-1046A]h:m;@",
			"[$-1046A]hh:mm;@"
		],
		"1131": [
			"[$-1046B]hh:mm:ss AM/PM;@",
			"[$-1046B]h:mm:ss AM/PM;@",
			"[$-1046B]h:mm:ss;@",
			"[$-1046B]hh:mm:ss;@",
			"[$-1046B]hh:mm AM/PM;@",
			"[$-1046B]h:mm AM/PM;@",
			"[$-1046B]h:mm;@",
			"[$-1046B]hh:mm;@"
		],
		"1132": [
			"[$-1046C]hh:mm:ss;@",
			"[$-1046C]hh:mm;@"
		],
		"1133": [
			"[$-1046D]h:mm:ss;@",
			"[$-1046D]h:mm;@"
		],
		"1134": [
			"[$-1046E]hh:mm:ss;@",
			"[$-1046E]h:mm:ss\" Auer\";@",
			"[$-1046E]hh:mm:ss\" Auer\";@",
			"[$-1046E]hh:mm;@",
			"[$-1046E]h:mm;@",
			"[$-1046E]h.mm;@",
			"[$-1046E]h.mm\" Auer\";@"
		],
		"1135": [
			"[$-1046F]hh:mm:ss;@",
			"[$-1046F]h:mm:ss;@",
			"[$-1046F]hh:mm;@",
			"[$-1046F]h:mm;@"
		],
		"1136": [
			"[$-10470]hh:mm:ss;@",
			"[$-10470]hh:mm;@"
		],
		"1137": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10471]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10471]hh:mm;@"
		],
		"1138": [
			"[$-10472]h:mm:ss AM/PM;@",
			"[$-10472]hh:mm:ss;@",
			"[$-10472]h:mm AM/PM;@",
			"[$-10472]hh:mm;@"
		],
		"1139": [
			"[$-10473]h:mm:ss AM/PM;@",
			"[$-10473]hh:mm:ss;@",
			"[$-10473]h:mm AM/PM;@",
			"[$-10473]hh:mm;@"
		],
		"1140": [
			"[$-10474]hh:mm:ss;@",
			"[$-10474]h:mm:ss;@",
			"[$-10474]hh:mm:ss AM/PM;@",
			"[$-10474]hh:mm;@",
			"[$-10474]h:mm;@",
			"[$-10474]hh:mm AM/PM;@"
		],
		"1141": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10475]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10475]hh:mm;@"
		],
		"1142": [
			"[$-10476]hh:mm:ss;@",
			"[$-10476]hh:mm;@"
		],
		"1143": [
			"[$-10477]h:mm:ss AM/PM;@",
			"[$-10477]hh:mm:ss;@",
			"[$-10477]h:mm AM/PM;@",
			"[$-10477]hh:mm;@"
		],
		"1144": [
			"[$-10478]AM/PM h:mm:ss;@",
			"[$-10478]h:mm:ss;@",
			"[$-10478]hh:mm:ss;@",
			"[$-10478]AM/PM h:mm;@",
			"[$-10478]h:mm;@",
			"[$-10478]hh:mm;@"
		],
		"1145": [
			"[$-10479]h:mm:ss;@",
			"[$-10479]hh:mm:ss;@",
			"[$-10479]h:mm;@",
			"[$-10479]hh:mm;@"
		],
		"1146": [
			"[$-1047A]h:mm:ss;@",
			"[$-1047A]hh:mm:ss;@",
			"[$-1047A]h:mm;@",
			"[$-1047A]hh:mm;@"
		],
		"1148": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-1047C]hh:mm:ss;@",
			"[$-1047C]h:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-1047C]hh:mm;@",
			"[$-1047C]h:mm;@"
		],
		"1150": [
			"[$-1047E]hh:mm:ss;@",
			"[$-1047E]hh:mm;@"
		],
		"1152": [
			"[$-10480]h:mm:ss;@",
			"[$-10480]hh:mm:ss;@",
			"[$-10480]AM/PM h:mm:ss;@",
			"[$-10480]AM/PM hh:mm:ss;@",
			"[$-10480]h:mm;@",
			"[$-10480]hh:mm;@",
			"[$-10480]AM/PM h:mm;@",
			"[$-10480]AM/PM hh:mm;@"
		],
		"1153": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10481]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10481]hh:mm;@"
		],
		"1154": [
			"[$-10482]h\"h\"mm:ss;@",
			"[$-10482]hh\"h\"mm ss \"seg\".;@",
			"[$-10482]h\"h\"mm;@",
			"[$-10482]hh\"h\"mm;@"
		],
		"1155": [
			"[$-10483]h:mm:ss;@",
			"[$-10483]hh:mm:ss;@",
			"[$-10483]hh:mm;@",
			"[$-10483]h:mm;@",
			"[$-10483]hh.mm;@"
		],
		"1156": [
			"[$-10484]hh:mm:ss;@",
			"[$-10484]h:mm:ss;@",
			"[$-10484]hh:mm;@",
			"[$-10484]h:mm;@",
			"[$-10484]hh.mm;@",
			"[$-10484]hh\" h \"mm;@",
			"[$-10484]hh\"h\"mm;@"
		],
		"1157": [
			"[$-10485]hh:mm:ss;@",
			"[$-10485]hh:mm;@"
		],
		"1158": [
			"[$-10486]h:mm:ss AM/PM;@",
			"[$-10486]hh:mm:ss;@",
			"[$-10486]h:mm AM/PM;@",
			"[$-10486]hh:mm;@"
		],
		"1159": [
			"[$-10487]hh:mm:ss;@",
			"[$-10487]hh:mm;@"
		],
		"1160": [
			"[$-10488]hh:mm:ss;@",
			"[$-10488]hh:mm;@"
		],
		"1164": [
			"[$-16048C]h:mm:ss;@",
			"[$-16048C]hh:mm:ss;@",
			"[$-16048C]h:mm;@",
			"[$-16048C]hh:mm;@"
		],
		"1169": [
			"[$-10491]hh:mm:ss;@",
			"[$-10491]hh:mm;@"
		],
		"2052": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"h\"时\"mm\"分\";@",
			"h\"时\"mm\"分\"ss\"秒\";@",
			"上午/下午h\"时\"mm\"分\";@",
			"上午/下午h\"时\"mm\"分\"ss\"秒\";@",
			"[DBNum1][$-804]h\"时\"mm\"分\";@",
			"[DBNum1][$-804]上午/下午h\"时\"mm\"分\";@"
		],
		"2055": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"h.mm\" h\";@",
			"hh.mm\" h\";@",
			"h.mm\" Uhr\";@"
		],
		"2057": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"[$-409]h:mm:ss AM/PM;@"
		],
		"2058": [
			"[$-80A]hh:mm:ss AM/PM;@",
			"[$-80A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"2060": [
			"h:mm:ss;@",
			"hh:mm:ss;@",
			"h.mm;@",
			"h\" h \"mm;@",
			"h\" h \"m\" min \"s\" s \";@"
		],
		"2064": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"h.mm\" h\";@"
		],
		"2067": [
			"h:mm:ss;@",
			"hh:mm:ss;@",
			"h.mm\" u.\";@",
			"h:mm;@"
		],
		"2068": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"\"kl \"hh.mm;@",
			"hh.mm.ss;@"
		],
		"2070": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d/m/yy h:mm AM/PM;@",
			"d/m/yy h:mm;@"
		],
		"2072": [
			"[$-10818]hh:mm:ss;@",
			"[$-10818]hh:mm;@"
		],
		"2073": [
			"[$-10819]hh:mm:ss;@",
			"[$-10819]hh:mm;@"
		],
		"2074": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"2077": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"\"kl \"h:mm;@"
		],
		"2092": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"2094": [
			"[$-1082E]hh:mm:ss;@",
			"[$-1082E]h:mm:ss\" góź.\";@",
			"[$-1082E]\"zeger \"h:mm:ss;@",
			"[$-1082E]hh:mm;@",
			"[$-1082E]h:mm;@",
			"[$-1082E]h:mm\" góź.\";@",
			"[$-1082E]\"zeger \"h:mm;@"
		],
		"2098": [
			"[$-10832]hh:mm:ss;@",
			"[$-10832]hh:mm;@"
		],
		"2107": [
			"[$-1083B]hh:mm:ss;@",
			"[$-1083B]h:mm:ss;@",
			"[$-1083B]hh:mm;@",
			"[$-1083B]h:mm;@"
		],
		"2108": [
			"[$-1083C]hh:mm:ss;@",
			"[$-1083C]hh:mm;@"
		],
		"2110": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"2115": [
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"2117": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10845]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10845]hh:mm;@"
		],
		"2118": [
			"[$-10409]h.mm.ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-10846]h:mm:ss;@",
			"[$-10846]hh:mm:ss;@",
			"[$-10409]h.mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-10846]h:mm;@",
			"[$-10846]hh:mm;@"
		],
		"2128": [
			"[$-10850]h:mm:ss;@",
			"[$-10850]h:mm;@"
		],
		"2137": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-10859]h:mm:ss;@",
			"[$-10859]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-10859]h:mm;@",
			"[$-10859]hh:mm;@"
		],
		"2141": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-1085D]hh:mm:ss;@",
			"[$-1085D]h:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-1085D]hh:mm;@",
			"[$-1085D]h:mm;@"
		],
		"2143": [
			"[$-1085F]h:mm:ss;@",
			"[$-1085F]hh:mm:ss;@",
			"[$-1085F]h:mm;@",
			"[$-1085F]hh:mm;@"
		],
		"2144": [
			"[$-10409]AM/PM h:mm:ss;@",
			"[$-10860]hh:mm:ss;@",
			"[$-10409]AM/PM h:mm;@",
			"[$-10860]hh:mm;@"
		],
		"2145": [
			"[$-10861]h:mm:ss AM/PM;@",
			"[$-10861]hh:mm:ss;@",
			"[$-10861]h:mm AM/PM;@",
			"[$-10861]hh:mm;@"
		],
		"2151": [
			"[$-10867]hh:mm:ss;@",
			"[$-10867]h:mm:ss;@",
			"[$-10867]hh:mm;@",
			"[$-10867]h:mm;@",
			"[$-10867]hh.mm;@",
			"[$-10867]hh\" h \"mm;@"
		],
		"2155": [
			"[$-1086B]h:mm:ss;@",
			"[$-1086B]hh:mm:ss;@",
			"[$-1086B]h:mm;@",
			"[$-1086B]hh:mm;@"
		],
		"2163": [
			"[$-10873]h:mm:ss AM/PM;@",
			"[$-10873]hh:mm:ss;@",
			"[$-10873]h:mm AM/PM;@",
			"[$-10873]hh:mm;@"
		],
		"3073": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"3076": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"3079": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"hh:mm;@",
			"hh:mm\" Uhr\";@"
		],
		"3081": [
			"[$-409]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"3082": [
			"h:mm;@",
			"[$-409]h:mm AM/PM;@",
			"h:mm:ss;@",
			"[$-409]h:mm:ss AM/PM;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-409]d-m-yy h:mm AM/PM;@",
			"d-m-yy h:mm;@"
		],
		"3084": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"h\" h \"mm;@",
			"h:mm;@"
		],
		"3098": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"3131": [
			"[$-10C3B]h:mm:ss;@",
			"[$-10C3B]hh:mm:ss;@",
			"[$-10C3B]h:mm;@",
			"[$-10C3B]hh:mm;@"
		],
		"3152": [
			"[$-10C50]h:mm:ss;@",
			"[$-10C50]h:mm;@"
		],
		"3179": [
			"[$-10C6B]hh:mm:ss AM/PM;@",
			"[$-10C6B]hh:mm:ss;@",
			"[$-10C6B]hh:mm AM/PM;@",
			"[$-10C6B]hh:mm;@"
		],
		"4097": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"4100": [
			"[$-409]AM/PM h:mm:ss;@",
			"[$-409]AM/PM hh:mm:ss;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"4103": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"h.mm;@",
			"h.mm\" Uhr \";@"
		],
		"4105": [
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"4106": [
			"[$-100A]hh:mm:ss AM/PM;@",
			"[$-100A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"4108": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"hh.mm\" h\";@"
		],
		"4122": [
			"[$-1101A]hh:mm:ss;@",
			"[$-1101A]hh:mm;@"
		],
		"4155": [
			"[$-1103B]hh:mm:ss;@",
			"[$-1103B]h:mm:ss;@",
			"[$-1103B]hh.mm.ss;@",
			"[$-1103B]hh:mm;@",
			"[$-1103B]h:mm;@",
			"[$-1103B]hh.mm;@"
		],
		"4191": [
			"h:mm;@",
			"[$-105F]h:mm;@",
			"h:mm:ss;@",
			"[$-105F]h:mm:ss;@",
			"mm:ss.0;@",
			"[h]:mm:ss;@",
			"[$-105F]dd-mm-yy h:mm;@",
			"dd-mm-yy h:mm;@"
		],
		"5121": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"5124": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"5127": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"h.mm\" h\";@",
			"hh.mm\" h\";@",
			"h.mm\" Uhr\";@"
		],
		"5129": [
			"[$-1409]h:mm:ss AM/PM;@",
			"[$-1409]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"5130": [
			"[$-140A]hh:mm:ss AM/PM;@",
			"[$-140A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"5132": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"hh.mm;@",
			"hh\" h \"mm;@"
		],
		"5146": [
			"[$-1141A]hh:mm:ss;@",
			"[$-1141A]hh:mm;@"
		],
		"5179": [
			"[$-1143B]hh:mm:ss;@",
			"[$-1143B]h:mm:ss;@",
			"[$-1143B]hh:mm;@",
			"[$-1143B]h:mm;@"
		],
		"6153": [
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"6154": [
			"[$-180A]hh:mm:ss AM/PM;@",
			"[$-180A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"6156": [
			"hh:mm:ss;@",
			"h:mm:ss;@",
			"hh.mm;@",
			"hh\" h \"mm;@"
		],
		"6170": [
			"[$-1181A]hh:mm:ss;@",
			"[$-1181A]hh:mm;@"
		],
		"6203": [
			"[$-1183B]hh:mm:ss;@",
			"[$-1183B]h:mm:ss;@",
			"[$-1183B]hh:mm;@",
			"[$-1183B]h:mm;@"
		],
		"7169": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"7177": [
			"[$-409]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"7178": [
			"[$-1C0A]hh:mm:ss AM/PM;@",
			"[$-1C0A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"7180": [
			"[$-11C0C]hh:mm:ss;@",
			"[$-11C0C]h:mm:ss;@",
			"[$-11C0C]hh:mm;@",
			"[$-11C0C]h:mm;@"
		],
		"7194": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"7227": [
			"[$-11C3B]hh:mm:ss;@",
			"[$-11C3B]h:mm:ss;@",
			"[$-11C3B]hh:mm;@",
			"[$-11C3B]h:mm;@"
		],
		"8193": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"8201": [
			"[$-409]hh:mm:ss AM/PM;@",
			"[$-409]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"8202": [
			"[$-200A]hh:mm:ss AM/PM;@",
			"[$-200A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"8204": [
			"[$-1200C]hh:mm:ss;@",
			"[$-1200C]hh:mm;@"
		],
		"8218": [
			"[$-1201A]h:mm:ss;@",
			"[$-1201A]hh:mm:ss;@",
			"[$-1201A]h:mm;@",
			"[$-1201A]hh:mm;@"
		],
		"8251": [
			"[$-1203B]h:mm:ss;@",
			"[$-1203B]hh:mm:ss;@",
			"[$-1203B]h:mm;@",
			"[$-1203B]hh:mm;@"
		],
		"9217": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"9225": [
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"9226": [
			"[$-240A]hh:mm:ss AM/PM;@",
			"[$-240A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"9228": [
			"[$-1240C]hh:mm:ss;@",
			"[$-1240C]hh:mm;@"
		],
		"9242": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"9275": [
			"[$-1243B]h:mm:ss;@",
			"[$-1243B]hh:mm:ss;@",
			"[$-1243B]h:mm;@",
			"[$-1243B]hh:mm;@"
		],
		"10241": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"10249": [
			"[$-409]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"10250": [
			"[$-280A]hh:mm:ss AM/PM;@",
			"[$-280A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"10252": [
			"[$-1280C]hh:mm:ss;@",
			"[$-1280C]hh:mm;@"
		],
		"10266": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"11265": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"11273": [
			"[$-409]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"11274": [
			"[$-2C0A]hh:mm:ss AM/PM;@",
			"[$-2C0A]h:mm:ss AM/PM;@",
			"hh:mm:ss;@",
			"h:mm:ss;@"
		],
		"11276": [
			"[$-12C0C]hh:mm:ss;@",
			"[$-12C0C]hh:mm;@"
		],
		"12289": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"12297": [
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"12298": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"12300": [
			"[$-1300C]hh:mm:ss;@",
			"[$-1300C]hh:mm;@"
		],
		"13313": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"13321": [
			"[$-409]h:mm:ss AM/PM;@",
			"[$-409]hh:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"13322": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"13324": [
			"[$-1340C]hh:mm:ss;@",
			"[$-1340C]hh:mm;@"
		],
		"14337": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"14345": [
			"[$-13809]hh:mm:ss;@",
			"[$-13809]h:mm:ss;@",
			"[$-13809]hh:mm;@",
			"[$-13809]h:mm;@"
		],
		"14346": [
			"[$-380A]hh:mm:ss AM/PM;@",
			"[$-380A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"14348": [
			"[$-1380C]hh:mm:ss;@",
			"[$-1380C]hh:mm;@"
		],
		"15361": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"15369": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-13C09]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-13C09]hh:mm;@"
		],
		"15370": [
			"[$-3C0A]hh:mm:ss AM/PM;@",
			"[$-3C0A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"15372": [
			"[$-13C0C]hh:mm:ss;@",
			"[$-13C0C]hh:mm;@"
		],
		"16385": [
			"[$-1000000]h:mm:ss;@",
			"[$-1000401]h:mm AM/PM;@",
			"[$-1000409]h:mm AM/PM;@",
			"[$-2000000]h:mm:ss;@",
			"[$-2000401]h:mm AM/PM;@",
			"[$-2000409]h:mm AM/PM;@"
		],
		"16393": [
			"[$-14009]hh:mm:ss;@",
			"[$-14009]h:mm:ss;@",
			"[$-10409]h.mm.ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-14009]hh:mm;@",
			"[$-14009]h:mm;@",
			"[$-10409]hh:mm AM/PM;@"
		],
		"16394": [
			"[$-400A]hh:mm:ss AM/PM;@",
			"[$-400A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"17417": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-14409]h:mm:ss;@",
			"[$-14409]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-14409]h:mm;@",
			"[$-14409]hh:mm;@"
		],
		"17418": [
			"[$-440A]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"18441": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-14809]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-14809]hh:mm;@"
		],
		"18442": [
			"[$-480A]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"19466": [
			"[$-4C0A]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"20490": [
			"[$-500A]hh:mm:ss AM/PM;@",
			"hh:mm:ss;@"
		],
		"21514": [
			"[$-10409]h:mm:ss AM/PM;@",
			"[$-10409]hh:mm:ss AM/PM;@",
			"[$-1540A]h:mm:ss;@",
			"[$-1540A]hh:mm:ss;@",
			"[$-10409]h:mm AM/PM;@",
			"[$-10409]hh:mm AM/PM;@",
			"[$-1540A]h:mm;@",
			"[$-1540A]hh:mm;@"
		],
		"22538": [
			"[$-580A]hh:mm:ss AM/PM;@",
			"[$-580A]h:mm:ss AM/PM;@",
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"64546": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"64547": [
			"h:mm:ss;@",
			"hh:mm:ss;@"
		],
		"64551": [
			"hh:mm:ss;@",
			"hh:mm;@"
		]
	};

	// Japanese era boundaries (Excel parity, see Microsoft KB c52091af).
	// Lookup by (year, month, day) keeps it independent of AscCommon.bDate1904.
	var g_aJapanEras = [
		{ latinShort: 'M', kanjiShort: '\u660e', kanjiFull: '\u660e\u6cbb', startYear: 1868, startMonth: 9,  startDay: 8  }, // Meiji
		{ latinShort: 'T', kanjiShort: '\u5927', kanjiFull: '\u5927\u6b63', startYear: 1912, startMonth: 7,  startDay: 30 }, // Taisho
		{ latinShort: 'S', kanjiShort: '\u662d', kanjiFull: '\u662d\u548c', startYear: 1926, startMonth: 12, startDay: 25 }, // Showa
		{ latinShort: 'H', kanjiShort: '\u5e73', kanjiFull: '\u5e73\u6210', startYear: 1989, startMonth: 1,  startDay: 8  }, // Heisei
		{ latinShort: 'R', kanjiShort: '\u4ee4', kanjiFull: '\u4ee4\u548c', startYear: 2019, startMonth: 5,  startDay: 1  }  // Reiwa
	];

	// LCIDs that activate Japanese era rendering for g/e tokens. Map shape kept
	// for future variants; Phase 1 has ja-JP only.
	var g_oJapanEraLcids = { 0x0411: true };

	function isJapanEraLcid(lcid) {
		return null != lcid && !!g_oJapanEraLcids[lcid & 0xFFFF];
	}

	function isJapanEraLid(lid) {
		var nLid = parseInt(lid, 16);
		return !isNaN(nLid) && isJapanEraLcid(nLid);
	}

	/**
	 * @param {number} year   four-digit Gregorian year
	 * @param {number} month  1-based month (caller must convert 0-based parseDate() result)
	 * @param {number} day    1-based day-of-month
	 * @returns {Object|null} era descriptor, or null for pre-Meiji input
	 */
	function getJapanEraByDate(year, month, day) {
		for (var i = g_aJapanEras.length - 1; i >= 0; --i) {
			var era = g_aJapanEras[i];
			if (year > era.startYear) {
				return era;
			}
			if (year === era.startYear) {
				if (month > era.startMonth) {
					return era;
				}
				if (month === era.startMonth && day >= era.startDay) {
					return era;
				}
			}
		}
		return null;
	}

	// Locale format-symbol overrides, keyed by BCP-47 / locale name.
	// Each entry is either an overrides object (keys that differ from the
	// invariant defaults) or a string alias to another canonical entry.
	// Consumed by ParseLocalFormatSymbol in NumFormat.js.
	//
	// Note on G:'X', g:'x' entries: in those locales the letter 'g'/'G' is
	// already occupied by a date/hour token, so the localized UI parser maps
	// the invariant Japanese-era 'g' token to 'x'/'X'. The invariant parser
	// path (used by styles.xml / API callers) always sees 'g' directly and
	// does not use these symbols.
	var g_oLocaleFormatSymbols = {
		// Finnish
		'fi': {
			Y:'V', y:'v', M:'K', m:'k', D:'P', d:'p', H:'T', h:'t',
			general:'Yleinen'
		},
		'smn':    'fi', 'sms':    'fi', 'fi-FI':  'fi', 'se-FI':  'fi',
		'smn-FI': 'fi', 'sms-FI': 'fi', 'sv-AX':  'fi', 'sv-FI':  'fi', 'en-FI': 'fi',

		// Dutch / Frisian
		'fy': { Y:'J', y:'j', H:'U', h:'u', general:'Standaard' },
		'nds':    'fy', 'nl':     'fy', 'en-NL':  'fy', 'fy-NL':  'fy',
		'nds-NL': 'fy', 'nl-BE':  'fy', 'nl-NL':  'fy',

		// Spanish / Catalan / Galician
		'ast': { Y:'A', y:'a', a:'o', general:'Estándar' },
		'eu':  'ast', 'gl': 'ast', 'ast-ES': 'ast', 'ca-ES':  'ast',
		'es-ES': 'ast', 'es-MX': 'ast', 'eu-ES': 'ast', 'gl-ES': 'ast', 'ca-ES-valencia': 'ast',

		// Portuguese (Brazil)
		'pt-BR': { Y:'A', y:'a', a:'o', general:'Geral' },
		'es-BR': 'pt-BR',

		// Portuguese (Portugal)
		'pt': { Y:'A', y:'a', a:'o', general:'Éstandar' },
		'pt-PT': 'pt',

		// Russian / Cyrillic
		'ba': {
			Y:'Г', y:'г', M:'М', m:'М', D:'Д', d:'д',
			H:'Ч', h:'ч', Minute:'М', minute:'м', S:'C', s:'с',
			general:'Основной'
		},
		'ce':     'ba', 'cu':    'ba', 'kk':     'ba', 'os':    'ba', 'rm':    'ba', 'ru':    'ba',
		'sah':    'ba', 'tt':    'ba', 'wae':    'ba', 'ba-RU': 'ba', 'ce-RU': 'ba',
		'cu-RU':  'ba', 'de-BE': 'ba', 'en-BE':  'ba', 'en-CH': 'ba', 'kk-KZ': 'ba',
		'os-RU':  'ba', 'pt-CH': 'ba', 'rm-CH':  'ba', 'ru-KZ': 'ba', 'ru-RU': 'ba',
		'sah-RU': 'ba', 'tt-RU': 'ba', 'wae-CH': 'ba',

		// French
		'oc': { Y:'A', y:'a', D:'J', d:'j', a:'o', general:'Standard' },
		'br':     'oc', 'co':    'oc', 'fr':     'oc', 'br-FR': 'oc', 'ca-FR':  'oc',
		'co-FR':  'oc', 'fr-BE': 'oc', 'fr-CA':  'oc', 'fr-CH': 'oc', 'fr-FR':  'oc', 'gsw-FR': 'oc',

		// German
		'de': {
			Y:'J', y:'j', M:'M', m:'M', D:'T', d:'t',
			Minute:'M', minute:'m', general:'Standard'
		},
		'ksh':    'de', 'dsb':   'de', 'hsb':    'de', 'de-AT': 'de', 'de-CH':  'de', 'de-DE': 'de',
		'dsb-DE': 'de', 'en-AT': 'de', 'en-DE':  'de', 'hsb-DE':'de', 'ksh-DE': 'de', 'nds-DE':'de',

		// Italian / Catalan (G->X: 'g'/'G' occupied by day token in this locale)
		'ca': {
			Y:'A', y:'a', D:'G', d:'g',
			G:'X', g:'x',
			a:'o', general:'Standard'
		},
		'it':    'ca', 'fur':   'ca', 'ca-IT':  'ca', 'de-IT': 'ca', 'fur-IT': 'ca',
		'it-CH': 'ca', 'it-IT': 'ca', 'it-VA':  'ca',

		// Swedish
		'sv': {
			Y:'Å', y:'å', M:'M', m:'M', H:'T', h:'t',
			Minute:'M', minute:'m', general:'Standard'
		},
		'en-SE':  'sv', 'se-SE':  'sv', 'sma-SE': 'sv', 'smj-SE': 'sv', 'sv-SE': 'sv',

		// Danish / Norwegian / Faroese / Sami
		'nb': { Y:'Å', y:'å', H:'T', h:'t', general:'Standard' },
		'nn':     'nb', 'se':    'nb', 'smj':    'nb', 'sma':   'nb', 'fo':    'nb', 'da':   'nb',
		'smj-NO': 'nb', 'sma-NO':'nb', 'se-NO':  'nb', 'nn-NO': 'nb', 'nb-SJ': 'nb',
		'nb-NO':  'nb', 'fo-DK': 'nb', 'da-DK':  'nb',

		// Chinese / Tibetan / Yi / Uyghur / Mongolian
		'bo': { general:'G/通用格式' },
		'ii':         'bo', 'ug':    'bo', 'zh':     'bo', 'bo-CN':   'bo', 'ii-CN':  'bo',
		'mn-Mong-CN': 'bo', 'ug-CN': 'bo', 'zh-CN':  'bo', 'zh-Hans': 'bo', 'zh-TW':  'bo',

		// Greek
		'el': {
			Y:'Ε', y:'ε', M:'Μ', m:'μ', D:'Η', d:'η',
			H:'Ω', h:'ω', Minute:'Λ', minute:'λ', S:'Δ', s:'δ',
			general:'Γενικός τύπος'
		},
		'el-GR': 'el',

		// Hungarian
		'hu': {
			Y:'É', y:'é', M:'H', m:'h', D:'N', d:'n',
			H:'Ó', h:'ó', Minute:'P', minute:'p', S:'M', s:'m',
			general:'Normál'
		},
		'hu-HU': 'hu',

		// Turkish (G->X: 'g'/'G' occupied by day token in this locale)
		'tr': {
			M:'A', m:'a', D:'G', d:'g',
			G:'X', g:'x',
			H:'S', h:'s', Minute:'D', minute:'d', S:'N', s:'n',
			a:'o', general:'Genel'
		},
		'tr-TR': 'tr',

		// Polish (G->X: 'g'/'G' occupied by hour token)
		'pl': {
			Y:'R', y:'r', H:'G', h:'g',
			G:'X', g:'x',
			general:'Standardowy'
		},
		'pl-PL': 'pl',

		// Czech
		'cs': { Y:'R', y:'r', general:'Vęeobecný' },
		'cs-CZ': 'cs',

		// Japanese
		'ja': { general:'G/標準' },
		'ja-JP': 'ja',

		// Korean
		'ko': { general:'G/표준' },
		'ko-KR': 'ko'
	};

	//---------------------------------------------------------export---------------------------------------------------
	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].g_aCultureInfos = g_aCultureInfos;
	window["AscCommon"].g_aAdditionalCurrencySymbols = g_aAdditionalCurrencySymbols;
	window["AscCommon"].c_oAscDateFormatExcel = c_oAscDateFormatExcel;
	window["AscCommon"].c_oAscTimeFormatExcel = c_oAscTimeFormatExcel;
	window["AscCommon"].g_aJapanEras = g_aJapanEras;
	window["AscCommon"].g_oJapanEraLcids = g_oJapanEraLcids;
	window["AscCommon"].isJapanEraLcid = isJapanEraLcid;
	window["AscCommon"].isJapanEraLid = isJapanEraLid;
	window["AscCommon"].getJapanEraByDate = getJapanEraByDate;
	window["AscCommon"].g_oLocaleFormatSymbols = g_oLocaleFormatSymbols;

})(window);
