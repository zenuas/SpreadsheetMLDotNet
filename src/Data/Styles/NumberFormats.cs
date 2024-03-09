namespace SpreadsheetMLDotNet.Data.Styles;

public enum NumberFormats
{
    General = 0,

    /// <summary>
    /// 0
    /// </summary>
    GeneralInt = 1,

    /// <summary>
    /// 0.00
    /// </summary>
    GeneralFloat = 2,

    /// <summary>
    /// #,##0
    /// </summary>
    GeneralIntSeparate = 3,

    /// <summary>
    /// #,##0.00
    /// </summary>
    GeneralFloatSeparate = 4,

    /// <summary>
    /// 0%
    /// </summary>
    GeneralIntPercent = 9,

    /// <summary>
    /// 0.00%
    /// </summary>
    GeneralFloatPercent = 10,

    /// <summary>
    /// 0.00E+00
    /// </summary>
    GeneralExponentialNotation = 11,

    /// <summary>
    /// # ?/?
    /// </summary>
    GeneralFractionShort = 12,

    /// <summary>
    /// # ??/??
    /// </summary>
    GeneralFractionLong = 13,

    /// <summary>
    /// mm-dd-yy
    /// </summary>
    GeneralDateShort = 14,

    /// <summary>
    /// d-mmm-yy
    /// </summary>
    GeneralDateLong = 15,

    /// <summary>
    /// d-mmm
    /// </summary>
    GeneralDateMonthDay = 16,

    /// <summary>
    /// mmm-yy
    /// </summary>
    GeneralDateYearMonth = 17,

    /// <summary>
    /// h:mm AM/PM
    /// </summary>
    GeneralTimeHourMinute_AMPM = 18,

    /// <summary>
    /// h:mm:ss AM/PM
    /// </summary>
    GeneralTimeHourMinuteSecond_AMPM = 19,

    /// <summary>
    /// h:mm
    /// </summary>
    GeneralTimeHourMinute = 20,

    /// <summary>
    /// h:mm:ss
    /// </summary>
    GeneralTimeHourMinuteSecond = 21,

    /// <summary>
    /// m/d/yy h:mm
    /// </summary>
    GeneralDateTime = 22,

    /// <summary>
    /// #,##0 ;(#,##0)
    /// </summary>
    GeneralIntNegativeParentheses = 37,

    /// <summary>
    /// #,##0 ;[Red](#,##0)
    /// </summary>
    GeneralIntNegativeColor = 38,

    /// <summary>
    /// #,##0.00 ;(#,##0.00)
    /// </summary>
    GeneralFloatNegativeParentheses = 39,

    /// <summary>
    /// #,##0.00 ;[Red](#,##0.00)
    /// </summary>
    GeneralFloatNegativeColor = 40,

    /// <summary>
    /// mm:ss
    /// </summary>
    GeneralTimeMinuteSecond = 45,

    /// <summary>
    /// [h]:mm:ss
    /// </summary>
    GeneralTimeOver24HourMinuteSecond = 46,

    /// <summary>
    /// mmss.0
    /// </summary>
    GeneralTimeMinuteSecondDecimal = 47,

    /// <summary>
    /// ##0.0E+0
    /// </summary>
    GeneralExponentialNotationSeparate = 48,

    /// <summary>
    /// @
    /// </summary>
    GeneralString = 49,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    LocalEraDateShort = 27,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    LocalEraDateLong = 28,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    LocalEraDateLong2 = 29,

    /// <summary>
    /// zh-tw: m/d/yy
    /// zh-cn: m-d-yy
    /// ja-jp: m/d/yy
    /// ko-kr: mm-dd-yy
    /// </summary>
    LocalDateShort = 30,

    /// <summary>
    /// zh-tw: yyyy"年"m"月"d"日"
    /// zh-cn: yyyy"年"m"月"d"日"
    /// ja-jp: yyyy"年"m"月"d"日"
    /// ko-kr: yyyy"년" mm"월" dd"일"
    /// </summary>
    LocalDateLong = 31,

    /// <summary>
    /// zh-tw: hh"時"mm"分"
    /// zh-cn: h"时"mm"分"
    /// ja-jp: h"時"mm"分"
    /// ko-kr: h"시" mm"분"
    /// </summary>
    LocalTimeHourMinute = 32,

    /// <summary>
    /// zh-tw: hh"時"mm"分"ss"秒"
    /// zh-cn: h"时"mm"分"ss"秒"
    /// ja-jp: h"時"mm"分"ss"秒"
    /// ko-kr: h"시" mm"분" ss"초"
    /// </summary>
    LocalTimeHourMinuteSecond = 33,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"
    /// zh-cn: 上午/下午 h"时"mm"分"
    /// ja-jp: yyyy"年"m"月"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    Local34 = 34,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"ss"秒"
    /// zh-cn: 上午/下午 h"时"mm"分"ss"秒"
    /// ja-jp: m"月"d"日"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    Local35 = 35,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    Local36 = 36,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    Local50 = 50,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    Local51 = 51,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: yyyy"年"m"月"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    Local52 = 52,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"ss"秒"
    /// zh-cn: m"月"d"日"
    /// ja-jp: m"月"d"日"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    Local53 = 53,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    Local54 = 54,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"
    /// zh-cn: 上午/下午 h"时"mm"分"
    /// ja-jp: yyyy"年"m"月"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    Local55 = 55,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"ss"秒"
    /// zh-cn: 上午/下午 h"时"mm"分"ss"秒"
    /// ja-jp: m"月"d"日"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    Local56 = 56,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    Local57 = 57,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    Local58 = 58,

    /// <summary>
    /// th-th: t0
    /// </summary>
    Local59 = 59,

    /// <summary>
    /// th-th: t0.00
    /// </summary>
    Local60 = 60,

    /// <summary>
    /// th-th: t#,##0
    /// </summary>
    Local61 = 61,

    /// <summary>
    /// th-th: t#,##0.00
    /// </summary>
    Local62 = 62,

    /// <summary>
    /// th-th: t0%
    /// </summary>
    Local67 = 67,

    /// <summary>
    /// th-th: t0.00%
    /// </summary>
    Local68 = 68,

    /// <summary>
    /// th-th: t# ?/?
    /// </summary>
    Local69 = 69,

    /// <summary>
    /// th-th: t# ??/??
    /// </summary>
    Local70 = 70,

    /// <summary>
    /// th-th: ว/ด/ปปปป
    /// </summary>
    Local71 = 71,

    /// <summary>
    /// th-th: ว-ดดด-ปป
    /// </summary>
    Local72 = 72,

    /// <summary>
    /// th-th: ว-ดดด
    /// </summary>
    Local73 = 73,

    /// <summary>
    /// th-th: ดดด-ปป
    /// </summary>
    Local74 = 74,

    /// <summary>
    /// th-th: ช:นน
    /// </summary>
    Local75 = 75,

    /// <summary>
    /// th-th: ช:นน:ทท
    /// </summary>
    Local76 = 76,

    /// <summary>
    /// th-th: ว/ด/ปปปป ช:นน
    /// </summary>
    Local77 = 77,

    /// <summary>
    /// th-th: นน:ทท
    /// </summary>
    Local78 = 78,

    /// <summary>
    /// th-th: [ช]:นน:ทท
    /// </summary>
    Local79 = 79,

    /// <summary>
    /// th-th: นน:ทท.0
    /// </summary>
    Local80 = 80,

    /// <summary>
    /// th-th: d/m/bb
    /// </summary>
    Local81 = 81,
}
