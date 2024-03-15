using SpreadsheetMLDotNet.Attributes;

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
    [IsDateTime]
    GeneralDateShort = 14,

    /// <summary>
    /// d-mmm-yy
    /// </summary>
    [IsDateTime]
    GeneralDateLong = 15,

    /// <summary>
    /// d-mmm
    /// </summary>
    [IsDateTime]
    GeneralDateMonthDay = 16,

    /// <summary>
    /// mmm-yy
    /// </summary>
    [IsDateTime]
    GeneralDateYearMonth = 17,

    /// <summary>
    /// h:mm AM/PM
    /// </summary>
    [IsDateTime]
    GeneralTimeHourMinute_AMPM = 18,

    /// <summary>
    /// h:mm:ss AM/PM
    /// </summary>
    [IsDateTime]
    GeneralTimeHourMinuteSecond_AMPM = 19,

    /// <summary>
    /// h:mm
    /// </summary>
    [IsDateTime]
    GeneralTimeHourMinute = 20,

    /// <summary>
    /// h:mm:ss
    /// </summary>
    [IsDateTime]
    GeneralTimeHourMinuteSecond = 21,

    /// <summary>
    /// m/d/yy h:mm
    /// </summary>
    [IsDateTime]
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
    [IsDateTime]
    GeneralTimeMinuteSecond = 45,

    /// <summary>
    /// [h]:mm:ss
    /// </summary>
    GeneralTimeOver24HourMinuteSecond = 46,

    /// <summary>
    /// mmss.0
    /// </summary>
    [IsDateTime]
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
    [IsDateTime]
    LocalEraDateShort = 27,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    [IsDateTime]
    LocalEraDateLong = 28,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    [IsDateTime]
    LocalEraDateLong2 = 29,

    /// <summary>
    /// zh-tw: m/d/yy
    /// zh-cn: m-d-yy
    /// ja-jp: m/d/yy
    /// ko-kr: mm-dd-yy
    /// </summary>
    [IsDateTime]
    LocalDateShort = 30,

    /// <summary>
    /// zh-tw: yyyy"年"m"月"d"日"
    /// zh-cn: yyyy"年"m"月"d"日"
    /// ja-jp: yyyy"年"m"月"d"日"
    /// ko-kr: yyyy"년" mm"월" dd"일"
    /// </summary>
    [IsDateTime]
    LocalDateLong = 31,

    /// <summary>
    /// zh-tw: hh"時"mm"分"
    /// zh-cn: h"时"mm"分"
    /// ja-jp: h"時"mm"分"
    /// ko-kr: h"시" mm"분"
    /// </summary>
    [IsDateTime]
    LocalTimeHourMinute = 32,

    /// <summary>
    /// zh-tw: hh"時"mm"分"ss"秒"
    /// zh-cn: h"时"mm"分"ss"秒"
    /// ja-jp: h"時"mm"分"ss"秒"
    /// ko-kr: h"시" mm"분" ss"초"
    /// </summary>
    [IsDateTime]
    LocalTimeHourMinuteSecond = 33,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"
    /// zh-cn: 上午/下午 h"时"mm"分"
    /// ja-jp: yyyy"年"m"月"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    [IsDateTime]
    Local34 = 34,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"ss"秒"
    /// zh-cn: 上午/下午 h"时"mm"分"ss"秒"
    /// ja-jp: m"月"d"日"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    [IsDateTime]
    Local35 = 35,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    [IsDateTime]
    Local36 = 36,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    [IsDateTime]
    Local50 = 50,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    [IsDateTime]
    Local51 = 51,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: yyyy"年"m"月"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    [IsDateTime]
    Local52 = 52,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"ss"秒"
    /// zh-cn: m"月"d"日"
    /// ja-jp: m"月"d"日"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    [IsDateTime]
    Local53 = 53,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    [IsDateTime]
    Local54 = 54,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"
    /// zh-cn: 上午/下午 h"时"mm"分"
    /// ja-jp: yyyy"年"m"月"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    [IsDateTime]
    Local55 = 55,

    /// <summary>
    /// zh-tw: 上午/下午 hh"時"mm"分"ss"秒"
    /// zh-cn: 上午/下午 h"时"mm"分"ss"秒"
    /// ja-jp: m"月"d"日"
    /// ko-kr: yyyy-mm-dd
    /// </summary>
    [IsDateTime]
    Local56 = 56,

    /// <summary>
    /// zh-tw: [$-404]e/m/d
    /// zh-cn: yyyy"年"m"月"
    /// ja-jp: [$-411]ge.m.d
    /// ko-kr: yyyy"年" mm"月" dd"日"
    /// </summary>
    [IsDateTime]
    Local57 = 57,

    /// <summary>
    /// zh-tw: [$-404]e"年"m"月"d"日"
    /// zh-cn: m"月"d"日"
    /// ja-jp: [$-411]ggge"年"m"月"d"日"
    /// ko-kr: mm-dd
    /// </summary>
    [IsDateTime]
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
    [IsDateTime]
    Local71 = 71,

    /// <summary>
    /// th-th: ว-ดดด-ปป
    /// </summary>
    [IsDateTime]
    Local72 = 72,

    /// <summary>
    /// th-th: ว-ดดด
    /// </summary>
    [IsDateTime]
    Local73 = 73,

    /// <summary>
    /// th-th: ดดด-ปป
    /// </summary>
    [IsDateTime]
    Local74 = 74,

    /// <summary>
    /// th-th: ช:นน
    /// </summary>
    [IsDateTime]
    Local75 = 75,

    /// <summary>
    /// th-th: ช:นน:ทท
    /// </summary>
    [IsDateTime]
    Local76 = 76,

    /// <summary>
    /// th-th: ว/ด/ปปปป ช:นน
    /// </summary>
    [IsDateTime]
    Local77 = 77,

    /// <summary>
    /// th-th: นน:ทท
    /// </summary>
    [IsDateTime]
    Local78 = 78,

    /// <summary>
    /// th-th: [ช]:นน:ทท
    /// </summary>
    [IsDateTime]
    Local79 = 79,

    /// <summary>
    /// th-th: นน:ทท.0
    /// </summary>
    [IsDateTime]
    Local80 = 80,

    /// <summary>
    /// th-th: d/m/bb
    /// </summary>
    [IsDateTime]
    Local81 = 81,
}
