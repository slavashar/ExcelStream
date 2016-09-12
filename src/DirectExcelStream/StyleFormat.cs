using System.ComponentModel;

namespace ExcelStream
{
    public enum StyleFormat : int
    {
        //[Description("0")]
        Number = 1,

        //[Description("0.00")]
        NumberWithDecimal = 2,
                
        //[Description("#,##0")]
        NumberWithComma = 3,
                
        //[Description("#,##0.00")]
        NumberWithCommaAndDecimal = 4,

        //[Description("#,##0;[Red]-#,##0")]
        NumberWithCommaAndNegativeRed = 38,

        //[Description("#,##0.00;[Red]-#,##0.00")]
        NumberWithCommaAndDecimalAndNegativeRed = 40,

        //[Description("$#,##0;[Red]-$#,##0")]
        CurrencyWithCommaAndNegativeRed = 6,

        //[Description("$#,##0.00;[Red]-$#,##0.00")]
        CurrencyWithCommaAndDecimalAndNegativeRed = 8,

        //[Description("0%")]
        Percentage = 9,
                
        //[Description("0.00%")]
        PercentageAndDecimal = 10,
                
        //[Description("0.00E+00")]
        Scientific = 11,
                
        //[Description("# ?/?")]
        Fraction = 12,

        //[Description("# ??/??")]
        FractionWithTwoDigits = 13,

        //[Description("m/d/yyyy")]
        Date = 14,
        
        //[Description("d-mmm-yy")]
        DayMonthYear = 15,

        //[Description("d-mmm")]
        DayMonth = 16,

        //[Description("mmm-yy")]
        MonthYear = 17,
        
        //[Description("h:mm AM/PM")]
        HourMinutePeriod = 18,
        
        //[Description("h:mm:ss AM/PM")]
        HourMinuteSecondPeriod = 19,
        
        //[Description("h:mm")]
        Time = 20,
        
        //[Description("h:mm:ss")]
        TimeWithSeconds = 21,
        
        //[Description("m/d/yyyy h:mm")]
        DateTime = 22,

        //[Description("mm:ss")]
        TimeSpan = 45,

        //[Description("[h]:mm:ss")]
        TimeSpanWithHours = 46,

        //[Description("mm:ss.0")]
        TimeSpanWithDecimal = 46,

        //[Description("@")]
        Text = 49
    }
}
