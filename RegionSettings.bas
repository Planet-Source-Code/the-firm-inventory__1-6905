Attribute VB_Name = "RegionSettings"


Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
        (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, _
        ByVal cchData As Long) As Long
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Public Const LOCALE_ILANGUAGE = &H1         '  language id
Public Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Public Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Public Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Public Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Public Const LOCALE_ICOUNTRY = &H5         '  country code
Public Const LOCALE_SCOUNTRY = &H6         '  localized name of country
Public Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country
Public Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
Public Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
Public Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Public Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country code
Public Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page

Public Const LOCALE_SLIST = &HC         '  list item separator
Public Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US

Public Const LOCALE_SDECIMAL = &HE         '  decimal separator
Public Const LOCALE_STHOUSAND = &HF         '  thousand separator
Public Const LOCALE_SGROUPING = &H10        '  digit grouping
Public Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Public Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Public Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9

Public Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Public Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Public Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Public Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Public Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Public Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Public Const LOCALE_INEGCURR = &H1C        '  negative currency mode

Public Const LOCALE_SDATE = &H1D        '  date separator
Public Const LOCALE_STIME = &H1E        '  time separator
Public Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Public Const LOCALE_SLONGDATE = &H20        '  long date format string
Public Const LOCALE_STIMEFORMAT = &H1003      '  time format string
Public Const LOCALE_IDATE = &H21        '  short date format ordering
Public Const LOCALE_ILDATE = &H22        '  long date format ordering
Public Const LOCALE_ITIME = &H23        '  time format specifier
Public Const LOCALE_ICENTURY = &H24        '  century format specifier
Public Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Public Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Public Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Public Const LOCALE_S1159 = &H28        '  AM designator
Public Const LOCALE_S2359 = &H29        '  PM designator

Public Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Public Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Public Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Public Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Public Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Public Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Public Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Public Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Public Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Public Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Public Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Public Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Public Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Public Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Public Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Public Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Public Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Public Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Public Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Public Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Public Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Public Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Public Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Public Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Public Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Public Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Public Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Public Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Public Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Public Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Public Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Public Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F

Public Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Public Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Public Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Public Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Public Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Public Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Public Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Public Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt

Public Sub GetLanguage()

    Dim Symbol As String
    Dim ll_Return As Long
    Dim iRet2 As Long
    Dim ls_Buffer As String
    Dim pos As Integer
    Dim ll_LocaleID As Long

    ll_LocaleID = GetUserDefaultLCID()
    ll_Return = GetLocaleInfo(ll_LocaleID, LOCALE_SLANGUAGE, ls_Buffer, 0)
    ls_Buffer = Space(ll_Return)
    iRet2 = GetLocaleInfo(ll_LocaleID, LOCALE_SLANGUAGE, ls_Buffer, ll_Return)
    Invent.Label1(9).Caption = "Language": Invent.Label2(9).Caption = ls_Buffer

End Sub
