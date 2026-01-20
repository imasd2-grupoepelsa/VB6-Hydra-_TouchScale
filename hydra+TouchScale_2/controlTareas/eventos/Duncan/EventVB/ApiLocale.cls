VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiLocale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2773
Attribute VB_Description="This class provides information about a LOCALE that is installed on the system."
Option Explicit

' ##MODULE_DESCRIPTION This class provides information about a LOCALE that is installed on the system.

' ##MODULE_DESCRIPTION A locale is a group of settings that identify the way a system _
displays information (currencies, day names, number formats etc.) in a given country.

' ##MODULE_DESCRIPTION For example, in Ireland the currency unit is the Euro, _
we write numbers with a full stop for a decimal separator and the days of the _
week are Monday, Tuesday, Wednesday etc.

' ##MODULE_DESCRIPTION You should use locale information to make your _
applications behave in a manner that it consistent with your users _
national experience and therefore easier for them to use.

' ##MODULE_DESCRIPTION How it works:
' ##MODULE_DESCRIPTION The locale ID is a 32 bit long number (0-31)
' ##MODULE_DESCRIPTION  Bits 20-31 are reserved for future use
' ##MODULE_DESCRIPTION  Bits 16-19 are the sort order (SORT_*)
' ##MODULE_DESCRIPTION  Bits 10-15 are the sub language (SUBLANG_*)
' ##MODULE_DESCRIPTION  Bits 0 - 9 are the main language ID (LANG_*)

Private mLocaleId As Long


Public Enum enLocaleSortOrders
    SORT_DEFAULT = &H0                          '  sorting default
    SORT_CHINESE_BIG5 = &H0                     '  Chinese BIG5 order
    SORT_CHINESE_UNICODE = &H1                  '  Chinese Unicode order
    SORT_CHINESE_PRC = &H2                      '  PRC Chinese Stroke Count order
    SORT_JAPANESE_UNICODE = &H1                 '  Japanese Unicode order
    SORT_JAPANESE_XJIS = &H0                    '  Japanese= &hJIS order
    SORT_KOREAN_KSC = &H0                       '  Korean KSC order
    SORT_KOREAN_UNICODE = &H1                   '  Korean Unicode order
    SORT_GERMAN_PHONEBOOK = &H1
End Enum

Public Enum enLocaleMainLanguage
    LANG_NEUTRAL = &H0
    LANG_AFRIKAANS = &H36
    LANG_ALBANIAN = &H1C
    LANG_ARABIC = &H1
    LANG_BASQUE = &H2D
    LANG_BELARUSIAN = &H23
    LANG_BULGARIAN = &H2
    LANG_CATALAN = &H3
    LANG_CHINESE = &H4
    LANG_CROATIAN = &H1A
    LANG_CZECH = &H5
    LANG_DANISH = &H6
    LANG_DUTCH = &H13
    LANG_ENGLISH = &H9
    LANG_ESTONIAN = &H25
    LANG_FAEROESE = &H38
    LANG_FARSI = &H29
    LANG_FINNISH = &HB
    LANG_FRENCH = &HC
    LANG_GERMAN = &H7
    LANG_GREEK = &H8
    LANG_HEBREW = &HD
    LANG_HINDI = &H39
    LANG_HUNGARIAN = &HE
    LANG_ICELANDIC = &HF
    LANG_INDONESIAN = &H21
    LANG_ITALIAN = &H10
    LANG_JAPANESE = &H11
    LANG_KOREAN = &H12
    LANG_LATVIAN = &H26
    LANG_LITHUANIAN = &H27
    LANG_MACEDONIAN = &H2F
    LANG_MALAY = &H3E
    LANG_NORWEGIAN = &H14
    LANG_POLISH = &H15
    LANG_PORTUGUESE = &H16
    LANG_ROMANIAN = &H18
    LANG_RUSSIAN = &H19
    LANG_SERBIAN = &H1A
    LANG_SLOVAK = &H1B
    LANG_SLOVENIAN = &H24
    LANG_SPANISH = &HA
    LANG_SWAHILI = &H41
    LANG_SWEDISH = &H1D
    LANG_THAI = &H1E
    LANG_TURKISH = &H1F
    LANG_UKRAINIAN = &H22
    LANG_VIETNAMESE = &H2A
End Enum

Public Enum enLocaleSublanguages
    SUBLANG_NEUTRAL = &H0                      ' language neutral
    SUBLANG_DEFAULT = &H1                      ' user default
    SUBLANG_SYS_DEFAULT = &H2                  ' system default
    SUBLANG_ARABIC_SAUDI_ARABIA = &H1          ' Arabic (Saudi Arabia)
    SUBLANG_ARABIC_IRAQ = &H2                  ' Arabic (Iraq)
    SUBLANG_ARABIC_EGYPT = &H3                 ' Arabic (Egypt)
    SUBLANG_ARABIC_LIBYA = &H4                 ' Arabic (Libya)
    SUBLANG_ARABIC_ALGERIA = &H5               ' Arabic (Algeria)
    SUBLANG_ARABIC_MOROCCO = &H6               ' Arabic (Morocco)
    SUBLANG_ARABIC_TUNISIA = &H7               ' Arabic (Tunisia)
    SUBLANG_ARABIC_OMAN = &H8                  ' Arabic (Oman)
    SUBLANG_ARABIC_YEMEN = &H9                 ' Arabic (Yemen)
    SUBLANG_ARABIC_SYRIA = &HA                 ' Arabic (Syria)
    SUBLANG_ARABIC_JORDAN = &HB                ' Arabic (Jordan)
    SUBLANG_ARABIC_LEBANON = &HC               ' Arabic (Lebanon)
    SUBLANG_ARABIC_KUWAIT = &HD                ' Arabic (Kuwait)
    SUBLANG_ARABIC_UAE = &HE                   ' Arabic (U.A.E)
    SUBLANG_ARABIC_BAHRAIN = &HF               ' Arabic (Bahrain)
    SUBLANG_ARABIC_QATAR = &H10                ' Arabic (Qatar)
    SUBLANG_CHINESE_TRADITIONAL = &H1          ' Chinese (Taiwan)
    SUBLANG_CHINESE_SIMPLIFIED = &H2           ' Chinese (PR China)
    SUBLANG_CHINESE_HONGKONG = &H3             ' Chinese (Hong Kong)
    SUBLANG_CHINESE_SINGAPORE = &H4            ' Chinese (Singapore)
    SUBLANG_CHINESE_MACAU = &H5                ' Chinese (Macau)
    SUBLANG_DUTCH = &H1                        ' Dutch
    SUBLANG_DUTCH_BELGIAN = &H2                ' Dutch (Belgian)
    SUBLANG_ENGLISH_US = &H1                   ' English (USA)
    SUBLANG_ENGLISH_UK = &H2                   ' English (UK)
    SUBLANG_ENGLISH_AUS = &H3                  ' English (Australian)
    SUBLANG_ENGLISH_CAN = &H4                  ' English (Canadian)
    SUBLANG_ENGLISH_NZ = &H5                   ' English (New Zealand)
    SUBLANG_ENGLISH_EIRE = &H6                 ' English (Irish)
    SUBLANG_ENGLISH_SOUTH_AFRICA = &H7         ' English (South Africa)
    SUBLANG_ENGLISH_JAMAICA = &H8              ' English (Jamaica)
    SUBLANG_ENGLISH_CARIBBEAN = &H9            ' English (Caribbean)
    SUBLANG_ENGLISH_BELIZE = &HA               ' English (Belize)
    SUBLANG_ENGLISH_TRINIDAD = &HB             ' English (Trinidad)
    SUBLANG_ENGLISH_ZIMBABWE = &HC             ' English (Zimbabwe)
    SUBLANG_ENGLISH_PHILIPPINES = &HD          ' English (Philippines)
    SUBLANG_FRENCH = &H1                       ' French
    SUBLANG_FRENCH_BELGIAN = &H2               ' French (Belgian)
    SUBLANG_FRENCH_CANADIAN = &H3              ' French (Canadian)
    SUBLANG_FRENCH_SWISS = &H4                 ' French (Swiss)
    SUBLANG_FRENCH_LUXEMBOURG = &H5            ' French (Luxembourg)
    SUBLANG_FRENCH_MONACO = &H6                ' French (Monaco)
    SUBLANG_GERMAN = &H1                       ' German
    SUBLANG_GERMAN_SWISS = &H2                 ' German (Swiss)
    SUBLANG_GERMAN_AUSTRIAN = &H3              ' German (Austrian)
    SUBLANG_GERMAN_LUXEMBOURG = &H4            ' German (Luxembourg)
    SUBLANG_GERMAN_LIECHTENSTEIN = &H5         ' German (Liechtenstein)
    SUBLANG_ITALIAN = &H1                      ' Italian
    SUBLANG_ITALIAN_SWISS = &H2                ' Italian (Swiss)
    SUBLANG_KOREAN = &H1                       ' Korean (Extended Wansung)
    SUBLANG_KOREAN_JOHAB = &H2                 ' Korean (Johab)
    SUBLANG_LITHUANIAN = &H1                   ' Lithuanian
    SUBLANG_LITHUANIAN_CLASSIC = &H2           ' Lithuanian (Classic)
    SUBLANG_MALAY_MALAYSIA = &H1               ' Malay (Malaysia)
    SUBLANG_MALAY_BRUNEI_DARUSSALAM = &H2      ' Malay (Brunei Darussalam)
    SUBLANG_NORWEGIAN_BOKMAL = &H1             ' Norwegian (Bokmal)
    SUBLANG_NORWEGIAN_NYNORSK = &H2            ' Norwegian (Nynorsk)
    SUBLANG_PORTUGUESE = &H2                   ' Portuguese
    SUBLANG_PORTUGUESE_BRAZILIAN = &H1         ' Portuguese (Brazilian)
    SUBLANG_SERBIAN_LATIN = &H2                ' Serbian (Latin)
    SUBLANG_SERBIAN_CYRILLIC = &H3             ' Serbian (Cyrillic)
    SUBLANG_SPANISH = &H1                      ' Spanish (Castilian)
    SUBLANG_SPANISH_MEXICAN = &H2              ' Spanish (Mexican)
    SUBLANG_SPANISH_MODERN = &H3               ' Spanish (Modern)
    SUBLANG_SPANISH_GUATEMALA = &H4            ' Spanish (Guatemala)
    SUBLANG_SPANISH_COSTA_RICA = &H5           ' Spanish (Costa Rica)
    SUBLANG_SPANISH_PANAMA = &H6               ' Spanish (Panama)
    SUBLANG_SPANISH_DOMINICAN_REPUBLIC = &H7   ' Spanish (Dominican Republic)
    SUBLANG_SPANISH_VENEZUELA = &H8            ' Spanish (Venezuela)
    SUBLANG_SPANISH_COLOMBIA = &H9             ' Spanish (Colombia)
    SUBLANG_SPANISH_PERU = &HA                 ' Spanish (Peru)
    SUBLANG_SPANISH_ARGENTINA = &HB            ' Spanish (Argentina)
    SUBLANG_SPANISH_ECUADOR = &HC              ' Spanish (Ecuador)
    SUBLANG_SPANISH_CHILE = &HD                ' Spanish (Chile)
    SUBLANG_SPANISH_URUGUAY = &HE              ' Spanish (Uruguay)
    SUBLANG_SPANISH_PARAGUAY = &HF             ' Spanish (Paraguay)
    SUBLANG_SPANISH_BOLIVIA = &H10             ' Spanish (Bolivia)
    SUBLANG_SPANISH_EL_SALVADOR = &H11         ' Spanish (El Salvador)
    SUBLANG_SPANISH_HONDURAS = &H12            ' Spanish (Honduras)
    SUBLANG_SPANISH_NICARAGUA = &H13           ' Spanish (Nicaragua)
    SUBLANG_SPANISH_PUERTO_RICO = &H14         ' Spanish (Puerto Rico)
    SUBLANG_SWEDISH = &H1                      ' Swedish
    SUBLANG_SWEDISH_FINLAND = &H2              ' Swedish (Finland)
End Enum


Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Enum LocaleInfoIndexes
      LOCALE_ILANGUAGE = &H1                      '\\ language id
      LOCALE_SLANGUAGE = &H2                      '\\ localized name of language
      LOCALE_SENGLANGUAGE = &H1001                '\\ English name of language
      LOCALE_SABBREVLANGNAME = &H3                '\\ abbreviated language name
      LOCALE_SNATIVELANGNAME = &H4                '\\ native name of language
      LOCALE_ICOUNTRY = &H5                       '\\ country code
      LOCALE_SCOUNTRY = &H6                       '\\ localized name of country
      LOCALE_SENGCOUNTRY = &H1002                 '\\ English name of country
      LOCALE_SABBREVCTRYNAME = &H7                '\\ abbreviated country name
      LOCALE_SNATIVECTRYNAME = &H8                '\\ native name of country
      LOCALE_IDEFAULTLANGUAGE = &H9               '\\ default language id
      LOCALE_IDEFAULTCOUNTRY = &HA                '\\ default country code
      LOCALE_IDEFAULTCODEPAGE = &HB               '\\ default oem code page
      LOCALE_IDEFAULTANSICODEPAGE = &H1004        '\\ default ansi code page
      LOCALE_IDEFAULTMACCODEPAGE = &H1011         '\\ default mac code page
      LOCALE_SLIST = &HC                          '\\ list item separator
      LOCALE_IMEASURE = &HD                       '\\ 0 = metric, 1 = US
      LOCALE_SDECIMAL = &HE                       '\\ decimal separator
      LOCALE_STHOUSAND = &HF                      '\\ thousand separator
      LOCALE_SGROUPING = &H10                     '\\ digit grouping
      LOCALE_IDIGITS = &H11                       '\\ number of fractional digits
      LOCALE_ILZERO = &H12                        '\\ leading zeros for decimal
      LOCALE_INEGNUMBER = &H1010                  '\\ negative number mode
      LOCALE_SNATIVEDIGITS = &H13                 '\\ native ascii 0-9
      LOCALE_SCURRENCY = &H14                     '\\ local monetary symbol
      LOCALE_SINTLSYMBOL = &H15                   '\\ intl monetary symbol
      LOCALE_SMONDECIMALSEP = &H16                '\\ monetary decimal separator
      LOCALE_SMONTHOUSANDSEP = &H17               '\\ monetary thousand separator
      LOCALE_SMONGROUPING = &H18                  '\\ monetary grouping
      LOCALE_ICURRDIGITS = &H19                   '\\ # local monetary digits
      LOCALE_IINTLCURRDIGITS = &H1A               '\\ # intl monetary digits
      LOCALE_ICURRENCY = &H1B                     '\\ positive currency mode
      LOCALE_INEGCURR = &H1C                      '\\ negative currency mode
      LOCALE_SDATE = &H1D                         '\\ date separator
      LOCALE_STIME = &H1E                         '\\ time separator
      LOCALE_SSHORTDATE = &H1F                    '\\ short date format string
      LOCALE_SLONGDATE = &H20                     '\\ long date format string
      LOCALE_STIMEFORMAT = &H1003                 '\\ time format string
      LOCALE_IDATE = &H21                         '\\ short date format ordering
      LOCALE_ILDATE = &H22                        '\\ long date format ordering
      LOCALE_ITIME = &H23                         '\\ time format specifier
      LOCALE_ITIMEMARKPOSN = &H1005               '\\ time marker position
      LOCALE_ICENTURY = &H24                      '\\ century format specifier (short date)
      LOCALE_ITLZERO = &H25                       '\\ leading zeros in time field
      LOCALE_IDAYLZERO = &H26                     '\\ leading zeros in day field (short date)
      LOCALE_IMONLZERO = &H27                     '\\ leading zeros in month field (short date)
      LOCALE_S1159 = &H28                         '\\ AM designator
      LOCALE_S2359 = &H29                         '\\ PM designator
      LOCALE_ICALENDARTYPE = &H1009               '\\ type of calendar specifier
      LOCALE_IOPTIONALCALENDAR = &H100B           '\\ additional calendar types specifier
      LOCALE_IFIRSTDAYOFWEEK = &H100C             '\\ first day of week specifier
      LOCALE_IFIRSTWEEKOFYEAR = &H100D            '\\ first week of year specifier
      LOCALE_SDAYNAME1 = &H2A                     '\\ long name for Monday
      LOCALE_SDAYNAME2 = &H2B                     '\\ long name for Tuesday
      LOCALE_SDAYNAME3 = &H2C                     '\\ long name for Wednesday
      LOCALE_SDAYNAME4 = &H2D                     '\\ long name for Thursday
      LOCALE_SDAYNAME5 = &H2E                     '\\ long name for Friday
      LOCALE_SDAYNAME6 = &H2F                     '\\ long name for Saturday
      LOCALE_SDAYNAME7 = &H30                     '\\ long name for Sunday
      LOCALE_SABBREVDAYNAME1 = &H31               '\\ abbreviated name for Monday
      LOCALE_SABBREVDAYNAME2 = &H32               '\\ abbreviated name for Tuesday
      LOCALE_SABBREVDAYNAME3 = &H33               '\\ abbreviated name for Wednesday
      LOCALE_SABBREVDAYNAME4 = &H34               '\\ abbreviated name for Thursday
      LOCALE_SABBREVDAYNAME5 = &H35               '\\ abbreviated name for Friday
      LOCALE_SABBREVDAYNAME6 = &H36               '\\ abbreviated name for Saturday
      LOCALE_SABBREVDAYNAME7 = &H37               '\\ abbreviated name for Sunday
      LOCALE_SMONTHNAME1 = &H38                   '\\ long name for January
      LOCALE_SMONTHNAME2 = &H39                   '\\ long name for February
      LOCALE_SMONTHNAME3 = &H3A                   '\\ long name for March
      LOCALE_SMONTHNAME4 = &H3B                   '\\ long name for April
      LOCALE_SMONTHNAME5 = &H3C                   '\\ long name for May
      LOCALE_SMONTHNAME6 = &H3D                   '\\ long name for June
      LOCALE_SMONTHNAME7 = &H3E                   '\\ long name for July
      LOCALE_SMONTHNAME8 = &H3F                   '\\ long name for August
      LOCALE_SMONTHNAME9 = &H40                   '\\ long name for September
      LOCALE_SMONTHNAME10 = &H41                  '\\ long name for October
      LOCALE_SMONTHNAME11 = &H42                  '\\ long name for November
      LOCALE_SMONTHNAME12 = &H43                  '\\ long name for December
      LOCALE_SMONTHNAME13 = &H100E                '\\ long name for 13th month (if exists)
      LOCALE_SABBREVMONTHNAME1 = &H44             '\\ abbreviated name for January
      LOCALE_SABBREVMONTHNAME2 = &H45             '\\ abbreviated name for February
      LOCALE_SABBREVMONTHNAME3 = &H46             '\\ abbreviated name for March
      LOCALE_SABBREVMONTHNAME4 = &H47             '\\ abbreviated name for April
      LOCALE_SABBREVMONTHNAME5 = &H48             '\\ abbreviated name for May
      LOCALE_SABBREVMONTHNAME6 = &H49             '\\ abbreviated name for June
      LOCALE_SABBREVMONTHNAME7 = &H4A             '\\ abbreviated name for July
      LOCALE_SABBREVMONTHNAME8 = &H4B             '\\ abbreviated name for August
      LOCALE_SABBREVMONTHNAME9 = &H4C             '\\ abbreviated name for September
      LOCALE_SABBREVMONTHNAME10 = &H4D            '\\ abbreviated name for October
      LOCALE_SABBREVMONTHNAME11 = &H4E            '\\ abbreviated name for November
      LOCALE_SABBREVMONTHNAME12 = &H4F            '\\ abbreviated name for December
      LOCALE_SABBREVMONTHNAME13 = &H100F          '\\ abbreviated name for 13th month (if exists)
      LOCALE_SPOSITIVESIGN = &H50                 '\\ positive sign
      LOCALE_SNEGATIVESIGN = &H51                 '\\ negative sign
      LOCALE_IPOSSIGNPOSN = &H52                  '\\ positive sign position
      LOCALE_INEGSIGNPOSN = &H53                  '\\ negative sign position
      LOCALE_IPOSSYMPRECEDES = &H54               '\\ mon sym precedes pos amt
      LOCALE_IPOSSEPBYSPACE = &H55                '\\ mon sym sep by space from pos amt
      LOCALE_INEGSYMPRECEDES = &H56               '\\ mon sym precedes neg amt
      LOCALE_INEGSEPBYSPACE = &H57                '\\ mon sym sep by space from neg amt
      LOCALE_FONTSIGNATURE = &H58                 '\\ font signature
      LOCALE_SISO639LANGNAME = &H59               '\\ ISO abbreviated language name
      LOCALE_SISO3166CTRYNAME = &H5A              '\\ ISO abbreviated country name
      LOCALE_IDEFAULTEBCDICCODEPAGE = &H1012      '\\ default ebcdic code page
      LOCALE_IPAPERSIZE = &H100A                  '\\ 0 = letter, 1 = a4, 2 = legal, 3 = a3
      LOCALE_SENGCURRNAME = &H1007                '\\ english name of currency
      LOCALE_SNATIVECURRNAME = &H1008             '\\ native name of currency
      LOCALE_SYEARMONTH = &H1006                  '\\ year month format string
      LOCALE_SSORTNAME = &H1013                   '\\ sort name
      LOCALE_IDIGITSUBSTITUTION = &H1014          '\\ 0 = none, 1 = context, 2 = native digit
End Enum


Public Enum vbMonthOfYear
    vbJanuary = 1
    vbFebruary = 2
    vbMarch = 3
    vbApril = 4
    vbMay = 5
    vbJune = 6
    vbJuly = 7
    vbAugust = 8
    vbSeptember = 9
    vbOctober = 10
    vbNovember = 11
    vbDecember = 12
    vbThirteenthMonth = 13
End Enum

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Get DayName(ByVal DayOfWeek As VbDayOfWeek) As String
Attribute DayName.VB_HelpID = 2774

Dim index As LocaleInfoIndexes

Select Case DayOfWeek
Case vbMonday
    index = LOCALE_SDAYNAME1
Case vbTuesday
    index = LOCALE_SDAYNAME2
Case vbWednesday
    index = LOCALE_SDAYNAME3
Case vbThursday
    index = LOCALE_SDAYNAME4
Case vbFriday
    index = LOCALE_SDAYNAME5
Case vbSaturday
    index = LOCALE_SDAYNAME6
Case vbSunday
    index = LOCALE_SDAYNAME7
End Select

DayName = GetLocaleInfoItem(index)

End Property


Public Property Get AbbreviatedDayName(ByVal DayOfWeek As VbDayOfWeek) As String
Attribute AbbreviatedDayName.VB_HelpID = 2775

Dim index As LocaleInfoIndexes

Select Case DayOfWeek
Case vbMonday
    index = LOCALE_SABBREVDAYNAME1
Case vbTuesday
    index = LOCALE_SABBREVDAYNAME2
Case vbWednesday
    index = LOCALE_SABBREVDAYNAME3
Case vbThursday
    index = LOCALE_SABBREVDAYNAME4
Case vbFriday
    index = LOCALE_SABBREVDAYNAME5
Case vbSaturday
    index = LOCALE_SABBREVDAYNAME6
Case vbSunday
    index = LOCALE_SABBREVDAYNAME7
End Select

AbbreviatedDayName = GetLocaleInfoItem(index)

End Property

Public Property Get DecimalSeparator() As String
Attribute DecimalSeparator.VB_HelpID = 2776

    DecimalSeparator = GetLocaleInfoItem(LOCALE_SDECIMAL)
    
End Property

Public Property Get EnglishCountryName() As String
Attribute EnglishCountryName.VB_HelpID = 2777

    EnglishCountryName = GetLocaleInfoItem(LOCALE_SENGCOUNTRY)
    
End Property

Public Property Get EnglishCurrencyName() As String
Attribute EnglishCurrencyName.VB_HelpID = 2778

    EnglishCurrencyName = GetLocaleInfoItem(LOCALE_SENGCURRNAME)
    
End Property

Public Property Get EnglishLanguageName() As String
Attribute EnglishLanguageName.VB_HelpID = 2779

    EnglishLanguageName = GetLocaleInfoItem(LOCALE_SENGLANGUAGE)
    
End Property

Private Function GetLocaleInfoItem(ByVal index As LocaleInfoIndexes) As String

Dim sBuff As String
Dim lRet As Long

sBuff = String$(1024, 0)
lRet = GetLocaleInfo(mLocaleId, index, sBuff, Len(sBuff))
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiLocale:GetItemInfo - " & index, GetLastSystemError
Else
    If lRet > 0 Then
        sBuff = Left$(sBuff, lRet)
        If InStr(sBuff, Chr$(0)) > 1 Then
            sBuff = Left$(sBuff, InStr(sBuff, Chr$(0)) - 1)
        End If
    End If
End If

GetLocaleInfoItem = sBuff

End Function


Public Property Get LocalCountryName() As String
Attribute LocalCountryName.VB_HelpID = 2780

    LocalCountryName = GetLocaleInfoItem(LOCALE_SCOUNTRY)
    
End Property

Public Property Get LocalCurrencyName() As String
Attribute LocalCurrencyName.VB_HelpID = 2781

    LocalCurrencyName = GetLocaleInfoItem(LOCALE_SNATIVECURRNAME)
    
End Property

Public Property Get LocaleId() As Long
Attribute LocaleId.VB_HelpID = 2782

    LocaleId = mLocaleId
    
End Property

Public Property Let LocaleId(ByVal newId As Long)

If mLocaleId <> newId Then
    mLocaleId = newId
End If

End Property

Public Property Get LocalLanguageName() As String
Attribute LocalLanguageName.VB_HelpID = 2783

    LocalLanguageName = GetLocaleInfoItem(LOCALE_SLANGUAGE)
    
End Property

Public Property Get Mainlanguage() As enLocaleMainLanguage
Attribute Mainlanguage.VB_HelpID = 2784

    Mainlanguage = APIDispenser.GetBits(mLocaleId, 0, 9)

End Property



Public Property Get MonthName(ByVal MonthOfYear As vbMonthOfYear) As String
Attribute MonthName.VB_HelpID = 2785

Dim index As LocaleInfoIndexes

index = LOCALE_SMONTHNAME1 + (MonthOfYear - 1)
MonthName = GetLocaleInfoItem(index)

End Property

Public Property Get AbbreviatedMonthName(ByVal MonthOfYear As vbMonthOfYear) As String
Attribute AbbreviatedMonthName.VB_HelpID = 2786

Dim index As LocaleInfoIndexes

index = LOCALE_SABBREVMONTHNAME1 + (MonthOfYear - 1)
AbbreviatedMonthName = GetLocaleInfoItem(index)

End Property
Public Property Get SortOrder() As enLocaleSortOrders
Attribute SortOrder.VB_HelpID = 2787

    SortOrder = APIDispenser.GetBits(mLocaleId, 16, 19)

End Property

Public Property Get Sublanguage() As enLocaleSublanguages
Attribute Sublanguage.VB_HelpID = 2788

Dim mSubLang As Long

mSubLang = APIDispenser.GetBits(mLocaleId, 10, 15)
'\\ Sub language depens on what the main language is....
Select Case Mainlanguage
Case LANG_ARABIC
    Select Case mSubLang
    Case &H1
        Sublanguage = SUBLANG_ARABIC_SAUDI_ARABIA
    Case &H2
        Sublanguage = SUBLANG_ARABIC_IRAQ
    Case &H3
        Sublanguage = SUBLANG_ARABIC_EGYPT
    Case &H4
        Sublanguage = SUBLANG_ARABIC_LIBYA
    Case &H5
        Sublanguage = SUBLANG_ARABIC_ALGERIA
    Case &H6
        Sublanguage = SUBLANG_ARABIC_MOROCCO
    Case &H7
        Sublanguage = SUBLANG_ARABIC_TUNISIA
    Case &H8
        Sublanguage = SUBLANG_ARABIC_OMAN
    Case &H9
        Sublanguage = SUBLANG_ARABIC_YEMEN
    Case &HA
        Sublanguage = SUBLANG_ARABIC_SYRIA
    Case &HB
        Sublanguage = SUBLANG_ARABIC_JORDAN
    Case &HC
        Sublanguage = SUBLANG_ARABIC_LEBANON
    Case &HD
        Sublanguage = SUBLANG_ARABIC_KUWAIT
    Case &HE
        Sublanguage = SUBLANG_ARABIC_UAE
    Case &HF
        Sublanguage = SUBLANG_ARABIC_BAHRAIN
    Case &H10
        Sublanguage = SUBLANG_ARABIC_QATAR
    End Select
Case LANG_CHINESE
    Select Case mSubLang
    Case &H1
        Sublanguage = SUBLANG_CHINESE_TRADITIONAL
    Case &H2
        Sublanguage = SUBLANG_CHINESE_SIMPLIFIED
    Case &H3
        Sublanguage = SUBLANG_CHINESE_HONGKONG
    Case &H4
        Sublanguage = SUBLANG_CHINESE_SINGAPORE
    Case &H5
        Sublanguage = SUBLANG_CHINESE_MACAU
    End Select
Case LANG_DUTCH
    Select Case mSubLang
    Case &H1
        Sublanguage = SUBLANG_DUTCH
    Case &H2
        Sublanguage = SUBLANG_DUTCH_BELGIAN
    End Select
Case LANG_ENGLISH
    Select Case mSubLang
    Case &H1 ' English (USA)
        Sublanguage = SUBLANG_ENGLISH_US
    Case &H2 ' English (UK)
        Sublanguage = SUBLANG_ENGLISH_US
    Case &H3 ' English (Australian)
        Sublanguage = SUBLANG_ENGLISH_AUS
    Case &H4 ' English (Canadian)
        Sublanguage = SUBLANG_ENGLISH_CAN
    Case &H5 ' English (New Zealand)
        Sublanguage = SUBLANG_ENGLISH_NZ
    Case &H6 ' English (Irish)
        Sublanguage = SUBLANG_ENGLISH_EIRE
    Case &H7 ' English (South Africa)
        Sublanguage = SUBLANG_ENGLISH_SOUTH_AFRICA
    Case &H8 ' English (Jamaica)
        Sublanguage = SUBLANG_ENGLISH_JAMAICA
    Case &H9 ' English (Caribbean)
        Sublanguage = SUBLANG_ENGLISH_CARIBBEAN
    Case &HA ' English (Belize)
        Sublanguage = SUBLANG_ENGLISH_BELIZE
    Case &HB ' English (Trinidad)
        Sublanguage = SUBLANG_ENGLISH_TRINIDAD
    Case &HC ' English (Zimbabwe)
        Sublanguage = SUBLANG_ENGLISH_ZIMBABWE
    Case &HD ' English (Philippines)
        Sublanguage = SUBLANG_ENGLISH_PHILIPPINES
    End Select
Case LANG_FRENCH
    Select Case mSubLang
    Case &H1
        Sublanguage = SUBLANG_FRENCH
    Case &H2
        Sublanguage = SUBLANG_FRENCH_BELGIAN
    Case &H3
        Sublanguage = SUBLANG_FRENCH_CANADIAN
    Case &H4
        Sublanguage = SUBLANG_FRENCH_SWISS
    Case &H5
        Sublanguage = SUBLANG_FRENCH_LUXEMBOURG
    Case &H6
        Sublanguage = SUBLANG_FRENCH_MONACO
    End Select
Case LANG_GERMAN
    Select Case mSubLang
    Case &H1
        Sublanguage = SUBLANG_GERMAN
    Case &H2
        Sublanguage = SUBLANG_GERMAN_SWISS
    Case &H3
        Sublanguage = SUBLANG_GERMAN_AUSTRIAN
    Case &H4
        Sublanguage = SUBLANG_GERMAN_LUXEMBOURG
    Case &H5
        Sublanguage = SUBLANG_GERMAN_LIECHTENSTEIN
    End Select
Case Else
    Sublanguage = mSubLang
End Select

End Property


Public Property Get ThousandSeparator() As String
Attribute ThousandSeparator.VB_HelpID = 2789

    ThousandSeparator = GetLocaleInfoItem(LOCALE_STHOUSAND)
    
End Property


