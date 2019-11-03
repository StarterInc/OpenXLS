/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.formats.XLS

import java.awt.Color

/**
 * /** Constants relevant to XLS Formatting
 *
 *
 * NOTE:
 * Excel automatically assigns
 * If you type     this number format
 * -------------------------------------------
 *
 *
 * 1.0             General
 * 1.123           General
 * 1.1%            0.00%
 * 1.1E+2          0.00E+00
 * 1 1/2           # ?/?
 * $1.11           Currency, 2 decimal places
 * 1/1/01          Date
 * 1:10            Time
 *
 *
 * CUSTOM FORMATS:
 *
 *
 * Each format that you create can have up to three sections for numbers and a fourth section for text.
 * <POSITIVE>;<NEGATIVE>;<ZERO>;<TEXT>
 *
 * @see CustomFormatHelper
</TEXT></ZERO></NEGATIVE></POSITIVE> */
interface FormatConstants {
companion object {


val DEFAULT_FONT_WEIGHT = 200
val DEFAULT_FONT_SIZE = 20
val DEFAULT_FONT_FACE = "Arial"


// Font Weights
val PLAIN = 400
val BOLD = 700

//	 Font UnderLine Style constants
val STYLE_UNDERLINE_NONE:Byte = 0x0
val STYLE_UNDERLINE_SINGLE:Byte = 0x1
val STYLE_UNDERLINE_DOUBLE:Byte = 0x2
val STYLE_UNDERLINE_SINGLE_ACCTG:Byte = 0x21
val STYLE_UNDERLINE_DOUBLE_ACCTG:Byte = 0x22


// Color Constants
val Black = Color(0, 0, 0)
val Brown = Color(153, 51, 0)
val OliveGreen = Color(51, 51, 0)
val Dark_Green = Color(0, 51, 0)
val Dark_Teal = Color(0, 51, 102)
val Dark_Blue = Color(0, 0, 128)
val Indigo = Color(51, 51, 153)
val Gray80 = Color(51, 51, 51)
// could we be wrong??  test it per claritas
// public static Color Gray50			= new Color(110,110,110);
// Excel 07 likes real 50 percent 			= new Color(127,127,127);
val Gray50 = Color(136, 136, 136)

val Dark_Red = Color(128, 0, 0)
val Orange = Color(255, 102, 0)
val Dark_Yellow = Color(128, 128, 0)
val Green = Color(0, 128, 0)
val Teal = Color(0, 128, 128)
val Blue = Color(0, 0, 255)
val BlueGray = Color(102, 102, 153)

val Red = Color(255, 0, 0)
val Light_Orange = Color(255, 153, 0)
val Lime = Color(153, 204, 0)
val SeaGreen = Color(51, 153, 102)
val Aqua = Color(51, 204, 204)
val Light_Blue = Color(0, 128, 255)
val Violet = Color(128, 0, 128)
val Gray40 = Color(150, 150, 150)

val Pink = Color(255, 0, 255)
val Gold = Color(255, 204, 0)
val Yellow = Color(255, 255, 0)
val BrightGreen = Color(0, 255, 0)
val Turquoise = Color(0, 255, 255)
val SkyBlue = Color(0, 204, 255)
val Plum = Color(153, 51, 102)
val Gray25 = Color(192, 192, 192)

val Rose = Color(255, 153, 204)
val Tan = Color(255, 204, 153)
val Light_Yellow = Color(255, 255, 153)
val Light_Green = Color(204, 255, 204)
val Light_Turquoise = Color(204, 255, 255)
val PaleBlue = Color(153, 204, 255)
val Lavender = Color(204, 153, 255)
val White = Color(255, 255, 255)
val Gray15 = Color(216, 216, 216)
val Dark_Purple = Color(102, 0, 102)
val Salmon = Color(255, 128, 128)
val Light_Purple = Color(204, 204, 255)
val Medium_Purple = Color(153, 153, 255)

/**
 * Excel's basic icv color table.  This may be altered by the Palette record.
 * See WorkBook.CCLORTABLE
*/
val COLORTABLE = arrayOf<Color>(FormatConstants.Black, FormatConstants.White, FormatConstants.Red, FormatConstants.BrightGreen, /// WRONG, was duplicate, incorrect GREEN
FormatConstants.Blue, FormatConstants.Yellow, FormatConstants.Pink, FormatConstants.Turquoise, FormatConstants.Black, /**
 * NOTE: from documentation:
 * if palette record exists 8-63 (0x3f) gets read from it,
 * otherwise If a Palette record exists in this file, these icv values
 * specify colors from the rgColor array in the Palette record.
*/
/*
 * Color 9 === AUTOMATIC color
 * the opposite of what is normal; for instance, background color
 * index 9= black, font color index 9= white. It is set here to
 * black, and in Fonts.getColor, a special case is put for icv= 9
*/
FormatConstants.White, // 9
FormatConstants.Red, // 10
FormatConstants.BrightGreen, FormatConstants.Blue, FormatConstants.Yellow, FormatConstants.Pink, FormatConstants.Turquoise, // 15
FormatConstants.Dark_Red, // Burgundy //16
FormatConstants.Green, FormatConstants.Dark_Blue, // purple
FormatConstants.Dark_Yellow, FormatConstants.Violet, // 20 dark pink
FormatConstants.Teal, FormatConstants.Gray25, FormatConstants.Gray50, FormatConstants.Medium_Purple, FormatConstants.Plum, // 25
FormatConstants.Light_Yellow, FormatConstants.Light_Turquoise, FormatConstants.Dark_Purple, FormatConstants.Salmon, FormatConstants.BlueGray, // 30
FormatConstants.Light_Purple, FormatConstants.Dark_Blue, FormatConstants.Pink, FormatConstants.Yellow, FormatConstants.Turquoise, // 35
FormatConstants.Violet, FormatConstants.Dark_Red, FormatConstants.Teal, // ////////////WRONG - WAS DUPLICATE
FormatConstants.Blue, FormatConstants.SkyBlue, // 40
FormatConstants.Light_Turquoise, FormatConstants.Light_Green, FormatConstants.Light_Yellow, FormatConstants.PaleBlue, FormatConstants.Rose, // 45
FormatConstants.Lavender, FormatConstants.Tan, FormatConstants.Dark_Yellow, FormatConstants.Aqua, FormatConstants.Lime, // 50
FormatConstants.Gold, FormatConstants.Light_Orange, FormatConstants.Orange, FormatConstants.BlueGray, FormatConstants.Gray40, // 55
FormatConstants.Dark_Teal, FormatConstants.SeaGreen, FormatConstants.Dark_Green, FormatConstants.OliveGreen, FormatConstants.Brown, // 60
FormatConstants.Plum, FormatConstants.Indigo, FormatConstants.Gray80, FormatConstants.White, // 64
FormatConstants.Black)// 65 WHY DO THESE CHANGE???

// Formatting Color Constants
val COLOR_BLACK = 0
val COLOR_WHITE = 1
val COLOR_RED = 2
val COLOR_BLUE = 4
val COLOR_YELLOW = 5
val COLOR_BLACK3 = 8
val COLOR_WHITE3 = 9
val COLOR_RED_CHART = 10    // chart fills use SPECIFIC icv numbers -- incorrect # results in black fill
val COLOR_BRIGHT_GREEN = 11
val COLOR_YELLOW_CHART = 13
val COLOR_DARK_RED = 16
val COLOR_GREEN = 17
val COLOR_DARK_BLUE = 18
val COLOR_OLIVE_GREEN_CHART = 19
val COLOR_DARK_YELLOW = 19
val COLOR_VIOLET = 20
val COLOR_TEAL = 21
val COLOR_GRAY25 = 22
val COLOR_GRAY50 = 23
val COLOR_MEDIUM_PURPLE = 24
val COLOR_PLUM = 25
val COLOR_LIGHT_YELLOW = 26
val COLOR_LIGHT_TURQUOISE = 27
val COLOR_DARK_PURPLE = 28
val COLOR_SALMON = 29
val COLOR_LIGHT_BLUE = 30
val COLOR_LIGHT_PURPLE = 31
val COLOR_PINK = 33
val COLOR_CYAN = 35
val COLOR_BLUE_CHART = 39
val COLOR_SKY_BLUE = 40

val COLOR_LIGHT_GREEN = 42
val COLOR_PALE_BLUE = 44
val COLOR_ROSE = 45
val COLOR_LAVENDER = 46
val COLOR_TAN = 47
val COLOR_TURQUOISE = 48
val COLOR_AQUA = 49
val COLOR_LIME = 50
val COLOR_GOLD = 51
val COLOR_LIGHT_ORANGE = 52
val COLOR_ORANGE = 53
val COLOR_BLUE_GRAY = 54
val COLOR_GRAY40 = 55
val COLOR_DARK_TEAL = 56
val COLOR_SEA_GREEN = 57
val COLOR_DARK_GREEN = 58
val COLOR_OLIVE_GREEN = 59
val COLOR_BROWN = 60
val COLOR_INDIGO = 62
val COLOR_GRAY80 = 63
val COLOR_WHITE2 = 64
val COLOR_BLACK2 = 65

/**
 * given color int ala COLORTABLE, interpret Excel Color name
*/
val COLORNAMES = arrayOf<String>("Black", "White", "Red", "BrightGreen", "Blue", "Yellow", "Pink", "Turquoise", "Black", "White", "Red", //10
"BrightGreen", "Blue", "Yellow", "Pink", "Turquoise", // 15
"Dark_Red", //Burgundy //16
"Green", "Dark_Blue", //purple
"Dark_Yellow", "Violet", //20 dark pink
"Teal", "Gray25", "Gray50", "MediumPurple", "Plum", //25
"Light_Yellow", "Light_Turquoise", "Dark_Purple", "Salmon", "BlueGray", //30
"Light_Purple", "Dark_Blue", "Pink", "Yellow", "Turquoise", // 35
"Violet", "Dark_Red", "Teal", "Blue", "SkyBlue", //40
"Light_Turquoise", "Light_Green", "Light_Yellow", "PaleBlue", "Rose", //45
"Lavender", "Tan", "Blue", "Aqua", "Lime", // 50
"Gold", "Light_Orange", "Orange", "BlueGray", "Gray40", //55
"Dark_Teal", "SeaGreen", "Dark_Green", "OliveGreen", "Brown", //60
"Plum", "Indigo", // WRONG
"Gray80", "White", // 64
"Black")// 65

/**
 * given color int ala COLORTABLE, interpret HTML Color name
*/
val HTMLCOLORNAMES = arrayOf<String>("Black", "White", "Red", "Green", "Blue", "Yellow", "Pink", "Turquoise", "Black", "White", "Red", //10
"Lime", //	BrightGreen",
"Blue", "Yellow", "Pink", "Turquoise", // 15
"Dark_Red", "Green", "DarkBlue", "DarkKhaki", //Dark_Yellow",
"Violet", //20
"Teal", "LightGray", // Gray25
"DarkGray", // Gray50"
"DodgerBlue", "DarkRed", //25
"LightYellow", "PaleTurquoise", // Light_Turquoise",
"DarkMagenta", "Red", "CadetBlue", // BlueGray", //30
"PowderBlue", // PaleBlue",
"DarkBlue", "Pink", "Yellow", "Turquoise", // 35
"Plum", "Brown", "SeaGreen", "Blue", "SkyBlue", //40
"PaleTurquoise", //Light_Turquoise",
"LightGreen", "LightYellow", "PowderBlue", // PaleBlue",
"Crimson", // Rose", //45
"Lavender", "Tan", "Blue", "Aqua", "Lime", // 50
"Gold", "Orange", //Light_Orange",
"Orange", "CadetBlue", // BlueGray",
"YellowGreen", // Gray40", //55
"DarkSlateBlue", //Dark_Teal",
"SeaGreen", "DarkGreen", "OliveDrab", //OliveGreen",
"Brown", //60
"Plum", "Indigo", // WRONG
"DimGray", // Gray80
"White", // 64
"Black")// 65

val SVGCOLORSTRINGS = arrayOf<String>("black", "white", "red", "lime", // = BrightGreen
"blue", "yellow", // 5
"fuchsia", // = Pink
"aqua", // = Turquoise
"black", "white", "red", // 10
"lime", "blue", "yellow", "fuchsia", "aqua", // 15
"maroon", // = Dark_Red
"green", "navy", // = Dark_Blue
"olive", // = Dark_Yellow
"purple", // = Violet // 20
"teal", "silver", // = Gray25
"gray", // = Gray50
"lightblue", //lightsteelblue",
"palevioletred", //mediumorchid",	// =Plum 	// 25
"darkolivegreen", "aqua", // = Light_Turqouise
"purple", // = Dark_Purple
"salmon", "slateblue", // = BlueGray		// 30
"aliceblue", // = Light_Purple
"navy", // = Dark_Blue
"fuchsia", // = Pink
"yellow", "aqua", // = Turquoise //35
"purple", // = Violet
"maroon", // = Dark_Red
"teal", "blue", "deepskyblue", // = SkyBlue				// 40
"azure", // = Light_Turquoise
"honeydew", // = Light_Green
"lightyellow", "powderblue", // = PaleBlue
"lightpink", // = Rose			// 45
"aliceblue", // = Lavendar
"navajowhite", // = Tan
"blue", "mediumturquoise", // = Aqua
"lawngreen", // = Lime	// 50
"gold", "orange", "darkorange", "slateblue", // = BlueGray
"darkgray", // = Gray40		//55
"midnightblue", // = Dark_Teal
"seagreen", "darkgreen", "darkolivegreen", "saddlebrown", // 60
"mediumvioletred", "darkslateblue", //= Indigo
"darkslategray", // = Gray80
"white", // 64
"black")// 65

// Formatting Pattern Constants
val PATTERN_NONE = 0
val PATTERN_FILLED = 1
val PATTERN_LIGHT_FILL = 2
val PATTERN_MED_FILL = 3
val PATTERN_HVY_FILL = 4
val PATTERN_HOR_STRIPES1 = 5
val PATTERN_VERT_STRIPES1 = 6
val PATTERN_DIAG_STRIPES1 = 7
val PATTERN_DIAG_STRIPES2 = 8
val PATTERN_CHECKERBOARD1 = 9
val PATTERN_CHECKERBOARD2 = 10
val PATTERN_HOR_STRIPES2 = 11
val PATTERN_VERT_STRIPES2 = 12
val PATTERN_DIAG_STRIPES3 = 13
val PATTERN_DIAG_STRIPES4 = 14
val PATTERN_GRID1 = 15
val PATTERN_CROSSPATCH1 = 16
val PATTERN_MED_DOTS = 17
val PATTERN_LIGHT_DOTS = 18
val PATTERN_HVY_DOTS = 19
val PATTERN_DIAG_STRIPES5 = 20
val PATTERN_DIAG_STRIPES6 = 21
val PATTERN_DIAG_STRIPES7 = 22
val PATTERN_DIAG_STRIPES8 = 23
val PATTERN_VERT_STRIPES3 = 24
val PATTERN_HOR_STRIPES3 = 25
val PATTERN_VERT_STRIPES4 = 26
val PATTERN_HOR_STRIPES4 = 27
val PATTERN_PATCHY1 = 28
val PATTERN_PATCHY2 = 29
val PATTERN_PATCHY3 = 30
val PATTERN_PATCHY4 = 31

// Border line styles

val BORDER_NONE:Short = 0x0
val BORDER_THIN:Short = 0x1
val BORDER_MEDIUM:Short = 0x2
val BORDER_DASHED:Short = 0x3
val BORDER_DOTTED:Short = 0x4
val BORDER_THICK:Short = 0x5
val BORDER_DOUBLE:Short = 0x6
val BORDER_HAIR:Short = 0x7
val BORDER_MEDIUM_DASHED:Short = 0x8
val BORDER_DASH_DOT:Short = 0x9
val BORDER_MEDIUM_DASH_DOT:Short = 0xA
val BORDER_DASH_DOT_DOT:Short = 0xB
val BORDER_MEDIUM_DASH_DOT_DOT:Short = 0xC
val BORDER_SLANTED_DASH_DOT:Short = 0xD

val BORDER_NAMES = arrayOf<String>("None", "Thin", "Medium", "Dashed", "Dotted", "Thick", "Double", "Hair", "Medium dashed", "Dash-dot", "Medium dash-dot", "Dash-dot-dot", "Medium dash-dot-dot", "Slanted dash-dot")

// interpret border sizes from specifications above
val BORDER_SIZES_HTML = arrayOf<String>("", "2px", // thin
"3px", // medium
"2px", "2px", "4px", // thick
"2ox", "1px", // hair
"3px", // medium dashed
"2px", "2px", "2px", "2px", "2px")

val BORDER_NAMES_HTML = arrayOf<String>("", "solid", // thin
"solid", // medium
"dashed", "dotted", "solid", // thick
"double", "solid", // hair
"dashed", "dashed", "dashed", "dotted", "dotted", "dashed")

val BORDER_STYLES_JSON = arrayOf<String>("none", // 0x0 No border
"thin", // 0x1 Thin line
"medium", // 0x2 Medium line
"dashed", // 0x3 Dashed line
"dotted", // 0x4 Dotted line
"thick", // 0x5 Thick line
"double", // 0x6 Double line
"hairline", // 0x7 Hairline
"medium_dashed", // 0x8 Medium dashed line
"dash-dot", // 0x9 Dash-dot line
"medium_dash-dot", // 0xA Medium dash-dot line
"dash-dot-dot", // 0xB Dash-dot-dot line
"medium_dash-dot-dot", // 0xC Medium dash-dot-dot line
"slant_dash-dot-dot"    // 0xD Slanted dash-dot-dot line
)

val FORMAT_SUBSCRIPT = 2
val FORMAT_SUPERSCRIPT = 1
val FORMAT_NOSCRIPT = 0

// Decoded Built-in Format Patterns + HEX Format IDs
// includes number, currency and date formats
/**
 * Please use the getBuiltinFormats in FormatConstantsImpl for locale specific formatting!
*/
val BUILTIN_FORMATS = arrayOf<Array<String>>(arrayOf<String>("General", "0"), arrayOf<String>("0", "1"), arrayOf<String>("0.00", "2"), arrayOf<String>("#,##0", "3"), arrayOf<String>("#,##0.00", "4"), arrayOf<String>("$#,##0;($#,##0)", "5"), arrayOf<String>("$#,##0;[Red]($#,##0)", "6"), arrayOf<String>("$#,##0.00;($#,##0.00)", "7"), arrayOf<String>("$#,##0.00;[Red]($#,##0.00)", "8"), arrayOf<String>("0%", "9"), arrayOf<String>("0.00%", "a"), arrayOf<String>("0.00E+00", "b"), arrayOf<String>("# ?/?", "c"), arrayOf<String>("# ??/??", "d"),

// dates
arrayOf<String>("mm-dd-yy", "E"), // REGIONAL FORMAT!	m/d/yyyy  or m/d/yy
arrayOf<String>("d-mmm-yy", "F"), arrayOf<String>("d-mmm", "10"), arrayOf<String>("mmm-yy", "11"), arrayOf<String>("h:mm AM/PM", "12"), arrayOf<String>("h:mm:ss AM/PM", "13"), arrayOf<String>("h:mm", "14"), arrayOf<String>("h:mm:ss", "15"), arrayOf<String>("m/d/yy h:mm", "16"),

// 17h through 24h are undocumented international formats!!!
arrayOf<String>("reserved", "17"), arrayOf<String>("reserved", "18"), arrayOf<String>("reserved", "19"), arrayOf<String>("reserved", "1a"), arrayOf<String>("reserved", "1b"), arrayOf<String>("reserved", "1c"), arrayOf<String>("reserved", "1d"), arrayOf<String>("reserved", "1e"), arrayOf<String>("reserved", "1f"), arrayOf<String>("reserved", "20"), arrayOf<String>("reserved", "21"), arrayOf<String>("reserved", "22"), arrayOf<String>("reserved", "23"), arrayOf<String>("reserved", "24"),

// more currency
arrayOf<String>("(#,##0_);(#,##0)", "25"), arrayOf<String>("(#,##0_);[Red](#,##0)", "26"), arrayOf<String>("(#,##0.00_);(#,##0.00)", "27"), arrayOf<String>("(#,##0.00_);[Red](#,##0.00)", "28"), arrayOf<String>("_(*#,##0_);_(*(#,##0);_(*\"-\"_);_(@_)", "29"), arrayOf<String>("_($* #,##0_);_($* (#,##0);_($*\"-\"_);_(@_)", "2a"), arrayOf<String>("_(* #,##0.00_);_(* (#,##0.00);_(*\"-\"??_);_(@_)", "2b"), arrayOf<String>("_($* #,##0.00;_($* (#,##0.00);_($* \"-\"??;_(@_)", "2c"),

// more dates
arrayOf<String>("mm:ss", "2D"), arrayOf<String>("[h]:mm:ss", "2E"), arrayOf<String>("mm:ss.0", "2F"),

// misc
arrayOf<String>("##0.0E+0", "30"), arrayOf<String>("@", "31"))


val BUILTIN_FORMATS_JP = arrayOf<Array<String>>(arrayOf<String>("General", "0"), arrayOf<String>("0", "1"), arrayOf<String>("0.00", "2"), arrayOf<String>("#,##0", "3"), arrayOf<String>("#,##0.00", "4"), arrayOf<String>("\u00a5#,##0;\u00a5-#,##0", "5"), arrayOf<String>("\u00a5#,##0;[Red]\u00a5-#,##0", "6"), arrayOf<String>("\u00a5#,##0.00;\u00a5-#,##0.00", "7"), arrayOf<String>("\u00a5#,##0.00;[Red]\u00a5-#,##0.00", "8"), arrayOf<String>("0%", "9"), arrayOf<String>("0.00%", "a"), arrayOf<String>("0.00E+00", "b"), arrayOf<String>("# ?/?", "c"), arrayOf<String>("# ??/??", "d"),

// dates
arrayOf<String>("yyyy/m/d", "e"), arrayOf<String>("d-mmm-yy", "f"), arrayOf<String>("d-mmm", "10"), arrayOf<String>("mmm-yy", "11"), arrayOf<String>("h:mm AM/PM", "12"), arrayOf<String>("h:mm:ss AM/PM", "13"), arrayOf<String>("h:mm", "14"), arrayOf<String>("h:mm:ss", "15"), arrayOf<String>("yyyy/m/d h:mm", "16"),

// 17h through 24h are undocumented international formats!!!
arrayOf<String>("($#,##0_);($#,##0)", "17"), arrayOf<String>("($#,##0_);[Red]($#,##0)", "18"), arrayOf<String>("($#,##0_);[Red]($#,##0)", "19"), arrayOf<String>("($#,##0.00_);[Red]($#,##0.00)", "1a"), arrayOf<String>("[$-411]ge.m.d", "1b"), arrayOf<String>("[$-411]ggge\u5E74m\u6708d\u65E5", "1c"), arrayOf<String>("reserved", "1d"), arrayOf<String>("m/d/yy", "1e"), arrayOf<String>("yyyy\u5E74m\u6708d\u65E5", "1f"), arrayOf<String>("h\u6642mm\u5206", "20"), arrayOf<String>("h\u6642mm\u5206ss\u79D2", "21"), arrayOf<String>("yyyy\u5E74m\u6708", "22"), arrayOf<String>("m\u6708d\u65E5", "23"), arrayOf<String>("reserved", "24"),

// more currency
arrayOf<String>("#,##0;-#,##0", "25"), arrayOf<String>("#,##0;[Red]-#,##0", "26"), arrayOf<String>("#,##0.00;-#,##0.00", "27"), arrayOf<String>("#,##0.00;[Red]-#,##0.00", "28"), arrayOf<String>("_ * #,##0_ ;_ * -#,##0_ ;_ * -_ ;_ @_ ", "29"), arrayOf<String>("_ \u00a5* #,##0_ ;_ \u00a5* -#,##0_ ;_ \u00a5* -_ ;_ @_ ", "2a"), arrayOf<String>("_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * -_ ;_ @_ ", "2b"), arrayOf<String>("_ \u00a5* #,##0.00_ ;_ \u00a5* -#,##0.00_ ;_ \u00a5* -_ ;_ @_ ", "2c"),

// misc
arrayOf<String>("mm:ss", "2d"), arrayOf<String>("[h]:mm:ss", "2e"), arrayOf<String>("mm:ss.0", "2f"), arrayOf<String>("##0.0E+0", "30"), arrayOf<String>("@", "31"), arrayOf<String>("yyyy\u5E74m\u6708", "37"), arrayOf<String>("m\u6708d\u65E5", "38"), arrayOf<String>("[$-411]ge.m.d", "39"), arrayOf<String>("[$-411]ggge\u5E74m\u6708d\u65E5", "3a"))

/* The NumberFormat element contains one of the following string constants
 * specifying the date or time format:
 * General Date, Short Date, Medium Date, Long Date, Short Time, Medium Time, or Long Time;
 * one of the following string constants specifying the numeric format:
 * Currency, Fixed, General, General Number, Percent, Scientific, or Standard (#,##0.00);
*/
// TODO: Add
// TODO: These may not be correct mappings
/* one of the following string constants specifying the Boolean values displayed: On/Off, True/False, or Yes/No;*/
val BUILTIN_FORMATS_MHTML = arrayOf<Array<String>>(arrayOf<String>("General Date", "M/D/YY"), arrayOf<String>("Short Date", "D-MMM"), arrayOf<String>("Medium Date", "D-MMM-YY"), arrayOf<String>("Long Date", "MMM-YY"), arrayOf<String>("Short Time", "mm:ss"), arrayOf<String>("Medium Time", "h:mm AM/PM"), arrayOf<String>("Long Time", "h:mm:ss AM/PM"), arrayOf<String>("Currency", "\"$\"#,##0_);(\"$\"#,##0)"), arrayOf<String>("Fixed", "0.00"), arrayOf<String>("General Number", "0"), arrayOf<String>("Percent", "0%"), arrayOf<String>("Scientific", "0.00E+00"), arrayOf<String>("Standard", "(#,##0.00);"))

/**
 * - Now this array includes the java format pattern.  If multiple patterns exist for negative
 * numbers, that is after a semicolon.  The [Red] needs to be stripped as well and applied outside
 * the Java formatting if desired.
*/
val NUMERIC_FORMATS = arrayOf<Array<String>>(arrayOf<String>("General", "0", ""), arrayOf<String>("0", "1", "0"), arrayOf<String>("0.00", "2", "0.00"),

// add below 2 to match excel
arrayOf<String>("#,##0", "3", "#,##0"), arrayOf<String>("#,##0.00", "4", "#,##0.00"),


arrayOf<String>("0%", "9", "0%"), // percent
arrayOf<String>("0.00%", "a", "#.00%"), // percent
arrayOf<String>("0.00E+00", "b", "0.00E00"), // scientific
arrayOf<String>("(#,##0.00_);(#,##0.00)", "27", "#,##0.00;(#,##0.00)"), //?
arrayOf<String>("_(*#,##0_);_(*(#,##0);_(*\"-\"_);_(@_)", "29", "#,##0;(#,##0)"), //?
arrayOf<String>("##00E+0", "30", "#.##E00"), //TODO: unable to write a good generic conversion string here, returning generic exponent
arrayOf<String>("@", "31", "@"), //?
// User-defined (Format ID > 164, but variable.  use Format ID= -1)
arrayOf<String>("0.0", "FF", "0.0"), arrayOf<String>("0.00;[Red]0.00", "FF", "#,##0.00;[Red]#,##0.00"), arrayOf<String>("0.00_);(0.00)", "FF", "#,##0.00###;(#,##0.00)"), arrayOf<String>("0.00_);[Red](0.00)", "FF", "#,##0.00###;[Red](#,##0.00)"))

// Decoded format Patterns, HEX ids, and java format strings for all date formats
/**
 * These formats are used in isDate handling and translating those into java format strings when possible
 * for toFormattedString behavior
*/
val DATE_FORMATS = arrayOf<Array<String>>(arrayOf<String>("mm-dd-yy", "E", "MM-dd-yy"), arrayOf<String>("d-mmm-yy", "F", "d-MMM-yy"), arrayOf<String>("d-mmm", "10", "d-MMM"), arrayOf<String>("mmm-yy", "11", "MMM-yy"), arrayOf<String>("h:mm AM/PM", "12", "h:mm aa"), arrayOf<String>("h:mm:ss AM/PM", "13", "h:mm:ss aa"), arrayOf<String>("h:mm", "14", "H:mm"), arrayOf<String>("h:mm:ss", "15", "H:mm:ss"), arrayOf<String>("m/d/yy h:mm", "16", "M/d/yy H:mm"), arrayOf<String>("mm:ss", "2D", "mm:ss"), arrayOf<String>("[h]:mm:ss", "2E"), arrayOf<String>("mm:ss.0", "2F", "mm:ss.S"),
// user-defined (Format ID > 164, but variable) - use Format ID of -1
arrayOf<String>("m/d/yyyy h:mm", "FF", "M/d/yyyy H:mm"), arrayOf<String>("m/d/yyyy;@", "FF", "M/d/yyyy"), arrayOf<String>("m/d/yy h:mm", "FF", "M/d/yy H:mm"), arrayOf<String>("hh:mm AM/PM", "FF", "hh:mm aa"), arrayOf<String>("m/d/yyyy", "FF", "M/d/yyyy"), arrayOf<String>("yyyy-mm-dd", "FF", "yyyy-MM-dd"),
// i know, weird formatting offset, but the xf says m/d/y and excel shows m/d/yyyy
arrayOf<String>("m/d/y", "FF", "M/d/yyyy"), arrayOf<String>("m/d/yy", "FF", "M/d/yy"), arrayOf<String>("m/d;@", "FF", "M/d"), arrayOf<String>("m/d/yy;@", "FF", "M/d/yy"), arrayOf<String>("mm/dd/yy;@", "FF", "MM/dd/yy"), arrayOf<String>("mm/dd/yy", "FF", "MM/dd/yy"), arrayOf<String>("mm/dd/yyyy", "FF", "MM/dd/yyyy"), arrayOf<String>("dd/mm/yy", "FF", "dd/MM/yy"), arrayOf<String>("yy/mm/dd;@", "FF", "yy/MM/dd"), arrayOf<String>("m/d/yy h:mm;@", "FF", "M/d/yy H:mm"), arrayOf<String>("hh:mm:ss.000", "FF", "HH:mm:ss.SSS"), arrayOf<String>("mmmm dd, yyyy", "FF", "MMMMM dd, yyyy"), arrayOf<String>("mmm dd, yyyy", "FF", "MMM dd, yyyy"), arrayOf<String>("m/d;@", "FF", "M/d"), arrayOf<String>("[$-409]d-mmm;@", "FF", "d-MMM"), arrayOf<String>("[$-409]d-mmm-yy;@", "FF", "d-MMM-yy"), arrayOf<String>("[$-409]d-mmm-yyyy;@", "FF", "d-MMM-yyyy"), arrayOf<String>("[$-409]dd-mmm-yy;@", "FF", "dd-Myyyy-mm-ddMM-yy"), arrayOf<String>("[$-409]dddd, mmmm dd, yyyy", "FF", "EEEE, MMMM dd, yyyy"), arrayOf<String>("[$-409]mmm-yy;@", "FF", "MMM-ss"), arrayOf<String>("[$-409]mmmm-yy;@", "FF", "MMMM-yy"), arrayOf<String>("[$-409]mmmm d, yyyy;@", "FF", "MMMM d, yyyy"), arrayOf<String>("[$-409]mmmmm;@", "FF", "MMMMM"), arrayOf<String>("[$-409]mmmmm-yy;@", "FF", "MMMMM-yy"), arrayOf<String>("[$-409]m/d/yy h:mm AM/PM;@", "FF", "M/d/yy H:mm aa"), arrayOf<String>("[$-409]mmmm d, yyyy;@", "FF", "MMMM d, yyyy"), arrayOf<String>("[$-409]h:mm:ss AM/PM", "FF", "H:mm:ss aa"), arrayOf<String>("[$-409]d-mmm;@", "FF", "d-MMM"), arrayOf<String>("[$-F800]dddd,mmmm dd,yyyy", "FF", "EEEE,MMMM dd, yyyy"), arrayOf<String>("[$-F800]dddd, mmmm dd, yyyy", "FF", "EEEE, MMMM dd, yyyy"), arrayOf<String>("dd/mm/yy hh:mm", "FF", "dd/MM/yy HH:mm"), arrayOf<String>("d-mmm-yyyy hh:mm", "FF", "d-MMM-yyyy HH:mm"), arrayOf<String>("yyyy-mm-dd hh:mm", "FF", "yyyy-MM-dd HH:mm"), arrayOf<String>("yyyy-mm-ss hh:mm:ss.000", "FF", "yyyy-MM-dd HH:mm:ss.SSS"), arrayOf<String>("yyyy-mm-dd", "FF", "yyyy-MM-dd"), arrayOf<String>("d/mm/yy", "FF", "d/MM/yy"), arrayOf<String>("d/mm/yy;@", "FF", "d/MM/yy"), arrayOf<String>("dddd, mmmm dd, yyyy", "FF", "EEEEE, MMMMM d, yyyy"), arrayOf<String>("[hhh]:mm", "FF", "h:mm"))/* ?? */
// Decoded Format Patterns and hex Format Ids
val CURRENCY_FORMATS = arrayOf<Array<String>>(
// below 4 do not match excel!
//   	{"$#,##0;($#,##0)", "1", "$#,##0;($#,##0)"},
//		{"$#,##0;[Red]($#,##0)", "2", "$#,##0;[Red]($#,##0)"},
//		{"$#,##0.00;($#,##0.00)", "3", "$#,##0.00;($#,##0.00)"},
//		{"$#,##0.00;[Red]($#,##0.00)", "4", "$#,##0.00;[Red]($#,##0.00)"},
arrayOf<String>("#,##0_;($#,##0)", "5", "#,##0;($#,##0)"), arrayOf<String>("#,##0_;[Red]($#,##0)", "6", "#,##0;[Red]($#,##0)"), arrayOf<String>("#,##0.00_;($#,##0.00)", "7", "#,##0.00###;($#,##0.00)"), arrayOf<String>("#,##0.00_;[Red]($#,##0.00)", "8", "#,##0.00;[Red]($#,##0.00)"), arrayOf<String>("_$*#,##0_;_($*($#,##0);_($*\"-\"_);_(@_)", "2a", "$#,##0;($#,##0)"), arrayOf<String>("_*#,##0_;_(*($#,##0);_(*\"-\"??_);_(@_)", "2b", "#,##0;(#,##0)"), arrayOf<String>("_$*#,##0_;_($*($#,##0);_($*\"-\"??_);_(@_)", "2c", "$#,##0;($#,##0)"),
// User-defined (Format ID > 164, but variable.  use Format ID= -1)
arrayOf<String>("_$*#,##0_;_($*($#,##0);_($*\"-\"??_);_(@_)", "FF", "$#,##0;($#,##0)"), arrayOf<String>("#,##0.00 [$-1]_);[Red](#,##0.00 [$-1])", "FF", "#,##0.00 ;[Red](#,##0.00 )"), arrayOf<String>("_ * #,##0.00_) [$-1]_ ;_ * (#,##0.00) [$-1]_ ;_ * -??_) [$-1]_ ;_ @_ ", "FF", "#,##0.00 ;(#,##0.00) "), arrayOf<String>("_ * #,##0.00_) [$€-1]_ ;_ * (#,##0.00) [$€-1]_ ;_ * -??_) [$€-1]_ ;_ @_ ", "FF", "#,##0.00 €;(#,##0.00) €"), arrayOf<String>("#,##0.00 [$€-1]", "FF", "#,##0.00 €"), arrayOf<String>("#,##0.00 [$-1]", "FF", "#,##0.00 "), arrayOf<String>("$#,##0.00", "FF", "$#,##0.00"), arrayOf<String>("_($* #,##0_);_($* (#,##0);_($* -??_);_(@_)", "FF", "$ #,##0;($* (#,##0)"), arrayOf<String>("#,##0.00 [$-1];[Red]#,##0.00 [$-1]", "FF", "#,##0.00 ;[Red]#,##0.00 "), arrayOf<String>("[$-411]#,##0.00", "FF", "#,##0.00"), arrayOf<String>("[$-411]#,##0.00;[Red][$-411]#,##0.00", "FF", "#,##0.00;[Red]#,##0.00"), arrayOf<String>("[$-411]#,##0.00;-[$-411]#,##0.00", "FF", "#,##0.00"), arrayOf<String>("[$-411]#,##0.00;[Red]-[$-411]#,##0.00", "FF", "#,##0.00;[Red]-#,##0.00"), arrayOf<String>("[$-809]#,##0.00", "FF", "#,##0.00"), arrayOf<String>("[$-809]#,##0.00;[Red][$-809]#,##0.00", "FF", "#,##0.00;[Red]#,##0.00"), arrayOf<String>("[$-809]#,##0.00;-[$-809]#,##0.00", "FF", "#,##0.00;-#,##0.00"), arrayOf<String>("[$-809]#,##0.00;[Red]-[$-809]#,##0.00", "FF", "#,##0.00;[Red]-#,##0.00"), arrayOf<String>("#,##0.0 [$$-C0C];[Red]#,##0.0 [$$-C0C]", "FF"),
// user-defined
arrayOf<String>("$#,##0;($#,##0)", "FF", "$#,##0;($#,##0)"), arrayOf<String>("$#,##0.00;($#,##0.00)", "FF", "$#,##0.00;($#,##0.00)"), arrayOf<String>("$#,##0;[Red]($#,##0)", "FF", "$#,##0;[Red]($#,##0)"), arrayOf<String>("$#,##0.00;[Red]($#,##0.00)", "FF", "$#,##0.00;[Red]($#,##0.00)"), arrayOf<String>("$#,##0.00;[Red]$#,##0.00", "FF", "$#,##0.00;[Red]$#,##0.00"), arrayOf<String>("$#,##0", "FF", "$#,##0"))

// this maps Java formats to Excel formats
val patternmap = arrayOf<Array<String>>(
// currency
arrayOf<String>("#,##0.00", "#,##0.00"), arrayOf<String>("$#,##0;($#,##0)", "(#,##0_);($#,##0)"), arrayOf<String>("$#,##0;[Red]($#,##0)", "$###,###;($###,###)"), arrayOf<String>("$#,##0.00;($#,##0.00)", "$###,###.##;($###,###.##)"), arrayOf<String>("$#,##0.00;[Red]($#,##0.00)", "$###,###.##;($###,###.##)"), arrayOf<String>("0%", "0%"), arrayOf<String>("0.00%", "0.00%"), arrayOf<String>("0.00E+00", "0.00E+00"), arrayOf<String>("# ?/?", "# ?/?"), arrayOf<String>("# ??/??", ""), arrayOf<String>("#,##0_;($#,##0)", ""), arrayOf<String>("#,##0_;[Red]($#,##0)", ""), arrayOf<String>("#,##0.00_;($#,##0.00)", ""), arrayOf<String>("#,##0.00_;[Red]($#,##0.00)", ""), arrayOf<String>("_*#,##0_;_(*($#,##0);_(*-_);_(@_)", ""), arrayOf<String>("_$*#,##0_;_($*($#,##0);_($*-_);_(@_)", ""), arrayOf<String>("_*#,##0_;_(*($#,##0);_(*-??_);_(@_)", ""), arrayOf<String>("_$*#,##0_;_($*($#,##0);_($*-??_);_(@_)", ""),

// date time
arrayOf<String>("mm-mmm-yy", "MM-dd-YY"), arrayOf<String>("m/d/y", "M/d/y"), arrayOf<String>("d-mmm-yy", "d-MMM-yy"), arrayOf<String>("d-mmm", "d-MMM"), arrayOf<String>("mmm-yy", "MMM-yy"), arrayOf<String>("h:mm AM/PM", "hh:mm a"), arrayOf<String>("h:mm:ss AM/PM", "hh:mm:ss a"), arrayOf<String>("h:mm", "hh:mm"), arrayOf<String>("h:mm:ss", "hh:mm:ss"), arrayOf<String>("m/d/yy h:mm", "M/d/yy hh:mm"))


// look up user-friendly format string and get Excel-qualified or encoded format string
// USED???
val EXCEL_FORMAT_LOOKUP = arrayOf<Array<String>>(
// currency
arrayOf<String>("0", ""), arrayOf<String>("0.00", ""), arrayOf<String>("#,##0", ""), arrayOf<String>("#,##0.00", ""), arrayOf<String>("($#,##0);($#,##0)", "\"$\"#,##0);(\"$\"#,##0)"), arrayOf<String>("($#,##0);[Red]($#,##0)", "\"$\"#,##0);[Red](\"$\"#,##0)"), arrayOf<String>("($#,##0.00);($#,##0.00)", "\"$\"#,##0.00);(\"$\"#,##0.00)"), arrayOf<String>("($#,##0.00);[Red]($#,##0.00)", "\"$\"#,##0.00);[Red](\"$\"#,##0.00)"), arrayOf<String>("0%", ""), arrayOf<String>("0.00%", ""), arrayOf<String>("0.00E+00", ""), arrayOf<String>("# ?/?", ""), arrayOf<String>("# ??/??", ""),
// Are these correct?  Should they be qualified???
arrayOf<String>("(#,##0_);($#,##0)", ""), arrayOf<String>("(#,##0_);[Red]($#,##0)", ""), arrayOf<String>("(#,##0.00_);($#,##0.00)", ""), arrayOf<String>("(#,##0.00_);[Red]($#,##0.00)", ""),


arrayOf<String>("(#,##0_);(#,##0)", ""), arrayOf<String>("(#,##0_);[Red](#,##0)", ""), arrayOf<String>("(#,##0.00_);(#,##0.00)", ""), arrayOf<String>("(#,##0.00_);[Red](#,##0.00)", ""), arrayOf<String>("_(*#,##0_);_(*($#,##0);_(*-_);_(@_)", "_(*#,##0_);_(*(\"$\"#,##0);_(*\"-\"_);_(@_)"), arrayOf<String>("_($*#,##0_);_($*($#,##0);_($*-_);_(@_)", "_(\"$\"*#,##0_);_(\"$\"*($#,##0);_(\"$\"*\"-\"_);_(@_)"), arrayOf<String>("_(*#,##0_);_(*($#,##0);_(*-??_);_(@_)", "_(*#,##0_);_(*(\"$\"#,##0);_(*\"-\"??_);_(@_)"), arrayOf<String>("_($*#,##0_);_($*($#,##0);_($*-??_);_(@_)", "_(\"$\"*#,##0_);_(\"$\"*(\"$\"#,##0);_(\"$\"*\"-\"??_);_(@_)"),

// dates
arrayOf<String>("mm-mmm-yy", ""), arrayOf<String>("m/d/y", ""), arrayOf<String>("d-mmm-yy", "d\\-mmm\\-yy"), arrayOf<String>("[$-409]mmmm d, yyyy;@", "mmmm\\d\\,\\-yyyy"), arrayOf<String>("d-mmm", "d\\-mmm"), arrayOf<String>("mmm-yy", "mmm\\-yy"), arrayOf<String>("h:mm AM/PM", "h:mm\\ AM/PM"), arrayOf<String>("h:mm:ss AM/PM", ""), arrayOf<String>("h:mm", ""), arrayOf<String>("h:mm:ss", ""), arrayOf<String>("m/d/yy h:mm", ""))
// Alignment's
val ALIGN_DEFAULT = 0
val ALIGN_LEFT = 1
val ALIGN_CENTER = 2
val ALIGN_RIGHT = 3
val ALIGN_FILL = 4
val ALIGN_JUSTIFY = 5
val ALIGN_CENTER_ACROSS_SELECTION = 6
val ALIGN_VERTICAL_TOP = 0
val ALIGN_VERTICAL_CENTER = 1
val ALIGN_VERTICAL_BOTTOM = 2
val ALIGN_VERTICAL_JUSTIFY = 3

val VERTICAL_ALIGNMENTS = arrayOf<String>("top", "middle", "bottom", "text-bottom", // justified
"text-bottom")// distributed"

/**
 * These Horizontal alignments are used to map to HTML output, so some
 * values may not be as expected.  For instance, "Center across selection" is
 * not something you can do in HTML.  We will use "Center" for this.
 *
 *
 * Not entirely True: HTML can center across merged table cells -- this is basically
 * the purpse AFAIK. -jm
*/
val HORIZONTAL_ALIGNMENTS = arrayOf<String>("Default", "Left", "Center", "Right", "Center", "Center", "Center")
}
}