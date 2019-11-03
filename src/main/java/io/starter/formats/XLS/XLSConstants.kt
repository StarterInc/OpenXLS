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

/**
 *
 */
interface XLSConstants {
    companion object {
        // Stream types
        val WK_GLOBALS: Short = 0x5
        val VB_MODULE: Short = 0x6
        val WK_WORKSHEET: Short = 0x10
        val WK_CHART: Short = 0x20
        val WK_MACROSHEET: Short = 0x40
        val WK_FILE: Short = 0x100

        // Cell types
        val TYPE_BLANK = -1
        val TYPE_STRING = 0
        val TYPE_FP = 1
        val TYPE_INT = 2
        val TYPE_FORMULA = 3
        val TYPE_BOOLEAN = 4
        val TYPE_DOUBLE = 5

        // Book Options and constants

        // CalculationOptions
        val CALCULATE_ALWAYS = 0
        val CALCULATE_EXPLICIT = 1
        val CALCULATE_AUTO = 2 // replacement for calc always
        val CALC_MODE_PROP = "io.starter.OpenXLS.calcmode"
        val REFTRACK_PROP = "io.starter.OpenXLS.trackreferences"
        val USETEMPFILE_PROP = "io.starter.formats.LEO.usetempfile"
        val VALIDATEWORKBOOK = "io.starter.formats.LEO.validateworkbook"
        // Debug options
        val DEBUG_LOW = 50
        val DEBUG_MEDIUM = 60
        val DEBUG_HIGH = 100

        // String table handling
        val STRING_ENCODING_AUTO = 0
        val STRING_ENCODING_UNICODE = 1
        val STRING_ENCODING_COMPRESSED = 2
        val ALLOWDUPES = 0
        val SHAREDUPES = 1
        val DEFAULTENCODING = "ISO-8859-1" // "UTF-8";"UTF-8";
        val UNICODEENCODING = "UTF-16LE" // "UnicodeLittleUnmarked";

        // XLSRecord Opcodes
        val EXCEL2K: Short = 0x1C0
        val GARBAGE: Short = -0x2
        val TXO: Short = 0x1B6
        val MSODRAWINGGROUP: Short = 0xEB
        val MSODRAWING: Short = 0xEC
        val MSODRAWINGSELECTION: Short = 0xED
        val PHONETIC: Short = 0xEF
        val CONTINUE: Short = 0x3C
        val COLINFO: Short = 0x7D
        val SST: Short = 0xFC
        val DSF: Short = 0x161
        val EXTSST: Short = 0xFF
        val ENDEXTSST: Short = 0xFE
        val BOF: Short = 0x809
        val FILEPASS: Short = 0x2F
        val INDEX: Short = 0x20B
        val DBCELL: Short = 0xD7
        val BOUNDSHEET: Short = 0x85
        val COUNTRY: Short = 0x8C    // record just after bound sheet
        val BOOKBOOL: Short = 0xDA
        val CALCCOUNT: Short = 0x0C
        val CALCMODE: Short = 0x0D
        val PRECISION: Short = 0x0E
        val REFMODE: Short = 0x0F
        val DELTA: Short = 0x10
        val ITERATION: Short = 0x11
        val DATE1904: Short = 0x22
        val BACKUP: Short = 0x40
        val PRINT_ROW_HEADERS: Short = 0x2A
        val PRINT_GRIDLINES: Short = 0x2B
        val HORIZONTAL_PAGE_BREAKS: Short = 0x1B
        val HLINK: Short = 0x1B8
        val VERTICAL_PAGE_BREAKS: Short = 0x1A
        val DEFAULTROWHEIGHT: Short = 0x225
        val FONT: Short = 0x31
        val HEADERREC: Short = 0x14
        val FOOTERREC: Short = 0x15
        val LEFT_MARGIN: Short = 0x26
        val RIGHT_MARGIN: Short = 0x27
        val TOP_MARGIN: Short = 0x28
        val BOTTOM_MARGIN: Short = 0x29
        val DCON: Short = 0x50
        val DEFCOLWIDTH: Short = 0x55
        val EXTERNCOUNT: Short = 0x16
        val EXTERNSHEET: Short = 0x17
        val EXTERNNAME: Short = 0x23
        val FORMAT: Short = 0x41E
        val XF: Short = 0xE0
        val NAME: Short = 0x18
        val DIMENSIONS: Short = 0x200
        val FILE_LOCK: Short = 0x195
        val RRD_INFO: Short = 0x196
        val RRD_HEAD: Short = 0x138
        val EOF: Short = 0x0A
        val BLANK: Short = 0x201
        val MERGEDCELLS: Short = 0xE5
        val MULBLANK: Short = 0xBE
        val MULRK: Short = 0xBD
        val NOTE: Short = 0x1C
        val NUMBER: Short = 0x203
        val LABEL: Short = 0x204
        val LABELSST: Short = 0xFD
        val BOOLERR: Short = 0x205
        val FORMULA: Short = 0x06 // 0x406;
        val ARRAY: Short = 0x221//0x21; //
        val SELECTION: Short = 0x1D
        val STYLE: Short = 0x293
        val ROW: Short = 0x208
        val RK: Short = 0x27E // this is wrong according to the documentation (0x27) ... -jm
        val RSTRING: Short = 0xD6
        val SHRFMLA: Short = 0x4BC // according to docs this is 0xBC
        val STRINGREC: Short = 0x207
        val TABLE: Short = 0x236
        val PANE: Short = 0x41
        val PASSWORD: Short = 0x13
        val INTERFACE_HDR: Short = 0xE1
        val USR_EXCL: Short = 0x194
        val PALETTE: Short = 0x92
        val PROTECT: Short = 0x12
        val OBJPROTECT: Short = 0x63
        val SCENPROTECT: Short = 0xDD
        val FEATHEADR: Short = 0x867    // extra protection settings + smarttag settings
        val SCL: Short = 0xA0 // zoom
        val SHEETPROTECTION: Short = 0x867    //
        val SHEETLAYOUT: Short = 0x862
        val RANGEPROTECTION: Short = 0x868
        val PROT4REV: Short = 0x1AF
        val WINDOW_PROTECT: Short = 0x19
        val WINDOW1: Short = 0x3D
        val WINDOW2: Short = 0x23E
        val PLV: Short = 0x88B
        val RTENTEXU: Short = 0x1B
        val DV: Short = 0x1BE
        val DVAL: Short = 0x1B2
        val RTMERGECELLS: Short = 0xE5
        val SUPBOOK: Short = 0x1AE
        val USERSVIEWBEGIN: Short = 0x1AA
        val USERSVIEWEND: Short = 0x1AB
        val USERBVIEW: Short = 0x1A9
        val PLS: Short = 0x4D
        val WSBOOL: Short = 0x81
        val OBJ: Short = 0x5D
        val OBPROJ: Short = 0xD3
        val XLS_MAX_COLS: Short = 0x100
        val TABID: Short = 0x13d
        val GUTS: Short = 0x80
        val CODENAME: Short = 0x1BA
        val XCT: Short = 0x59    // 20080122 KSC:
        val CRN: Short = 0x5A    // ""

        // Pivot Table records
        val SXVIEW: Short = 0xB0
        val TABLESTYLES: Short = 0x88E
        val SXSTREAMID: Short = 0xD5
        val SXVS: Short = 0xE3
        val SXADDL: Short = 0x864
        val SXVDEX: Short = 0x100
        val SXPI: Short = 0xB6
        val SXDI: Short = 0xC5
        val SXDB: Short = 0xC6
        val SXFDB: Short = 0xC7
        val SXEX: Short = 0xF1
        val QSISXTAG: Short = 0x802
        val SXVIEWEX9: Short = 0x810
        val DCONREF: Short = 0x51
        val DCONNAME: Short = 0x52
        val DCONBIN: Short = 0x1B5
        val SXFORMAT: Short = 0xFB
        val SXLI: Short = 0xB5
        val SXVI: Short = 0xB2
        val SXVD: Short = 0xB1
        val SXIVD: Short = 0xB4
        val SXDBEX: Short = 0x122
        val SXFDBTYPE: Short = 0x1BB
        val SXDBB: Short = 0xC8
        val SXNUM: Short = 0xC9
        val SXBOOL: Short = 0xCA
        val SXSTRING: Short = 0xCD

        // Printing records
        val SETUP: Short = 0xA1
        val HCENTER: Short = 0x83
        val VCENTER: Short = 0x84
        val LEFTMARGIN: Short = 0x26
        val RIGHTMARGIN: Short = 0x27
        val TOPMARGIN: Short = 0x28
        val BOTTOMMARGIN: Short = 0x29
        val PRINTGRID: Short = 0x2B
        val PRINTROWCOL: Short = 0x2A

        // Conditional Formatting
        val CF: Short = 0x1B1
        val CONDFMT: Short = 0x1B0
        // 2007 Conditional Formatting
        val CF12: Short = 0x87A
        val CONDFMT12: Short = 0x879

        // AutoFilter
        val AUTOFILTER: Short = 0x9E

        // Chart items
        val UNITS: Short = 0x1001
        val CHART: Short = 0x1002
        val SERIES: Short = 0x1003
        val DATAFORMAT: Short = 0x1006
        val LINEFORMAT: Short = 0x1007
        val MARKERFORMAT: Short = 0x1009
        val AREAFORMAT: Short = 0x100A
        val PIEFORMAT: Short = 0x100B
        val ATTACHEDLABEL: Short = 0x100C
        val SERIESTEXT: Short = 0x100D
        val CHARTFORMAT: Short = 0x1014
        val LEGEND: Short = 0x1015
        val SERIESLIST: Short = 0x1016
        val BAR: Short = 0x1017
        val LINE: Short = 0x1018
        val PIE: Short = 0x1019
        val AREA: Short = 0x101A
        val SCATTER: Short = 0x101B
        val CHARTLINE: Short = 0x101C
        val AXIS: Short = 0x101D
        val TICK: Short = 0x101E
        val VALUERANGE: Short = 0x101F
        val CATSERRANGE: Short = 0x1020
        val AXISLINEFORMAT: Short = 0x1021
        val CHARTFORMATLINK: Short = 0x1022
        val DEFAULTTEXT: Short = 0x1024
        val TEXTDISP: Short = 0x1025
        val FONTX: Short = 0x1026
        val OBJECTLINK: Short = 0x1027
        val FRAME: Short = 0x1032
        val BEGIN: Short = 0x1033
        val END: Short = 0x1034
        val PLOTAREA: Short = 0x1035
        val THREED: Short = 0x103A
        val PICF: Short = 0x103C
        val DROPBAR: Short = 0x103D
        val RADAR: Short = 0x103E
        val SURFACE: Short = 0x103F
        val RADARAREA: Short = 0x1040
        val AXISPARENT: Short = 0x1041
        val LEGENDXN: Short = 0x1043
        val SHTPROPS: Short = 0x1044
        val SERTOCRT: Short = 0x1045
        val AXESUSED: Short = 0x1046
        val SBASEREF: Short = 0x1048
        val SERPARENT: Short = 0x104A
        val SERAUXTREND: Short = 0x104B
        val IFMT: Short = 0x104E
        val POS: Short = 0x104F
        val ALRUNS: Short = 0x1450
        val AI: Short = 0x1051
        val SERAUXERRBAR: Short = 0x105B
        val SERFMT: Short = 0x105D
        val CHART3DBARSHAPE: Short = 0x105F
        val FBI: Short = 0x1460
        val BOPPOP: Short = 0x1061
        val AXCENT: Short = 0x1062
        val DAT: Short = 0x1063
        val PLOTGROWTH: Short = 0x1064
        val SIIINDEX: Short = 0x1065
        val GELFRAME: Short = 0x1066
        val BOPPOPCUSTOM: Short = 0x1067
        //20080703 KSC: Excel 9 Chart Records
        val CRTLAYOUT12: Short = 0x89D
        val CRTLAYOUT12A: Short = 0x08A7
        val CHARTFRTINFO: Short = 0x0850
        val FRTWRAPPER: Short = 0x0851
        val STARTBLOCK: Short = 0x0852
        val ENDBLOCK: Short = 0x0853
        val STARTOBJECT: Short = 0x0854
        val ENDOBJECT: Short = 0x0855
        val CATLAB: Short = 0x0856
        val YMULT: Short = 0x0857
        val SXVIEWLINK: Short = 0x0858
        val PIVOTCHARTBITS: Short = 0x0859
        val FRTFONTLIST: Short = 0x085A
        val PIVOTCHARTLINK: Short = 0x0861
        val DATALABEXT: Short = 0x086A
        val DATALABEXTCONTENTS: Short = 0x086B
        val FONTBASIS: Short = 0x1060

        // max size of records -- depends on XLS version
        val MAXRECLEN = 8224
        val MAXROWS_BIFF8 = 65536
        val MAXCOLS_BIFF8 = 256
        val MAXROWS = 1048576 // new Excel 2007 limits
        val MAXCOLS = 16384    // new Excel 2007 limits
    }

}