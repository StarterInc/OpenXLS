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
package io.starter.formats.XLS;

/**
 *
 */
public interface XLSConstants {
    // Stream types
    short WK_GLOBALS = 0x5;
    short VB_MODULE = 0x6;
    short WK_WORKSHEET = 0x10;
    short WK_CHART = 0x20;
    short WK_MACROSHEET = 0x40;
    short WK_FILE = 0x100;

    // Cell types
    int
            TYPE_BLANK = -1,
            TYPE_STRING = 0,
            TYPE_FP = 1,
            TYPE_INT = 2,
            TYPE_FORMULA = 3,
            TYPE_BOOLEAN = 4,
            TYPE_DOUBLE = 5;

    // Book Options and constants

    // CalculationOptions
    int CALCULATE_ALWAYS = 0;
    int CALCULATE_EXPLICIT = 1;
    int CALCULATE_AUTO = 2; // replacement for calc always
    String CALC_MODE_PROP = "io.starter.OpenXLS.calcmode";
    String REFTRACK_PROP = "io.starter.OpenXLS.trackreferences";
    String USETEMPFILE_PROP = "io.starter.formats.LEO.usetempfile";
    String VALIDATEWORKBOOK = "io.starter.formats.LEO.validateworkbook";
    // Debug options
    int DEBUG_LOW = 50;
    int DEBUG_MEDIUM = 60;
    int DEBUG_HIGH = 100;

    // String table handling
    int STRING_ENCODING_AUTO = 0;
    int STRING_ENCODING_UNICODE = 1;
    int STRING_ENCODING_COMPRESSED = 2;
    int ALLOWDUPES = 0;
    int SHAREDUPES = 1;
    String DEFAULTENCODING = "ISO-8859-1"; // "UTF-8";"UTF-8";
    String UNICODEENCODING = "UTF-16LE"; // "UnicodeLittleUnmarked";

    // XLSRecord Opcodes
    short EXCEL2K = 0x1C0;
    short GARBAGE = 0xFFFFFFFE;
    short TXO = 0x1B6;
    short MSODRAWINGGROUP = 0xEB;
    short MSODRAWING = 0xEC;
    short MSODRAWINGSELECTION = 0xED;
    short PHONETIC = 0xEF;
    short CONTINUE = 0x3C;
    short COLINFO = 0x7D;
    short SST = 0xFC;
    short DSF = 0x161;
    short EXTSST = 0xFF;
    short ENDEXTSST = 0xFE;
    short BOF = 0x809;
    short FILEPASS = 0x2F;
    short INDEX = 0x20B;
    short DBCELL = 0xD7;
    short BOUNDSHEET = 0x85;
    short COUNTRY = 0x8C;    // record just after bound sheet
    short BOOKBOOL = 0xDA;
    short CALCCOUNT = 0x0C;
    short CALCMODE = 0x0D;
    short PRECISION = 0x0E;
    short REFMODE = 0x0F;
    short DELTA = 0x10;
    short ITERATION = 0x11;
    short DATE1904 = 0x22;
    short BACKUP = 0x40;
    short PRINT_ROW_HEADERS = 0x2A;
    short PRINT_GRIDLINES = 0x2B;
    short HORIZONTAL_PAGE_BREAKS = 0x1B;
    short HLINK = 0x1B8;
    short VERTICAL_PAGE_BREAKS = 0x1A;
    short DEFAULTROWHEIGHT = 0x225;
    short FONT = 0x31;
    short HEADERREC = 0x14;
    short FOOTERREC = 0x15;
    short LEFT_MARGIN = 0x26;
    short RIGHT_MARGIN = 0x27;
    short TOP_MARGIN = 0x28;
    short BOTTOM_MARGIN = 0x29;
    short DCON = 0x50;
    short DEFCOLWIDTH = 0x55;
    short EXTERNCOUNT = 0x16;
    short EXTERNSHEET = 0x17;
    short EXTERNNAME = 0x23;
    short FORMAT = 0x41E;
    short XF = 0xE0;
    short NAME = 0x18;
    short DIMENSIONS = 0x200;
    short FILE_LOCK = 0x195;
    short RRD_INFO = 0x196;
    short RRD_HEAD = 0x138;
    short EOF = 0x0A;
    short BLANK = 0x201;
    short MERGEDCELLS = 0xE5;
    short MULBLANK = 0xBE;
    short MULRK = 0xBD;
    short NOTE = 0x1C;
    short NUMBER = 0x203;
    short LABEL = 0x204;
    short LABELSST = 0xFD;
    short BOOLERR = 0x205;
    short FORMULA = 0x06; // 0x406;
    short ARRAY = 0x221;//0x21; //
    short SELECTION = 0x1D;
    short STYLE = 0x293;
    short ROW = 0x208;
    short RK = 0x27E; // this is wrong according to the documentation (0x27) ... -jm
    short RSTRING = 0xD6;
    short SHRFMLA = 0x4BC; // according to docs this is 0xBC
    short STRINGREC = 0x207;
    short TABLE = 0x236;
    short PANE = 0x41;
    short PASSWORD = 0x13;
    short INTERFACE_HDR = 0xE1;
    short USR_EXCL = 0x194;
    short PALETTE = 0x92;
    short PROTECT = 0x12;
    short OBJPROTECT = 0x63;
    short SCENPROTECT = 0xDD;
    short FEATHEADR = 0x867;    // extra protection settings + smarttag settings
    short SCL = 0xA0; // zoom
    short SHEETPROTECTION = 0x867;    //
    short SHEETLAYOUT = 0x862;
    short RANGEPROTECTION = 0x868;
    short PROT4REV = 0x1AF;
    short WINDOW_PROTECT = 0x19;
    short WINDOW1 = 0x3D;
    short WINDOW2 = 0x23E;
    short PLV = 0x88B;
    short RTENTEXU = 0x1B;
    short DV = 0x1BE;
    short DVAL = 0x1B2;
    short RTMERGECELLS = 0xE5;
    short SUPBOOK = 0x1AE;
    short USERSVIEWBEGIN = 0x1AA;
    short USERSVIEWEND = 0x1AB;
    short USERBVIEW = 0x1A9;
    short PLS = 0x4D;
    short WSBOOL = 0x81;
    short OBJ = 0x5D;
    short OBPROJ = 0xD3;
    short XLS_MAX_COLS = 0x100;
    short TABID = 0x13d;
    short GUTS = 0x80;
    short CODENAME = 0x1BA;
    short XCT = 0x59;    // 20080122 KSC:
    short CRN = 0x5A;    // ""

    // Pivot Table records
    short SXVIEW = 0xB0;
    short TABLESTYLES = 0x88E;
    short SXSTREAMID = 0xD5;
    short SXVS = 0xE3;
    short SXADDL = 0x864;
    short SXVDEX = 0x100;
    short SXPI = 0xB6;
    short SXDI = 0xC5;
    short SXDB = 0xC6;
    short SXFDB = 0xC7;
    short SXEX = 0xF1;
    short QSISXTAG = 0x802;
    short SXVIEWEX9 = 0x810;
    short DCONREF = 0x51;
    short DCONNAME = 0x52;
    short DCONBIN = 0x1B5;
    short SXFORMAT = 0xFB;
    short SXLI = 0xB5;
    short SXVI = 0xB2;
    short SXVD = 0xB1;
    short SXIVD = 0xB4;
    short SXDBEX = 0x122;
    short SXFDBTYPE = 0x1BB;
    short SXDBB = 0xC8;
    short SXNUM = 0xC9;
    short SXBOOL = 0xCA;
    short SXSTRING = 0xCD;

    // Printing records
    short SETUP = 0xA1;
    short HCENTER = 0x83;
    short VCENTER = 0x84;
    short LEFTMARGIN = 0x26;
    short RIGHTMARGIN = 0x27;
    short TOPMARGIN = 0x28;
    short BOTTOMMARGIN = 0x29;
    short PRINTGRID = 0x2B;
    short PRINTROWCOL = 0x2A;

    // Conditional Formatting
    short CF = 0x1B1;
    short CONDFMT = 0x1B0;
    // 2007 Conditional Formatting
    short CF12 = 0x87A;
    short CONDFMT12 = 0x879;

    // AutoFilter
    short AUTOFILTER = 0x9E;

    // Chart items
    short UNITS = 0x1001;
    short CHART = 0x1002;
    short SERIES = 0x1003;
    short DATAFORMAT = 0x1006;
    short LINEFORMAT = 0x1007;
    short MARKERFORMAT = 0x1009;
    short AREAFORMAT = 0x100A;
    short PIEFORMAT = 0x100B;
    short ATTACHEDLABEL = 0x100C;
    short SERIESTEXT = 0x100D;
    short CHARTFORMAT = 0x1014;
    short LEGEND = 0x1015;
    short SERIESLIST = 0x1016;
    short BAR = 0x1017;
    short LINE = 0x1018;
    short PIE = 0x1019;
    short AREA = 0x101A;
    short SCATTER = 0x101B;
    short CHARTLINE = 0x101C;
    short AXIS = 0x101D;
    short TICK = 0x101E;
    short VALUERANGE = 0x101F;
    short CATSERRANGE = 0x1020;
    short AXISLINEFORMAT = 0x1021;
    short CHARTFORMATLINK = 0x1022;
    short DEFAULTTEXT = 0x1024;
    short TEXTDISP = 0x1025;
    short FONTX = 0x1026;
    short OBJECTLINK = 0x1027;
    short FRAME = 0x1032;
    short BEGIN = 0x1033;
    short END = 0x1034;
    short PLOTAREA = 0x1035;
    short THREED = 0x103A;
    short PICF = 0x103C;
    short DROPBAR = 0x103D;
    short RADAR = 0x103E;
    short SURFACE = 0x103F;
    short RADARAREA = 0x1040;
    short AXISPARENT = 0x1041;
    short LEGENDXN = 0x1043;
    short SHTPROPS = 0x1044;
    short SERTOCRT = 0x1045;
    short AXESUSED = 0x1046;
    short SBASEREF = 0x1048;
    short SERPARENT = 0x104A;
    short SERAUXTREND = 0x104B;
    short IFMT = 0x104E;
    short POS = 0x104F;
    short ALRUNS = 0x1450;
    short AI = 0x1051;
    short SERAUXERRBAR = 0x105B;
    short SERFMT = 0x105D;
    short CHART3DBARSHAPE = 0x105F;
    short FBI = 0x1460;
    short BOPPOP = 0x1061;
    short AXCENT = 0x1062;
    short DAT = 0x1063;
    short PLOTGROWTH = 0x1064;
    short SIIINDEX = 0x1065;
    short GELFRAME = 0x1066;
    short BOPPOPCUSTOM = 0x1067;
    //20080703 KSC: Excel 9 Chart Records
    short CRTLAYOUT12 = 0x89D;
    short CRTLAYOUT12A = 0x08A7;
    short CHARTFRTINFO = 0x0850;
    short FRTWRAPPER = 0x0851;
    short STARTBLOCK = 0x0852;
    short ENDBLOCK = 0x0853;
    short STARTOBJECT = 0x0854;
    short ENDOBJECT = 0x0855;
    short CATLAB = 0x0856;
    short YMULT = 0x0857;
    short SXVIEWLINK = 0x0858;
    short PIVOTCHARTBITS = 0x0859;
    short FRTFONTLIST = 0x085A;
    short PIVOTCHARTLINK = 0x0861;
    short DATALABEXT = 0x086A;
    short DATALABEXTCONTENTS = 0x086B;
    short FONTBASIS = 0x1060;

    // max size of records -- depends on XLS version
    int MAXRECLEN = 8224;
    int MAXROWS_BIFF8 = 65536;
    int MAXCOLS_BIFF8 = 256;
    int MAXROWS = 1048576; // new Excel 2007 limits
    int MAXCOLS = 16384;    // new Excel 2007 limits

}