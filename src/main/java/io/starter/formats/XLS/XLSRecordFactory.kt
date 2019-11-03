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

import io.starter.formats.XLS.charts.*
import io.starter.formats.XLS.formulas.Ptg

/**
 * Factory to create XLSRecords and Ptgs.
 */
class XLSRecordFactory
/**
 * This class is static only, prohibit construction.
 */
private constructor() : XLSConstants {

    init {
        throw UnsupportedOperationException(
                "XLSRecordFactory is purely static")
    }

    companion object {

        // 20060504 KSC: add separate array for unary prefix operators:
        //				 maps regular operator to unary pseudo version located in ptgLookup below
        var ptgPrefixOperators = arrayOf<Array<Any>>(arrayOf<Any>("+", "u+"), arrayOf<Any>("-", "u-"))// should $ be in here???

        /**
         * Maps BIFF8 record opcodes to classes.
         */
        // Removed static object instance creations as very difficult (impossible) to dereference so as can release memeory ...
        // TRY :	private static final Map records;

        // subset of ptgLookup below used for pattern matching in formula strings
        var ptgOps = arrayOf<Array<String>>(arrayOf("#VALUE!", "PtgErr"), arrayOf("#NULL!", "PtgErr"), arrayOf("#DIV/0!", "PtgErr"), arrayOf("#VALUE!", "PtgErr"), arrayOf("#REF!", "PtgErr"), arrayOf("#NUM!", "PtgErr"), arrayOf("#N/A", "PtgErr"),
                // operators
                arrayOf("(", "PtgParen"), arrayOf("*", "PtgMlt"), arrayOf("/", "PtgDiv"), arrayOf("^", "PtgPower"), arrayOf("&", "PtgConcat"), arrayOf("<>", "PtgNE"), arrayOf("<=", "PtgLE"), arrayOf("<", "PtgLT"), arrayOf(">=", "PtgGE"), arrayOf(">", "PtgGT"), arrayOf("=", "PtgEQ"), arrayOf("!=", "PtgNE"), arrayOf("+", "PtgAdd"), // moved to AFTER other operators
                arrayOf("-", "PtgSub"))
        // this is how you init a 2D array of Objects
        var ptgLookup = arrayOf<Array<String>>(
                // 20070215 KSC: added constant strings to avoid PtgRef3d matches
                arrayOf("#VALUE!", "PtgErr"), arrayOf("#NULL!", "PtgErr"), arrayOf("#DIV/0!", "PtgErr"), arrayOf("#VALUE!", "PtgErr"), arrayOf("#REF!", "PtgErr"), arrayOf("#NUM!", "PtgErr"), arrayOf("#N/A", "PtgErr"),
                // operators
                arrayOf("(", "PtgParen"), arrayOf("*", "PtgMlt"), arrayOf("/", "PtgDiv"), arrayOf("^", "PtgPower"), arrayOf("&", "PtgConcat"), arrayOf("<>", "PtgNE"), arrayOf("<=", "PtgLE"), arrayOf("<", "PtgLT"), arrayOf(">=", "PtgGE"), arrayOf(">", "PtgGT"), arrayOf("=", "PtgEQ"), arrayOf("!=", "PtgNE"),
                //{" ","PtgIsect"}, intersection operator, need to work out how to use a space as operator-- SEE PTGMEMFUNC/MEMAREA -- only valid when parsing complex ranges
                //{",","PtgUnion"},  problems matching as a separator in string parsing
                //{":","PtgRange"}, // may have issues with ptgArea?
                arrayOf("+", "PtgAdd"), // moved to AFTER other operators
                arrayOf("-", "PtgSub"),
                //operands
                arrayOf("PtgStr", "PtgStr"), arrayOf("PtgNumber", "PtgNumber"), arrayOf("PtgInt", "PtgInt"), arrayOf("PtgRef", "PtgRef"), arrayOf("PtgArea", "PtgArea"), arrayOf("false", "PtgBool"), arrayOf("true", "PtgBool"), arrayOf("PtgArea3d", "PtgArea3d"), arrayOf("PtgRef3d", "PtgRef3d"), arrayOf("PtgArray", "PtgArray"), arrayOf("PtgMissArg", "PtgMissArg"), arrayOf("PtgMemFunc", "PtgMemFunc"), arrayOf("PtgAtr", "PtgAtr"),
                //functions
                arrayOf("PtgFunc", "PtgFunc"), arrayOf("PtgFuncVar", "PtgFuncVar"),
                // unary prefix operators:  see FormulaParser.splitString
                arrayOf("u+", "PtgUPlus"), arrayOf("u-", "PtgUMinus"), arrayOf(")", "PtgParen"))

        /*  DO DIFFERENTLY SO TO AVOID HANGING OBJECT REFERENCES
        static {
            HashMap recmap = new HashMap();

            // Most Frequent
            case BLANK), new Blank() );
            case ROW), new Row() );
            case XF), new Xf() );
            case INDEX), new Index() );
            case COUNTRY), new Country() );
            case CALCMODE), new CalcMode() );
            case DIMENSIONS), new Dimensions() );
            case SELECTION), new Selection() );
            case DEFAULTROWHEIGHT),  new DefaultRowHeight() );
            case DEFCOLWIDTH), new DefColWidth() );
            case DBCELL), new Dbcell() );
            case BOF), new Bof() );
            case BOUNDSHEET), new Boundsheet() );
            case EOF), new Eof() );
            case FORMAT), new Format() );
            case STYLE), new Style() );
            case PASSWORD), new Password() );
            case PALETTE), new Palette() );
            case ARRAY), new Array() );
            case BOOLERR), new Boolerr() );
            case EXTERNSHEET), new Externsheet() );
            case EXTERNNAME), new Externname() );
            case FORMULA), new Formula() );
            case LABEL), new Label() );
            case TXO), new Txo() );
            case CONTINUE), new Continue() );
            case SST), new Sst() );
            case GUTS), new Guts() );
            case EXTSST), new Extsst() );
            case HLINK), new Hlink() );
            case LABELSST), new Labelsst() );
            case NUMBER), new NumberRec() );
            case MERGEDCELLS), new Mergedcells() );
            case MULBLANK), new Mulblank() );
            case MULRK), new Mulrk() );
            case RK), new Rk() );
            case RSTRING), new Rstring() );
            case SHRFMLA), new Shrfmla() );
            case STRINGREC), new StringRec() );
            case SUPBOOK), new Supbook() );
            case DV), new Dv() );
            case DVAL), new Dval() );
            case SETUP), new Setup() );
            case HCENTER), new HCenter() );
            case VCENTER), new VCenter() );
            case LEFTMARGIN), new LeftMargin() );
            case RIGHTMARGIN), new RightMargin() );
            case TOPMARGIN), new TopMargin() );
            case BOTTOMMARGIN), new BottomMargin() );
            case PRINTGRID), new PrintGrid() );
            case PRINTROWCOL), new PrintRowCol() );
            case XCT), new Xct() );
            case CRN), new Crn() );
            case NOTE), new Note() );

            // Named Ranges and References
            case NAME), new Name() );

            // Workbook Settings
            case FONT), new Font() );
            case DSF), new Dsf() );
            case WINDOW1), new Window1() );
            case WINDOW2), new Window2() );
            case CODENAME), new Codename() );
            case PROTECT), new Protect() );
            case OBJPROTECT), new ObjProtect() );
            case SCENPROTECT), new ScenProtect() );
            case FEATHEADR), new FeatHeadr() );
            case PROT4REV), new Prot4rev() );
            case COLINFO), new Colinfo() );
            case USERSVIEWBEGIN), new Usersviewbegin() );
            case USERSVIEWEND), new Usersviewend() );
            case WSBOOL), new WsBool() );
            case BOOKBOOL), new BookBool() );
            case USR_EXCL), new UsrExcl() );
            case INTERFACE_HDR), new InterfaceHdr() );
            case RRD_INFO), new RrdInfo() );
            case RRD_HEAD), new RrdHead() );
            case FILE_LOCK), new FileLock() );
            case PLS), new Pls() );
            case HEADERREC), new Headerrec() );
            case DATE1904), new NineteenOhFour() );

            // Sheet Settings
            case OBJ), new Obj() );
            case OBPROJ), new Obproj() );
            case FOOTERREC), new Footerrec() );
            case TABID), new TabID() );
            case PANE),  new Pane() );
            case SCL),  new Scl() );

            // Protection settings
            case FILEPASS), new Filepass() );

            // Conditional Formatting
            case CF), new Cf() );
            case CONDFMT), new Condfmt() );

            // Auto filter
            case AUTOFILTER),  new AutoFilter() );

            // Chart Records
            case CHART), new Chart() );
            case SERIES), new Series() );
            case SERIESTEXT), new SeriesText() );
            case SERIESLIST), new SeriesList() );
            case AI), new Ai() );
            case BEGIN), new Begin() );
            case END), new End() );
            case UNITS), new Units() );
            case CHART), new Chart() );
            case DATAFORMAT), new DataFormat() );
            case LINEFORMAT), new LineFormat() );
            case MARKERFORMAT), new MarkerFormat() );
            case AREAFORMAT), new AreaFormat() );
            case PIEFORMAT), new PieFormat() );
            case ATTACHEDLABEL), new AttachedLabel() );
            case CHARTFORMAT), new ChartFormat() );
            case LEGEND), new Legend() );
            case BAR), new Bar() );
            case LINE), new Line() );
            case PIE), new Pie() );
            case AREA), new Area() );
            case SCATTER), new Scatter() );
            case CHARTLINE), new ChartLine() );
            case AXIS), new Axis() );
            case TICK), new Tick() );
            case VALUERANGE), new ValueRange() );
            case CATSERRANGE), new CatserRange() );
            case AXISLINEFORMAT), new AxisLineFormat() );
            case CHARTFORMATLINK), new ChartFormatLink() );
            case DEFAULTTEXT), new DefaultText() );
            case TEXTDISP), new TextDisp() );
            case FONTX), new Fontx() );
            case OBJECTLINK), new ObjectLink() );
            case FRAME), new Frame() );
            case BEGIN), new Begin() );
            case END), new End() );
            case PLOTAREA), new PlotArea() );
            case THREED), new ThreeD() );
            case PICF), new Picf() );
            case DROPBAR), new Dropbar() );
            case RADAR), new Radar() );
            case SURFACE), new Surface() );
            case RADARAREA), new RadarArea() );
            case AXISPARENT), new AxisParent() );
            case LEGENDXN), new Legendxn() );
            case SHTPROPS), new ShtProps() );
            case SERTOCRT), new SerToCrt() );
            case AXESUSED), new Axesused() );
            case SBASEREF), new SbaseRef() );
            case SERPARENT), new SerParent() );
            case SERAUXTREND), new SerauxTrend() );
            case IFMT), new Ifmt() );
            case POS), new Pos() );
            case ALRUNS), new AlRuns() );
            case AI), new Ai() );
            case SERAUXERRBAR), new SerauxErrBar() );
            case SERFMT), new Serfmt() );
            case CHART3DBARSHAPE), new Chart3DBarShape() );
            case FBI), new Fbi() );
            case BOPPOP), new Boppop() );
            case AXCENT), new Axcent() );
            case DAT), new Dat() );
            case PLOTGROWTH), new PlotGrowth() );
            case SIIINDEX), new SiIndex() );
            case GELFRAME), new GelFrame() );
            case BOPPOPCUSTOM), new BoppopCustom() );
            case FONTBASIS),  new FontBasis() );

            // PivotTable Records
            case SXVIEW), new Sxview() );
            case SXFORMAT), new Sxformat() );
            case SXLI), new Sxli() );
            case SXVI), new Sxvi() );
            case SXVD), new Sxvd() );
            case SXIVD), new Sxivd() );

            // Object and Picture Records
            case PHONETIC), new Phonetic() );
            case MSODRAWING), new MSODrawing() );
            case MSODRAWINGGROUP), new MSODrawingGroup() );
            case MSODRAWINGSELECTION), new MSODrawingSelection() );

            // Excel 9 Chart Records
            case CHARTFRTINFO), new ChartFrtInfo() );
            case FRTWRAPPER), new FrtWrapper() );
            case STARTBLOCK), new StartBlock() );
            case ENDBLOCK), new EndBlock() );
            case STARTOBJECT), new StartObject() );
            case ENDOBJECT), new EndObject() );
            case CATLAB), new CatLab() );
            case YMULT), new YMult() );
            case SXVIEWLINK), new SxViewLink() );
            case PIVOTCHARTBITS), new PivotChartBits() );
            case FRTFONTLIST), new FrtFontList() );
            case PIVOTCHARTLINK), new PivotChartLink() );
            case DATALABEXTCONTENTS), new DataLabExtContents() );
            case DATALABEXT), new DataLabExt() );

            records = Collections.unmodifiableMap( recmap );
        }
    */
        /*
     * Create a ptg record from a name, will be init'ed elsewhere if needed.
     * I am keeping this seperate from getBiffRecord for performance reasons.  Why search
     * through all the ptg/formula stuff every time you deal with a XLS record an vice-versa.  Also,
     * init'ing may be different.  Small duplication of code, but I think it is worth it.
     */
        @Throws(InvalidRecordException::class)
        fun getPtgRecord(name: String): Ptg? {
            for (t in ptgLookup.indices) {
                if (ptgLookup[t][0].equals(name, ignoreCase = true)) {
                    try {
                        val classname = ptgLookup[t][1]
                        return Class.forName("io.starter.formats.XLS.formulas.$classname").newInstance() as Ptg
                    } catch (e: Exception) {
                        throw InvalidRecordException("ERROR: Creating Record: " + name + "failed: " + e.toString())
                    }

                }
            }
            return null
        }


        /**
         * Get an instance of the record type corresponding to the given opcode.
         *
         * @param opcode the BIFF8 record opcode to be resolved
         * @return an instance of the class corresponding to the given opcode
         * or an XLSRecord if the opcode is unknown
         * @throws RuntimeException if instantiation of the record fails
         */
        fun getBiffRecord(opcode: Short): BiffRec {
            // TRY THIS:
            var record: BiffRec? = null
            try {
                when (opcode) {
                    XLSConstants.BLANK -> record = Blank()
                    XLSConstants.ROW -> record = Row()
                    XLSConstants.XF -> record = Xf()
                    XLSConstants.INDEX -> record = Index()
                    XLSConstants.COUNTRY -> record = Country()
                    XLSConstants.CALCMODE -> record = CalcMode()
                    XLSConstants.DIMENSIONS -> record = Dimensions()
                    XLSConstants.SELECTION -> record = Selection()
                    XLSConstants.DEFAULTROWHEIGHT -> record = DefaultRowHeight()
                    XLSConstants.DEFCOLWIDTH -> record = DefColWidth()
                    XLSConstants.DBCELL -> record = Dbcell()
                    XLSConstants.BOF -> record = Bof()
                    XLSConstants.BOUNDSHEET -> record = Boundsheet()
                    XLSConstants.EOF -> record = Eof()
                    XLSConstants.FORMAT -> record = Format()
                    XLSConstants.STYLE -> record = Style()
                    XLSConstants.PASSWORD -> record = Password()
                    XLSConstants.PALETTE -> record = Palette()
                    XLSConstants.ARRAY -> record = Array()
                    XLSConstants.BOOLERR -> record = Boolerr()
                    XLSConstants.EXTERNSHEET -> record = Externsheet()
                    XLSConstants.EXTERNNAME -> record = Externname()
                    XLSConstants.FORMULA -> record = Formula()
                    XLSConstants.LABEL -> record = Label()
                    XLSConstants.TXO -> record = Txo()
                    XLSConstants.CONTINUE -> record = Continue()
                    XLSConstants.SST -> record = Sst()
                    XLSConstants.GUTS -> record = Guts()
                    XLSConstants.EXTSST -> record = Extsst()
                    XLSConstants.HLINK -> record = Hlink()
                    XLSConstants.LABELSST -> record = Labelsst()
                    XLSConstants.NUMBER -> record = NumberRec()
                    XLSConstants.MERGEDCELLS -> record = Mergedcells()
                    XLSConstants.MULBLANK -> record = Mulblank()
                    XLSConstants.MULRK -> record = Mulrk()
                    XLSConstants.RK -> record = Rk()
                    XLSConstants.RSTRING -> record = Rstring()
                    XLSConstants.SHRFMLA -> record = Shrfmla()
                    XLSConstants.STRINGREC -> record = StringRec()
                    XLSConstants.SUPBOOK -> record = Supbook()
                    XLSConstants.DV -> record = Dv()
                    XLSConstants.DVAL -> record = Dval()
                    XLSConstants.SETUP -> record = Setup()
                    XLSConstants.HCENTER -> record = HCenter()
                    XLSConstants.VCENTER -> record = VCenter()
                    XLSConstants.LEFTMARGIN -> record = LeftMargin()
                    XLSConstants.RIGHTMARGIN -> record = RightMargin()
                    XLSConstants.TOPMARGIN -> record = TopMargin()
                    XLSConstants.BOTTOMMARGIN -> record = BottomMargin()
                    XLSConstants.PRINTGRID -> record = PrintGrid()
                    XLSConstants.PRINTROWCOL -> record = PrintRowCol()
                    XLSConstants.XCT -> record = Xct()
                    XLSConstants.CRN -> record = Crn()
                    XLSConstants.NOTE -> record = Note()

                    // Named Ranges and References
                    XLSConstants.NAME -> record = Name()

                    // Workbook Settings
                    XLSConstants.FONT -> record = Font()
                    XLSConstants.DSF -> record = Dsf()
                    XLSConstants.WINDOW1 -> record = Window1()
                    XLSConstants.WINDOW2 -> record = Window2()
                    XLSConstants.PLV -> record = PLV()
                    XLSConstants.CODENAME -> record = Codename()
                    XLSConstants.PROTECT -> record = Protect()
                    XLSConstants.OBJPROTECT -> record = ObjProtect()
                    XLSConstants.SCENPROTECT -> record = ScenProtect()
                    XLSConstants.FEATHEADR -> record = FeatHeadr()
                    XLSConstants.PROT4REV -> record = Prot4rev()
                    XLSConstants.COLINFO -> record = Colinfo()
                    XLSConstants.USERSVIEWBEGIN -> record = Usersviewbegin()
                    XLSConstants.USERSVIEWEND -> record = Usersviewend()
                    XLSConstants.WSBOOL -> record = WsBool()
                    XLSConstants.BOOKBOOL -> record = BookBool()
                    XLSConstants.USR_EXCL -> record = UsrExcl()
                    XLSConstants.INTERFACE_HDR -> record = InterfaceHdr()
                    XLSConstants.RRD_INFO -> record = RrdInfo()
                    XLSConstants.RRD_HEAD -> record = RrdHead()
                    XLSConstants.FILE_LOCK -> record = FileLock()
                    XLSConstants.PLS -> record = Pls()
                    XLSConstants.HEADERREC -> record = Headerrec()
                    XLSConstants.DATE1904 -> record = NineteenOhFour()

                    // Sheet Settings
                    XLSConstants.OBJ -> record = Obj()
                    XLSConstants.OBPROJ -> record = Obproj()
                    XLSConstants.FOOTERREC -> record = Footerrec()
                    XLSConstants.TABID -> record = TabID()
                    XLSConstants.PANE -> record = Pane()
                    XLSConstants.SCL -> record = Scl()

                    // Conditional Formatting
                    XLSConstants.CF -> record = Cf()
                    XLSConstants.CONDFMT -> record = Condfmt()

                    // Auto filter
                    XLSConstants.AUTOFILTER -> record = AutoFilter()

                    // Chart Records
                    XLSConstants.CHART -> record = Chart()
                    XLSConstants.SERIES -> record = Series()
                    XLSConstants.SERIESTEXT -> record = SeriesText()
                    XLSConstants.SERIESLIST -> record = SeriesList()
                    XLSConstants.AI -> record = Ai()
                    XLSConstants.UNITS -> record = Units()
                    XLSConstants.DATAFORMAT -> record = DataFormat()
                    XLSConstants.LINEFORMAT -> record = LineFormat()
                    XLSConstants.MARKERFORMAT -> record = MarkerFormat()
                    XLSConstants.AREAFORMAT -> record = AreaFormat()
                    XLSConstants.PIEFORMAT -> record = PieFormat()
                    XLSConstants.ATTACHEDLABEL -> record = AttachedLabel()
                    XLSConstants.CHARTFORMAT -> record = ChartFormat()
                    XLSConstants.LEGEND -> record = Legend()
                    XLSConstants.BAR -> record = Bar()
                    XLSConstants.LINE -> record = Line()
                    XLSConstants.PIE -> record = Pie()
                    XLSConstants.AREA -> record = Area()
                    XLSConstants.SCATTER -> record = Scatter()
                    XLSConstants.CHARTLINE -> record = ChartLine()
                    XLSConstants.AXIS -> record = Axis()
                    XLSConstants.TICK -> record = Tick()
                    XLSConstants.VALUERANGE -> record = ValueRange()
                    XLSConstants.CATSERRANGE -> record = CatserRange()
                    XLSConstants.AXISLINEFORMAT -> record = AxisLineFormat()
                    XLSConstants.CHARTFORMATLINK -> record = ChartFormatLink()
                    XLSConstants.DEFAULTTEXT -> record = DefaultText()
                    XLSConstants.TEXTDISP -> record = TextDisp()
                    XLSConstants.FONTX -> record = Fontx()
                    XLSConstants.OBJECTLINK -> record = ObjectLink()
                    XLSConstants.FRAME -> record = Frame()
                    XLSConstants.BEGIN -> record = Begin()
                    XLSConstants.END -> record = End()
                    XLSConstants.PLOTAREA -> record = PlotArea()
                    XLSConstants.THREED -> record = ThreeD()
                    XLSConstants.PICF -> record = Picf()
                    XLSConstants.DROPBAR -> record = Dropbar()
                    XLSConstants.RADAR -> record = Radar()
                    XLSConstants.SURFACE -> record = Surface()
                    XLSConstants.RADARAREA -> record = RadarArea()
                    XLSConstants.AXISPARENT -> record = AxisParent()
                    XLSConstants.LEGENDXN -> record = Legendxn()
                    XLSConstants.SHTPROPS -> record = ShtProps()
                    XLSConstants.SERTOCRT -> record = SerToCrt()
                    XLSConstants.AXESUSED -> record = Axesused()
                    XLSConstants.SBASEREF -> record = SbaseRef()
                    XLSConstants.SERPARENT -> record = SerParent()
                    XLSConstants.SERAUXTREND -> record = SerauxTrend()
                    XLSConstants.IFMT -> record = Ifmt()
                    XLSConstants.POS -> record = Pos()
                    XLSConstants.ALRUNS -> record = AlRuns()
                    XLSConstants.SERAUXERRBAR -> record = SerauxErrBar()
                    XLSConstants.SERFMT -> record = Serfmt()
                    XLSConstants.CHART3DBARSHAPE -> record = Chart3DBarShape()
                    XLSConstants.FBI -> record = Fbi()
                    XLSConstants.BOPPOP -> record = Boppop()
                    XLSConstants.AXCENT -> record = Axcent()
                    XLSConstants.DAT -> record = Dat()
                    XLSConstants.PLOTGROWTH -> record = PlotGrowth()
                    XLSConstants.SIIINDEX -> record = SiIndex()
                    XLSConstants.GELFRAME -> record = GelFrame()
                    XLSConstants.BOPPOPCUSTOM -> record = BoppopCustom()
                    XLSConstants.FONTBASIS -> record = FontBasis()

                    // PivotTable Records
                    XLSConstants.SXVIEW -> record = Sxview()
                    XLSConstants.TABLESTYLES -> record = TableStyles()
                    XLSConstants.SXFORMAT -> record = Sxformat()
                    XLSConstants.SXLI -> record = Sxli()
                    XLSConstants.SXVI -> record = Sxvi()
                    XLSConstants.SXVD -> record = Sxvd()
                    XLSConstants.SXIVD -> record = Sxivd()
                    XLSConstants.SXSTREAMID -> record = SxStreamID()
                    XLSConstants.SXVS -> record = SxVS()
                    XLSConstants.SXADDL -> record = SxAddl()
                    XLSConstants.SXVDEX -> record = SxVdEX()
                    XLSConstants.SXPI -> record = SxPI()
                    XLSConstants.SXDI -> record = SxDI()
                    XLSConstants.SXDB -> record = SxDB()
                    XLSConstants.SXFDB -> record = SxFDB()
                    XLSConstants.SXDBEX -> record = SXDBEx()
                    XLSConstants.SXFDBTYPE -> record = SXFDBType()
                    XLSConstants.SXSTRING -> record = SXString()
                    XLSConstants.SXNUM -> record = SXNum()
                    XLSConstants.SXDBB -> record = SxDBB()
                    XLSConstants.SXEX -> record = SxEX()
                    XLSConstants.QSISXTAG -> record = QsiSXTag()
                    XLSConstants.SXVIEWEX9 -> record = SxVIEWEX9()
                    XLSConstants.DCONREF -> record = DConRef()
                    XLSConstants.DCONNAME -> record = DConName()
                    XLSConstants.DCONBIN -> record = DConBin()


                    // Object and Picture Records
                    XLSConstants.PHONETIC -> record = Phonetic()
                    XLSConstants.MSODRAWING -> record = MSODrawing()
                    XLSConstants.MSODRAWINGGROUP -> record = MSODrawingGroup()
                    XLSConstants.MSODRAWINGSELECTION -> record = MSODrawingSelection()

                    // Excel 9 Chart Records
                    XLSConstants.CHARTFRTINFO -> record = ChartFrtInfo()
                    XLSConstants.FRTWRAPPER -> record = FrtWrapper()
                    XLSConstants.STARTBLOCK -> record = StartBlock()
                    XLSConstants.ENDBLOCK -> record = EndBlock()
                    XLSConstants.STARTOBJECT -> record = StartObject()
                    XLSConstants.ENDOBJECT -> record = EndObject()
                    XLSConstants.CATLAB -> record = CatLab()
                    XLSConstants.YMULT -> record = YMult()
                    XLSConstants.SXVIEWLINK -> record = SxViewLink()
                    XLSConstants.PIVOTCHARTBITS -> record = PivotChartBits()
                    XLSConstants.FRTFONTLIST -> record = FrtFontList()
                    XLSConstants.PIVOTCHARTLINK -> record = PivotChartLink()
                    XLSConstants.DATALABEXTCONTENTS -> record = DataLabExtContents()
                    XLSConstants.DATALABEXT -> record = DataLabExt()
                    XLSConstants.CRTLAYOUT12A -> record = CrtLayout12A()
                    XLSConstants.CRTLAYOUT12 -> record = CrtLayout12()
                    else -> record = XLSRecord()
                }
                record.opcode = opcode
            } catch (e: Exception) {
                throw RuntimeException("failed to instantiate record", e)
            }

            return record
        }
    }

}