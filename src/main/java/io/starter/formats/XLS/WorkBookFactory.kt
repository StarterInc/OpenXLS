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

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.LEO.BlockByteConsumer
import io.starter.formats.LEO.BlockByteReader
import io.starter.formats.LEO.LEOFile
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.ProgressListener

import java.io.Serializable
import java.util.LinkedHashMap


/**
 * Factory for creating WorkBook objects from Byte streams.
 *
 * @see WorkBook
 *
 * @see XLSRecord
 */
class WorkBookFactory : io.starter.toolkit.ProgressNotifier, XLSConstants, Serializable {
    var debugLevel = 0

    internal var leoFile: LEOFile
        set
    // end ProgressNotifier methods

    /**
     * get the file name for the WorkBook
     */
    /**
     * sets the workbook filename associated with this wbfactory
     *
     * @param f
     */
    var fileName: String? = null
        get() {
            if (field == null)
                fileName = leoFile.fileName
            return field
        }

    // Methods from ProgressNotifier
    private var progresslistener: ProgressListener? = null
    override var progress = 0
    private var done = false
    override var progressText = ""

    override fun register(j: ProgressListener) {
        progresslistener = j
        j.addTarget(this)
    }

    override fun fireProgressChanged() {
        // if (progresslistener != null) {
        // progresslistener.updateProgress();
        // }
    }

    override fun iscompleted(): Boolean {
        return done
    }

    /**
     * return the next opcode/length in the Stream from the given record.
     */
    fun lookAhead(rec: BiffRec): Short {
        val i = rec.offset + rec.length
        val parsedata = rec.byteReader ?: return rec.opcode

        val b1 = parsedata.getHeaderBytes(i)

        return ByteTools.readShort(b1[0].toInt(), b1[1].toInt())
    }

    /**
     * read in a WorkBook from a byte array.
     */
    @Throws(InvalidRecordException::class)
    fun getWorkBook(parsedata: BlockByteReader, leo: LEOFile): Book {
        val book = WorkBook()
        return this.initWorkBook(book, parsedata, leo)
    }

    /**
     * Initialize the workbook
     *
     * @param book
     * @param parsedata
     * @param leo
     * @return
     * @throws InvalidRecordException
     */
    @Throws(InvalidRecordException::class)
    fun initWorkBook(book: Book, parsedata: BlockByteReader, leo: LEOFile): Book {

        parsedata.applyRelativePosition = true

        /** KSC: record-level validation  */
        var bPerformRecordLevelValidation = false    // perform record-level validation if set
        if (System.getProperty(XLSConstants.VALIDATEWORKBOOK) != null) {
            if (System.getProperty(XLSConstants.VALIDATEWORKBOOK) == "true")
                bPerformRecordLevelValidation = true
        }
        var curSubstream: java.util.LinkedHashMap<Short, R>? = null
        var sheetSubstream: java.util.LinkedHashMap<Short, R>? = null
        if (bPerformRecordLevelValidation) {
            val globalSubstream = java.util.LinkedHashMap()
            fillGlobalSubstream(globalSubstream)
            sheetSubstream = java.util.LinkedHashMap()
            fillWorksSheetSubstream(sheetSubstream)
            curSubstream = globalSubstream
        }

        leoFile = leo

        book.setDebugLevel(this.debugLevel)
        book.factory = this
        var infile = false
        var isWBBOF = true
        var opcode: Short = 0x00
        var reclen: Short = 0x00
        var lastOpcode: Short = 0x00
        var BofCount = 0 // track the number of 'Bof' records

        var rec: BiffRec? = null
        val blen = parsedata.length

        // init the progress listener
        this.progressText = "Initializing Workbook..."
        progress = 0
        if (progresslistener != null)
            progresslistener!!.setMaxProgress(blen)
        this.fireProgressChanged()
        if (debugLevel > 1)
            Logger.logInfo("XLS File Size: $blen")

        var i = 0
        while (i <= blen - 4) {

            this.fireProgressChanged() // ""
            val headerbytes = parsedata.getHeaderBytes(i)
            opcode = ByteTools.readShort(headerbytes[0].toInt(), headerbytes[1].toInt())
            reclen = ByteTools.readShort(headerbytes[2].toInt(), headerbytes[3].toInt())

            if (lastOpcode == XLSConstants.EOF && opcode.toInt() == 0 || opcode.toInt() == -0x1) {
                val startpos = i - 3
                var junkreclen = 0
                val offset = 0

                if (offset != 0)
                    junkreclen = offset - startpos
                else
                    junkreclen = blen - i
                i += junkreclen - 1
                i += junkreclen - 1
            } else { // REAL REC
                // sanity checks
                if (reclen < 0)
                    throw InvalidRecordException("WorkBookFactory.getWorkBook() Negative Reclen encountered pos:" + i + " opcode:0x" + Integer.toHexString(opcode.toInt()))
                else if (reclen + 1 > blen)
                    throw InvalidRecordException("WorkBookFactory.getWorkBook() Reclen longer than data pos:" + i + " opcode:0x" + Integer.toHexString(opcode.toInt()))

                if (opcode == XLSConstants.BOF || infile) { // if the first Bof has been
                    // reached, start
                    infile = true

                    // Init Record'
                    rec = this.parse(book, opcode, i, reclen.toInt(), parsedata)

                    if (progresslistener != null) {
                        progresslistener!!.setValue(i)
                    }

                    /**** KSC: record-level validation  */
                    if (bPerformRecordLevelValidation && curSubstream != null) {
                        markRecord(curSubstream, rec, opcode)
                    }

                    // write to the dump file if necessary
                    if (WorkBookHandle.dump_input != null)
                        try {
                            WorkBookHandle.dump_input!!.write("-------------------------------------" + "-------------------------\n" + (rec as XLSRecord).recDesc + ByteTools.getByteDump(parsedata.get((rec as XLSRecord?)!!, 0, reclen.toInt()), 0) + "\n")
                            WorkBookHandle.dump_input!!.flush()
                        } catch (e: Exception) {
                            Logger.logErr("error writing to dump file, ceasing dump output: $e")
                            WorkBookHandle.dump_input = null
                        }

                    if (rec == null) { // Effectively an EOF
                        if (debugLevel > 0)
                            Logger.logInfo("done parsing WorkBook storage.")
                        this.done = true
                        this.progressText = "Done Reading WorkBook."
                        this.fireProgressChanged()
                        return book
                    } else {
                        // not used anymore ((XLSRecord)rec).resetCacheBytes();
                    }
                    // int reco = rec.getOffset() ;
                    // int recl = rec.getLength();
                    var thisrecpos = i + reclen.toInt() + 4

                    if (opcode == XLSConstants.BOF) {
                        if (isWBBOF) {    // do first Bof initialization
                            book.setFirstBof(rec as io.starter.formats.XLS.Bof?)
                            isWBBOF = false
                        } else if (bPerformRecordLevelValidation && BofCount == 0 && lastOpcode != XLSConstants.EOF && curSubstream != null) {
                            /***** KSC: record-level validation  */
                            // invalid record structure-  no EOF before BOF
                            validateRecords(curSubstream, book, rec.sheet)
                        }

                        BofCount++
                        /***** KSC: record-level validation  */
                        if (bPerformRecordLevelValidation && curSubstream == null) {
                            // after global substream is processed, switch to sheet substream
                            reInitSubstream(sheetSubstream!!)
                            curSubstream = sheetSubstream
                            curSubstream[XLSConstants.BOF].isPresent = true
                            curSubstream[XLSConstants.BOF].recordPos = 0
                        }
                    } else if (opcode == XLSConstants.EOF) {
                        BofCount--

                        /***** KSC: record-level validation  */
                        if (bPerformRecordLevelValidation && BofCount == 0 && curSubstream != null) {
                            validateRecords(curSubstream, book, rec.sheet)
                            curSubstream = null
                        }
                    }
                    // end of Workbook
                    if (BofCount == -1) {
                        if (debugLevel > 2)
                            Logger.logInfo("Last Bof")
                        i += reclen.toInt()
                        thisrecpos = blen
                    }
                    if (thisrecpos > 0)
                        i = thisrecpos
                    lastOpcode = opcode
                } else {
                    throw InvalidRecordException("No valid record found.")
                }

            }
        }
        if (debugLevel > 0)
            Logger.logInfo("done")
        progress = blen
        this.progressText = "Done Reading WorkBook."
        this.fireProgressChanged()

        this.done = true
        // flag the book so we know it's ready for shared access
        // book.setReady(true); ENTERPRISE ONLY
        // recordata.setApplyRelativePosition(false);
        return book
    }

    /**
     * create the individual records based on type
     */
    @Synchronized
    @Throws(InvalidRecordException::class)
    protected fun parse(book: Book, opcode: Short, offset: Int, datalen: Int, bytebuf: BlockByteReader): BiffRec {

        // sanity checks
        if (datalen < 0 || datalen > XLSConstants.MAXRECLEN)
            throw InvalidRecordException("InvalidRecordException BAD RECORD LENGTH: " + " off: " + offset + " op: " + Integer.toHexString(opcode.toInt()) + " len: " + datalen)
        if (offset + datalen > bytebuf.length)
            throw InvalidRecordException("InvalidRecordException RECORD LENGTH LONGER THAN FILE: " + " off: " + offset + " op: " + Integer.toHexString(opcode.toInt()) + " len: " + datalen + " buflen:" + bytebuf.length)

        // Create a new Record
        val rec = XLSRecordFactory.getBiffRecord(opcode)

        // init the mighty rec
        rec.workBook = book as WorkBook
        rec.byteReader = bytebuf
        rec.length = datalen.toShort()
        rec.offset = offset
        rec.setDebugLevel(this.debugLevel)
        rec.streamer = book.streamer

        // send it to the CONTINUE handler
        book.continueHandler.addRec(rec, datalen.toShort().toInt())
        // add it to the record stream
        return book.addRecord(rec, true)
    }

    /**
     * record-level validation:
     *  * have ordered list of records for each substream (wb/global, worksheet)
     *  * upon each record that is processed, look up in list and mark present + record pos in correseponding list (streamer or sheet)
     *  * upon EOF for stream, traverse thru list, if required and not present, add
     * <br></br>Limitations:
     * This methodology does NOT validate Chart records, Chart-only sheets, Macro Sheets, Dialog Sheets or Custom Views
     */
    /**
     * fill map with EVERY possible global (workbook)-level record, IN ORDER, plus flag if they are required or not
     * Used in record-level validation
     *
     * @param map
     */
    private fun fillGlobalSubstream(map: LinkedHashMap<Short, R>) {
        // ordered list of all records in the global (workbook) substream, along
        // with if they are required or not
        map[XLSConstants.BOF] = R(true)
        map[134.toShort()] = R(false) // WriteProtect
        map[XLSConstants.FILEPASS] = R(false)
        map[96.toShort()] = R(false) // Template
        map[XLSConstants.INTERFACE_HDR] = R(true)
        map[193.toShort()] = R(true) // Mms
        map[226.toShort()] = R(true) // InterfaceEnd
        map[92.toShort()] = R(true) // WriteAccess
        map[91.toShort()] = R(false) // FileSharing
        map[66.toShort()] = R(true) // CodePage
        map[441.toShort()] = R(false) // LEL - SHOULD BE is required??
        map[353.toShort()] = R(true) // DSF - is required??
        map[448.toShort()] = R(false) // Excel9File
        map[XLSConstants.TABID] = R(true)
        map[211.toShort()] = R(false) // ObjProj
        map[445.toShort()] = R(false) // *[ObNoMacros]
        map[XLSConstants.CODENAME] = R(false)
        map[156.toShort()] = R(false) // BuiltInFnGroupCount
        map[154.toShort()] = R(false) // FnGroupName
        map[2200.toShort()] = R(false) // FnGrp12
        map[222.toShort()] = R(false) // OleObjectSize
        map[XLSConstants.WINDOW_PROTECT] = R(true)
        map[XLSConstants.PROTECT] = R(true)
        map[XLSConstants.PASSWORD] = R(true)
        map[XLSConstants.PROT4REV] = R(true)
        map[444.toShort()] = R(true) // Prot4RevPass
        map[XLSConstants.WINDOW1] = R(true)
        map[XLSConstants.BACKUP] = R(true)
        map[141.toShort()] = R(true) // HideObj
        map[XLSConstants.DATE1904] = R(true)
        map[14.toShort()] = R(true) // CalcPrecision
        map[439.toShort()] = R(true) // RefreshAll
        map[XLSConstants.BOOKBOOL] = R(true)
        map[XLSConstants.FONT] = R(true) // Is required???
        map[XLSConstants.FORMAT] = R(true) // Is required??
        map[XLSConstants.XF] = R(true) // Is required???
        map[2172.toShort()] = R(false) // XfCRC
        map[2173.toShort()] = R(false) // XfExt
        map[2189.toShort()] = R(false) // DXF- Is required???
        map[XLSConstants.STYLE] = R(true) // Is required???
        map[2194.toShort()] = R(false) // StyleExt
        map[XLSConstants.TABLESTYLES] = R(false)
        map[2191.toShort()] = R(false) // TableStyle
        map[2192.toShort()] = R(false) // TABLESTYLEELEMENT
        map[XLSConstants.PALETTE] = R(false)
        map[4188.toShort()] = R(false) // CLRTCLIENT
        map[XLSConstants.SXSTREAMID] = R(false)
        map[XLSConstants.SXVS] = R(false)
        map[XLSConstants.DCONNAME] = R(false)
        map[XLSConstants.DCONBIN] = R(false)
        map[XLSConstants.DCONREF] = R(false)
        map[208.toShort()] = R(false) // SXTbl
        map[210.toShort()] = R(false) // SxTbpg
        map[209.toShort()] = R(false) // SXTBRGIITM
        map[XLSConstants.SXSTRING] = R(false)
        map[220.toShort()] = R(false) // DbOrParamQry
        map[XLSConstants.SXADDL] = R(false)

        map[184.toShort()] = R(false) // DocRoute
        map[185.toShort()] = R(false) // RECIPNAME
        map[XLSConstants.USERBVIEW] = R(false) // SHOULD BE Is but isn't present
        // *************************
        map[352.toShort()] = R(true) // UsesELFs
        map[XLSConstants.BOUNDSHEET] = R(true)
        map[2180.toShort()] = R(false) // MDTInfo
        // *ContinueFrt12
        map[2181.toShort()] = R(false) // MDXStr
        // *ContinueFrt12
        map[2182.toShort()] = R(false) // MDXTuple
        // *ContinueFrt12
        map[2183.toShort()] = R(false) // MDXSet
        map[2184.toShort()] = R(false) // MDXProp
        map[2185.toShort()] = R(false) // MDXKPI
        map[2186.toShort()] = R(false) // MDB
        map[2202.toShort()] = R(false) // MTRSettings
        map[2211.toShort()] = R(false) // ForceFullCalculation
        map[XLSConstants.COUNTRY] = R(true)
        map[XLSConstants.SUPBOOK] = R(false) // SHOULD BE IsRequired but isn't present ******************
        map[XLSConstants.EXTERNNAME] = R(false)
        map[XLSConstants.XCT] = R(false)
        map[XLSConstants.CRN] = R(false)
        map[XLSConstants.EXTERNSHEET] = R(false)
        map[XLSConstants.NAME] = R(false) // Name is Required???
        map[2196.toShort()] = R(false) // NameCmt
        map[2201.toShort()] = R(false) // NameFnGrp12
        map[2195.toShort()] = R(false) // NamePublish
        map[2067.toShort()] = R(false) // RealTimeData A required record
        // is not present: 430
        // *ContinueFrt
        map[449.toShort()] = R(false) // RecalcId
        map[2150.toShort()] = R(false) // HFPicture should be required????
        map[XLSConstants.MSODRAWINGGROUP] = R(false) // should be required??
        // Continue
        map[XLSConstants.SST] = R(true)        // should be required??
        // Continue
        map[XLSConstants.EXTSST] = R(true)
        map[2049.toShort()] = R(false) // WebPub should be required???
        map[2059.toShort()] = R(false) // WOpt
        map[2149.toShort()] = R(false) // CrErr
        map[2147.toShort()] = R(false) // BookExt
        map[2151.toShort()] = R(false) // FeatHdr should be required???
        map[2166.toShort()] = R(false) // DConn should be required???
        map[2198.toShort()] = R(false) // Theme
        // *ContinueFrt12
        map[2203.toShort()] = R(false) // CompressPictures
        map[2188.toShort()] = R(false) // Compat12
        map[2199.toShort()] = R(false) // GUIDTypeLib
        map[XLSConstants.EOF] = R(true)
    }

    /**
     * fill map with EVERY possible worksheet-level record, IN ORDER, plus flag if they are required or not.
     * Used in record-level validation
     *
     * @param map
     */
    private fun fillWorksSheetSubstream(map: LinkedHashMap<Short, R>) {
        map[XLSConstants.BOF] = R(true)
        map[94.toShort()] = R(false)    //	Uncalced
        map[XLSConstants.INDEX] = R(true)
        map[XLSConstants.CALCMODE] = R(true)
        map[XLSConstants.CALCCOUNT] = R(true)
        map[15.toShort()] = R(true)    //CalcRefMode
        map[17.toShort()] = R(true)    // CalcIter
        map[16.toShort()] = R(true)    // CalcDelta
        map[95.toShort()] = R(true)    // CalcSaveRecalc
        map[XLSConstants.PRINTROWCOL] = R(true)
        map[XLSConstants.PRINTGRID] = R(true)
        map[130.toShort()] = R(true)    // GridSet
        map[XLSConstants.GUTS] = R(true)
        map[XLSConstants.DEFAULTROWHEIGHT] = R(true)
        map[XLSConstants.WSBOOL] = R(true)
        map[151.toShort()] = R(false) //	[Sync]
        map[152.toShort()] = R(false) //	[LPr]
        map[XLSConstants.HORIZONTAL_PAGE_BREAKS] = R(false) //[HorizontalPageBreaks]
        map[XLSConstants.VERTICAL_PAGE_BREAKS] = R(false)    //[VerticalPageBreaks]
        map[XLSConstants.HEADERREC] = R(true)
        map[XLSConstants.FOOTERREC] = R(true)
        map[XLSConstants.HCENTER] = R(true)
        map[XLSConstants.VCENTER] = R(true)
        map[XLSConstants.LEFT_MARGIN] = R(false)
        map[XLSConstants.RIGHT_MARGIN] = R(false)
        map[XLSConstants.TOP_MARGIN] = R(false)
        map[XLSConstants.BOTTOM_MARGIN] = R(false)
        map[XLSConstants.PLS] = R(false)
        //	Continue
        map[XLSConstants.SETUP] = R(true)
        map[0x89c.toShort()] = R(false)    //[HeaderFooter]
        map[233.toShort()] = R(false)    //[ BkHim ]
        map[1048.toShort()] = R(false)    // BigName
        //       *ContinueBigName
        map[XLSConstants.PROTECT] = R(false)
        map[XLSConstants.SCENPROTECT] = R(false)
        map[XLSConstants.OBJPROTECT] = R(false)
        map[XLSConstants.PASSWORD] = R(false)
        map[XLSConstants.DEFCOLWIDTH] = R(true)
        map[XLSConstants.COLINFO] = R(false)
        map[174.toShort()] = R(false)    // [ScenMan *(
        map[175.toShort()] = R(false)    // SCENARIO
        // Continue
        map[144.toShort()] = R(false)    // Sort
        map[2197.toShort()] = R(false)    // SortData
        // ContinueFrt12
        map[155.toShort()] = R(false)    // Filtermode
        map[2164.toShort()] = R(false)    // DropDownObjIds
        map[157.toShort()] = R(false)    // AutoFilterInfo
        map[XLSConstants.AUTOFILTER] = R(false)
        map[2174.toShort()] = R(false)    // AutoFilter12
        map[2197.toShort()] = R(false)    // SortData
        //*ContinueFrt12]
        map[XLSConstants.DIMENSIONS] = R(true)
        //	[CELLTABLE]
        map[XLSConstants.ROW] = R(false)
        // celltable entries
        map[XLSConstants.BLANK] = R(false)
        map[XLSConstants.MULBLANK] = R(false)
        map[XLSConstants.RK] = R(false)
        map[XLSConstants.BOOLERR] = R(false)
        map[XLSConstants.NUMBER] = R(false)
        map[XLSConstants.LABELSST] = R(false)
        map[94.toShort()] = R(false)    // uncalced
        map[XLSConstants.FORMULA] = R(false)
        map[XLSConstants.MULRK] = R(false)
        map[XLSConstants.STRINGREC] = R(false)
        map[XLSConstants.SHRFMLA] = R(false)
        map[XLSConstants.ARRAY] = R(false)
        map[XLSConstants.TABLE] = R(false)
        map[450.toShort()] = R(false)    // EntExU2
        //	[OBJECTS]
        map[XLSConstants.MSODRAWING] = R(false)
        map[XLSConstants.TXO] = R(false)
        map[XLSConstants.OBJ] = R(false)
        map[XLSConstants.OBJ].altPrecedor = shortArrayOf(XLSConstants.MSODRAWING)        // Obj can follow MSODRAWING or TXO
        //Charts -- most record opcodes are > 4000 - ignore !!
        map[XLSConstants.CHARTFRTINFO] = R(false)
        // do not include the majority of chart records but do include these:
        map[XLSConstants.STARTBLOCK] = R(false)
        map[XLSConstants.CRTLAYOUT12] = R(false)
        map[XLSConstants.CRTLAYOUT12A] = R(false)
        map[XLSConstants.CATLAB] = R(false)
        map[XLSConstants.DATALABEXT] = R(false)
        map[XLSConstants.DATALABEXTCONTENTS] = R(false)
        map[2138.toShort()] = R(false)    // FrtFontList
        map[2213.toShort()] = R(false)    // TextPropsStream
        map[2214.toShort()] = R(false)    // RichTextStream
        map[2212.toShort()] = R(false)    // ShapePropsStream
        map[2206.toShort()] = R(false)    // CtrlMlFrt
        map[XLSConstants.ENDBLOCK] = R(false)
        map[XLSConstants.STARTOBJECT] = R(false)
        map[XLSConstants.FRTWRAPPER] = R(false)
        map[XLSConstants.ENDOBJECT] = R(false)
        //	CHARTSHEETCONTENT = BOF [WriteProtect] [SheetExt] [WebPub] *HFPicture PAGESETUP PrintSize [HeaderFooter] [BACKGROUND] *Fbi *Fbi2 [ClrtClient] [PROTECTION] [Palette] [SXViewLink] [PivotChartBits] [SBaseRef] [MsoDrawingGroup] OBJECTS Units CHARTFOMATS SERIESDATA *WINDOW *CUSTOMVIEW [CodeName] [CRTMLFRT] EOF
        //Chart Begin *2FONTLIST SclPlotGrowth [FRAME] *SERIESFORMAT *SS ShtProps *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL [CRTMLFRT] *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End
        map[XLSConstants.MSODRAWINGSELECTION] = R(false)
        map[2150.toShort()] = R(false)    //*HFPicture
        map[XLSConstants.NOTE] = R(false)
        //*PIVOTVIEW
        map[XLSConstants.SXVIEW] = R(false)
        map[XLSConstants.SXVD] = R(false)
        map[XLSConstants.SXVI] = R(false)
        map[XLSConstants.SXVDEX] = R(false)
        map[XLSConstants.SXIVD] = R(false)
        map[XLSConstants.SXPI] = R(false)
        map[XLSConstants.SXDI] = R(false)
        map[XLSConstants.SXLI] = R(false)
        map[XLSConstants.SXEX] = R(false)
        map[247.toShort()] = R(false)    //SXSelect
        map[240.toShort()] = R(false)    // SXRULE
        map[XLSConstants.SXFORMAT] = R(false)
        //	PIVOTRULE
        map[244.toShort()] = R(false)    // SXDXF
        map[XLSConstants.QSISXTAG] = R(false)
        //DBQUERYEXT
        map[2051.toShort()] = R(false)    // DBQUERYEX
        map[2052.toShort()] = R(false)    // EXTSTRING
        map[XLSConstants.SXSTRING] = R(false)
        map[2060.toShort()] = R(false)    // SXVIEWEX
        map[2061.toShort()] = R(false)    // SXTH
        map[2062.toShort()] = R(false)    // SXPIEx
        map[256.toShort()] = R(false)    // SXVDTEx
        map[XLSConstants.SXVIEWEX9] = R(false)
        map[XLSConstants.SXADDL] = R(false)

        map[XLSConstants.DCON] = R(false)
        map[XLSConstants.DCONNAME] = R(false)
        map[XLSConstants.DCONBIN] = R(false)
        map[XLSConstants.DCONREF] = R(false)
        map[XLSConstants.WINDOW2] = R(true)
        map[XLSConstants.PLV] = R(false)
        map[XLSConstants.SCL] = R(false)
        map[XLSConstants.PANE] = R(false)
        map[XLSConstants.SELECTION] = R(false)
        map[426.toShort()] = R(false)    // UserSViewBegin
        /*		map.put(SELECTION, new R(false));
		map.put(HORIZONTAL_PAGE_BREAKS, new R(false));
		map.put(VERTICAL_PAGE_BREAKS, new R(false));
		map.put(HEADERREC, new R(false));
		map.put(FOOTERREC, new R(false));
		map.put(HCENTER,  new R(false));
		map.put(VCENTER, new R(false));
		map.put(LEFT_MARGIN, new R(false));
		map.put(RIGHT_MARGIN,  new R(false));
		map.put(TOP_MARGIN,  new R(false));
		map.put(BOTTOM_MARGIN,  new R(false));
		map.put(PLS, new R(false));
		map.put(SETUP, new R(false));
		map.put((short) 51, new R(false));	//([PrintSize]
		map.put((short) 2204, new R(false));	//[HeaderFooter]*/
        // here ??? [AUTOFILTER]
        map[427.toShort()] = R(false)    //UserSViewEnd

        map[319.toShort()] = R(false)    // RRSort
        map[153.toShort()] = R(false)    //[DxGCol]
        map[XLSConstants.MERGEDCELLS] = R(false)    //*MergeCells
        map[351.toShort()] = R(false)    // [LRng]
        //	*QUERYTABLE
        map[XLSConstants.PHONETIC] = R(false)
        map[XLSConstants.CONDFMT] = R(false)
        map[XLSConstants.CF] = R(false)
        map[XLSConstants.CONDFMT12] = R(false)
        map[XLSConstants.CF12] = R(false)
        map[2171.toShort()] = R(false)    // CFEx
        map[XLSConstants.HLINK] = R(false)
        map[2048.toShort()] = R(false)    //HLinkTooltip
        map[XLSConstants.DVAL] = R(false)
        map[XLSConstants.DV] = R(false)
        map[XLSConstants.CODENAME] = R(false)    // [CodeName]
        map[2049.toShort()] = R(false)    // *WebPub
        map[2156.toShort()] = R(false)    // *CellWatch
        map[2146.toShort()] = R(false)    // SheetExt]
        map[XLSConstants.FEATHEADR] = R(false)
        map[2152.toShort()] = R(false)    // FEAT
        //		*FEAT11
        //		*RECORD12
        map[2248.toShort()] = R(false)    // UNKNOWN RECORD!!!
        map[XLSConstants.EOF] = R(true)
    }

    /**
     * add all missing REQUIRED records from the appropriate substream for record-level validation
     *
     * @param map  substream list storing if present, required ...
     * @param book
     */
    private fun validateRecords(map: LinkedHashMap<Short, R>, book: Book, bs: Boundsheet?) {
        // use array instead of iterator as may have to traverse more than once
        var opcodes = arrayOfNulls<Short>(0)
        val ss = map.keys
        opcodes = ss.toTypedArray<Short>()
        var lastR = R(false)
        lastR.recordPos = 0
        var lastOp = -1
        if (bs != null && bs.isChartOnlySheet)
            return     // don't validate chart-only sheets

        // traverse thru stream list, ensuring required records are present; create if not
        for (i in opcodes.indices) {
            val op = opcodes[i]
            val r = map.get(op)
            if (!r.isPresent && r.isRequired) {
                // io.starter.toolkit.Logger.log("A required record is not present: " +  op);
                // Create a new Record
                val rec = createMissingRequiredRecord(op, book, bs)
                rec!!.setDebugLevel(this.debugLevel)
                var recPos = lastR.recordPos + 1
                try {
                    while (true) {    // now get to LAST record if there are multiple records
                        val lastRec = if (bs == null) book.streamer.getRecordAt(recPos) else bs.sheetRecs[recPos] as BiffRec
                        if (lastRec.opcode.toInt() != lastOp)
                            break
                        recPos++
                    }
                } catch (e: IndexOutOfBoundsException) {
                }

                book.streamer.addRecordAt(rec, recPos)
                book.addRecord(rec, false)    // false= already added to streamer or sheet
                if (bs != null)
                    rec.setSheet(bs)

                r.recordPos = rec.recordIndex

                // now must adjust ensuing record positions to account for inserted record
                for (zz in opcodes.indices) {
                    val nextR = map.get(opcodes[zz])
                    if (nextR.isPresent && r != nextR && nextR.recordPos >= r.recordPos)
                        nextR.recordPos++
                }
                r.isPresent = true
            }
            if (r.isPresent) {
                lastR = r
                lastOp = op.toInt()
            }
        }
        // go thru 1 more time to ensure record order is correct
        validateRecordOrder(map, if (bs == null) book.streamer.records else bs.sheetRecs, opcodes)
    }

    /**
     * given substream map, actual list of records and list of ordered opcodes, validate record order in substream
     * <br></br>if necessary, reorgainze records to follow order in map
     *
     * @param map     ordered map of substream records indexed by opcodes
     * @param list    list of actual records
     * @param opcodes ordered list of opcodes
     */
    private fun validateRecordOrder(map: LinkedHashMap<Short, R>, list: MutableList<*>, opcodes: Array<Short>) {
        /* debugging:
	io.starter.toolkit.Logger.log("BeFORE order:");
	for (int zz= 0; zz < list.size(); zz++) {
		    io.starter.toolkit.Logger.log(zz + "-" + list.get(zz));
	}*/
        var lastR = map[XLSConstants.BOF]
        var lastOp = XLSConstants.BOF
        for (i in 1 until opcodes.size) {
            val op = opcodes[i]
            val r = map[op]
            if (r.isPresent) {
                // compare ordered list of records (R) to actual order (denoted via recordPos)
                if (r.recordPos < lastR.recordPos && r.recordPos >= 0) { // Out Of Order (NOTE: CellTable entries will have a record pos = -1)
                    if (r.altPrecedor != null) {    // record can have more than 1 valid predecessor
                        // TODO: this is ugly - do a different way ...
                        for (zz in r.altPrecedor!!.indices) {
                            val newR = map[r.altPrecedor!![zz]]
                            if (r.recordPos > newR.recordPos) {
                                lastR = newR
                                break
                            }
                        }
                        if (r.recordPos > lastR.recordPos)
                            continue
                    }

                    val origRecPos = r.recordPos
                    //io.starter.toolkit.Logger.log("Record out of order:  r:" + op + "/" + r.recordPos + " lastR:" + lastOp + "/" + lastR.recordPos);
                    // find correct insertion point by looking at ordered map
                    for (zz in i - 1 downTo 1) {
                        val prevr = map[opcodes[zz]]
                        if (prevr.isPresent && r.recordPos < prevr.recordPos) {
                            //io.starter.toolkit.Logger.log("\tInsert at " + prevr.recordPos + " before op= " + opcodes[zz]);
                            var recsMovedCount = 0
                            var recToMove = list[origRecPos] as BiffRec
                            do {
                                list.removeAt(origRecPos)
                                list.add(prevr.recordPos /*+ recsMovedCount*/, recToMove)
                                if (recsMovedCount == 0)
                                    r.recordPos = recToMove.recordIndex
                                //io.starter.toolkit.Logger.log("\tMoved To " + recToMove.getRecordIndex());
                                recsMovedCount++
                                recToMove = list[origRecPos] as BiffRec
                            } while (recToMove.opcode == op)

                            // after moved all the records necessary, adjust record positions
                            for (jj in opcodes.indices) {
                                val nextR = map[opcodes[jj]]
                                if (nextR.isPresent && nextR.recordPos >= origRecPos && nextR.recordPos <= r.recordPos && opcodes[jj] != op)
                                    nextR.recordPos -= recsMovedCount
                            }
                            break
                        }
                    }
                }
                lastR = r
                lastOp = op
            }
        }
        /*io.starter.toolkit.Logger.log("AFTER order:");
	for (int zz= 0; zz < list.size(); zz++) {
		    io.starter.toolkit.Logger.log(zz + "-" + list.get(zz));
	}*/
    }

    // TODO: handle fonts,formats,xf, style -- what's minimum necessary??
    // TODO: handle TabId
    // TODO: is Index OK?

    /**
     * create missing required records for record-level validation
     *
     * @param opcode opcode of missing record
     * @param book
     * @param bs     boundsheet- if null it's a wb level record
     * @return new BiffRec
     */
    private fun createMissingRequiredRecord(opcode: Short, book: Book, bs: Boundsheet?): BiffRec? {
        val record = XLSRecordFactory.getBiffRecord(opcode)
        if (bs != null)
            record.setSheet(bs)
        var data: ByteArray? = null
        try {
            when (opcode) {
                XLSConstants.INTERFACE_HDR -> data = byteArrayOf(0xB0.toByte(), 0x04)    //codePage (2 bytes): An unsigned integer that specifies the code page. 1200==Unicode
                193.toShort()    // Mms
                -> data = byteArrayOf(0, 0)
                226.toShort()    // InterfaceEnd
                -> data = byteArrayOf()
                92.toShort()        // WriteAccess	-- user name - can be all blank
                -> data = byteArrayOf(0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20)
                66.toShort()        // CodePage	same as interfacehdr 1200==Unicode
                -> data = byteArrayOf(0xB0.toByte(), 0x04)
                353.toShort()    // DSF- truly is required?
                -> data = byteArrayOf(0, 0)
                XLSConstants.TABID            //
                -> {
                    // count how many sheets in wb
                    var nSheets = 0
                    run {
                        var z = book.streamer.records.size - 1
                        while (z > 0) {
                            var b = book.streamer.records[z] as BiffRec
                            if (b.opcode == XLSConstants.BOUNDSHEET) {
                                while (z > 0 && b.opcode == XLSConstants.BOUNDSHEET) {
                                    nSheets++
                                    b = book.streamer.records[--z] as BiffRec
                                }
                                break
                            }
                            z--
                        }
                    }
                    data = byteArrayOf()
                    for (i in 0 until nSheets) {
                        data = ByteTools.append(ByteTools.shortToLEBytes(i.toShort()), data)
                    }
                }
                XLSConstants.WINDOW_PROTECT -> data = byteArrayOf(0, 0)
                XLSConstants.PROTECT -> data = byteArrayOf(0, 0)
                XLSConstants.PASSWORD -> data = byteArrayOf(0, 0)
                XLSConstants.PROT4REV -> data = byteArrayOf(0, 0)
                444.toShort()    // Prot4RevPass
                -> data = byteArrayOf(0, 0)
                XLSConstants.WINDOW1            // very general window settings
                -> data = byteArrayOf(0xE0.toByte(), 1, 0x69, 0, 0x13, 0x38, 0x1F, 0x1D, 0x38, 0, 0, 0, 0, 0, 1, 0, 0x58, 0x02)
                XLSConstants.BACKUP -> data = byteArrayOf(0, 0)
                141.toShort()    // HideObj
                -> data = byteArrayOf(0, 0)
                XLSConstants.DATE1904 -> data = byteArrayOf(0, 0)
                14.toShort()    // CalcPrecision
                -> data = byteArrayOf(1, 0)
                439.toShort()    // RefreshAll
                -> data = byteArrayOf(0, 0)
                XLSConstants.BOOKBOOL -> data = byteArrayOf(0, 0)
                XLSConstants.FONT    // truly required??
                -> data = byteArrayOf(0xC8.toByte(), 0, 0, 0, 0xFF.toByte(), 0x7F, 0x90.toByte(), 1, 0, 0, 0, 0, 0, 0, 5, 1, 0x41, 0, 0x72, 0, 0x69, 0, 0x61, 0, 0x6C, 0)
                //		case FORMAT:	// truly required??
                //		    break;
                XLSConstants.XF    // truly is required??
                -> data = byteArrayOf(0, 0, 0, 0, 0xF5.toByte(), 0xFF.toByte(), 0x20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0xC0.toByte(), 0x20)
                //		case STYLE:	// truly is required?
                //		    break;
                XLSConstants.COUNTRY -> data = byteArrayOf(1, 0, 1, 0)
                XLSConstants.SST -> data = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0)
                XLSConstants.EXTSST -> data = byteArrayOf(0, 0)    // should be rebuilt upon output ...
                352.toShort()    // UsesELFs
                -> data = byteArrayOf(0, 0)
                XLSConstants.INDEX -> data = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0x4E, 6, 0, 0)
                XLSConstants.CALCMODE -> data = byteArrayOf(1, 0)
                XLSConstants.CALCCOUNT -> data = byteArrayOf(0x64, 0)
                15 //CalcRefMode
                -> data = byteArrayOf(1, 0)
                17    // CalcIter
                -> data = byteArrayOf(0, 0)
                16    // CalcDelta	An Xnum value that specifies the amount of change in value for a given cell from the previously calculated value for that cell that MUST exist for the iteration to continue. The value MUST be greater than or equal to 0.
                -> {
                    data = byteArrayOf(0xFC.toByte(), 0xA9.toByte(), 0xF1.toByte(), 0xD2.toByte(), 0x4D, 0x62, 0x50, 0x3F)
                    data = byteArrayOf(1, 0)
                }
                95    // CalcSaveRecalc
                -> data = byteArrayOf(1, 0)
                XLSConstants.PRINTROWCOL -> data = byteArrayOf(0, 0)
                XLSConstants.PRINTGRID -> data = byteArrayOf(0, 0)
                130    // GridSet
                -> data = byteArrayOf(0, 0)
                XLSConstants.GUTS -> data = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0)
                XLSConstants.DEFAULTROWHEIGHT -> data = byteArrayOf(0, 0, 0xFF.toByte(), 0)
                XLSConstants.WSBOOL -> data = byteArrayOf(0xC1.toByte(), 4)
                XLSConstants.HEADERREC, XLSConstants.FOOTERREC -> data = byteArrayOf()
                XLSConstants.HCENTER -> data = byteArrayOf(0, 0)
                XLSConstants.VCENTER -> data = byteArrayOf(0, 0)
                XLSConstants.SETUP -> data = byteArrayOf(0, 0, 0xFF.toByte(), 0, 1, 0, 1, 0, 1, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0xE0.toByte(), 0x3F, 0, 0, 0, 0, 0, 0, 0xE0.toByte(), 0x3F, 0, 0)
                XLSConstants.DEFCOLWIDTH -> data = byteArrayOf(8, 0)
                XLSConstants.DIMENSIONS -> {
                    data = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    record.data = data
                    if (bs != null) { // see if rows have already been added
                        bs.dimensions = record as Dimensions
                        if (bs.rowMap.size > 0) {
                            val z = bs.rows[bs.rowMap.size - 1].rowNumber
                            record.setRowLast(z)
                        }

                        //			z= ((Colinfo)bs.getColinfos().get(bs.getColinfos().size())).getColLast();
                        //			((Dimensions) record).setColLast(z);
                    }
                }
                XLSConstants.WINDOW2 -> data = byteArrayOf(0xB6.toByte(), 6, 0, 0, 0, 0, 0x40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                XLSConstants.EOF -> {
                }
                else -> io.starter.toolkit.Logger.log("Must create required rec: $opcode")
            }


            record.data = data
            try {
                record.init()
            } catch (e: Exception) {
                // TODO: ExtSst needs sst ...
            }

        } catch (e: Exception) {
            Logger.logErr("Record Validation: Error creating missing record: $opcode")
        }

        return record

    }

    /**
     * reset Record members in stream so can be used again for record-level validation
     *
     * @param map
     */
    private fun reInitSubstream(map: LinkedHashMap<Short, R>) {
        val ii = map.keys.iterator()
        while (ii.hasNext()) {
            val r = map[ii.next()]
            r.isPresent = false
            r.recordPos = -1
        }
    }

    /**
     * mark record present and record pertinent information for record-level validation
     */
    private fun markRecord(map: LinkedHashMap<Short, R>, rec: BiffRec?, opcode: Short) {
        try {
            val r = map[opcode]
            if (!r.isPresent) {  // THIS STATEMENT FOR SOME REASON IS VERY VERY VERY TIME CONSUMING -- see testPerformance3
                r.isPresent = true
                r.recordPos = rec!!.recordIndex
            }
        } catch (ne: NullPointerException) {
/*	    if (opcode != CONTINUE && opcode!=DBCELL
		    && opcode < 4000 /* chart records * /
		    && !rec.isValueForCell()) // ignore CELLTABLE records
//		io.starter.toolkit.Logger.log("COULDN'T FIND Opcode: " + opcode);*/

        }

}

/**
 * debug utility
 */
    private fun displayRecsInStream(map:LinkedHashMap<Short, R>) {
val ii = map.keys.iterator()
io.starter.toolkit.Logger.log("Present Records")
while (ii.hasNext())
{
val op = ii.next()
val r = map.get(op)
if (r.isPresent)
{
io.starter.toolkit.Logger.log(op + " at " + r.recordPos)
}
}
}

companion object {

const val serialVersionUID = 1233423412323L

/**
 * Find the location of the next particular opcode
 */
    protected fun getNextOpcodeOffset(op:Short, rec:BlockByteConsumer, parsedata:BlockByteReader):Int {
var found = false
var x = rec.offset
var opcode:Short = 0x0
while (!found && x < parsedata.length - 2)
{
opcode = ByteTools.readShort(parsedata.get(rec, x).toInt(), parsedata.get(rec, ++x).toInt())
if (opcode == op)
{
found = true
break
}
}
if (!found)
return 0
return x - 3
}
}
}

/**
 * represents pertinent info for a BiffRec in an easy-to-access class
 */
internal class R( var isRequired:Boolean) {
 var isPresent:Boolean = false
 var recordPos = -1
 var altPrecedor:ShortArray? = null
}

 /*
 * missing records:
 * notes: 2262 record is ????? see TestInsertRows.testInsertRow0FormulaMovement
 * workingdir + "InsertRowBug1.xls"); TestColumns.testInsertColMoveReferences
 * (this.workingdir + "InsertColumnBug1.xls");
 *
 * TestReadWrite.testNPEOnOpen: workingdir + "equilar/proxycomp3_pwc.xls");
 * COULDN'T FIND Opcode: 194
 *
 * Sheet-level recs:
 * COULDN'T FIND Opcode: 171 ???  sheet substream just before 153 then EOF -- see testRecalc
   COULDN'T FIND Opcode: 148 -->TestNPEOnOpen
 *
 *
 *
 *
 * testCorruption.testOutOfSpec -- missing required records: workingdir +
 * "Caribou_North_And_South.xls" A required record is not present: 225 A
 * required record is not present: 193 A required record is not present: 226 A
 * required record is not present: 92 A required record is not present: 25 A
 * required record is not present: 19 A required record is not present: 431 A
 * required record is not present: 444
 *
 * TestFormulaCalculator.testRecalc: workingdir + "Sakonnet/smile.xls" A
 * required record is not present: 225 A required record is not present: 193 A
 * required record is not present: 226 A required record is not present: 353 A
 * required record is not present: 317 A required record is not present: 18 A
 * required record is not present: 431 A required record is not present: 444 A
 * required record is not present: 64 A required record is not present: 141 A
 * required record is not present: 439 A required record is not present: 218 A
 * required record is not present: 352 A required record is not present: 255
 */

 /*
 * record order issues: 659= Style 352= UsesELFs 140= Country 146= Palette
 *
 * 2173= XfExt 2189= DXF 2190= TableStyles 2211= ForceFullCalculation
 *
 *
 * testValidationHandle.testEquilarFile workingdir +
 * "equilar/proxycomp3_formatted_new.xls" Record out of order: r:659/682
 * lastR:2189/903 Record out of order: r:352/732 lastR:2190/956 Record out of
 * order: r:140/753 lastR:2211/958
 *
 * TestScenarios.borderbrokenxx, formularesultxx, numericerrorxx
 * TestFormulaParse.testRajeshOOM: workingdir + "OOXML/proxycomp3_cap_new.xls"
 * Record out of order: r:659/1304 lastR:2189/1557 Record out of order:
 * r:146/1375 lastR:2190/1610 Record out of order: r:140/1403 lastR:2211/1612
 *
 * TestAutoFilter.XXX: workingdir + "testAutoFilter.xls" Record out of order:
 * r:659/128 lastR:2173/204 Record out of order: r:352/175 lastR:2190/292 Record
 * out of order: r:140/178 lastR:2211/294
 *
 * also see TestAddAF w.r.t. insertion position -- supbook is wrong? ...
 *
 * TestMemoryUsage.testMemUsageWithManyReferences: workingdir +
 * "equilar/proxycomp3_pmp_Model_T_new.xls" Record out of order: r:659/1202
 * lastR:2189/1448 Record out of order: r:146/1273 lastR:2190/1501 Record out of
 * order: r:140/1297 lastR:2211/1503 Record out of order: r:89/1300
 * lastR:35/1308
 *
 * TestInsertRows.testInsertRowsWithFormulas.this.workingsheetsterdir +
 * "aviasphere/AirframeMaster_F20Import.xls" Record out of order: r:659/274
 * lastR:2173/352 Record out of order: r:352/323 lastR:2190/537 Record out of
 * order: r:140/330 lastR:2211/539
 *
 * TestXMLSS.testClasicXMLReadWriteError: workingdir + "xmlsavecorruption.xls"
 * Record out of order: r:659/144 lastR:2173/208 Record out of order: r:352/194
 * lastR:2190/364 Record out of order: r:140/196 lastR:2211/366
 *
 * TestRowCols.testRowInsertionFormatLoss.workingdir + "claritas/input/400.xls"
 * Record out of order: r:659/230 lastR:2189/426 Record out of order: r:352/280
 * lastR:2190/481 Record out of order: r:140/288 lastR:2211/483
 *
 * TEstRowCols.testRowFormats:testRowFormats.xls Record out of order: r:659/278
 * lastR:2173/356 Record out of order: r:352/327 lastR:2190/541 Record out of
 * order: r:140/334 lastR:2211/543
 *
 * TestRowCols.testAutoAjdustRowHeight.C:\eclipse\workspace\Testfiles\OpenXLS\input
 * \equilar/tcr_formatted_2003.xls Record out of order: r:659/709 lastR:2189/973
 * Record out of order: r:352/774 lastR:2190/1038 Record out of order: r:140/800
 * lastR:2211/1040
 */