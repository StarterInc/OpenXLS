/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.formats.XLS.charts

import io.starter.formats.XLS.*
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import java.util.Stack

/**
 * **Ai: Linked Chart Data (1051h)**<br></br>
 *
 *
 * This record specifies linked series data or text.
 *
 *
 * Offset           Name    Size    Contents
 * --
 * 4               id      1       index id: (0=title or text, 1=series vals, 2=series cats + (3=bubbles)
 * 5               rt      1       reference type(0=default,1=text in formula bar, 2=worksheet, 4=error)
 * 6               grbit   2       flags
 * 8               ifmt    2       Index to number format (used if fCustomIfmt= true)
 * 10              cce     2       size of rgce (in bytes)
 * 12              rgce    var     Parsed formula of link
 *
 *
 * The grbit field contains the following option flags.
 *
 *
 * 0	0	0x1		fCustomIfmt		TRUE if this object has a custom number format; FALSE if number format is linked to data source
 * a	1	0x2		(reserved)		Reserved; must be zero
 * 0	5-2	0x3C	st				Source type (always zero)
 * a	7-6	0xCO	(reserved)		Reserved; must be zero
 * 1	7-0	0xFF	(reserved)		Reserved; must be zero
 *
 *
 *
 * @see Formula
 *
 * @see Chart
 */

class Ai : GenericChartObject(), ChartObject {
    private var expression: Stack<Ptg>? = null
    /**
     * Return the type (ID) of this Ai record
     *
     * @return int rt
     */
    var type = -1
        protected set
    /**
     * return the custom number format for this AI (category, series or bubble)
     * <br></br>if 0, use default number format as specified by Axis or Chart
     *
     * @return
     */
    var ifmt = -1
        protected set
    protected var cce = -1
    private var grbit: Short = -1
    private var rt: Short = -1
    private var fCustomIfmt = false
    // private CompatibleVector xlsrecs = new
    // CompatibleVector();
    private var st: SeriesText? = null

    /**
     * This is for storage of the boundsheet name when the ai record is being moved from
     * one workbook to another.  In these cases, populate this value, and pull it back out
     * to locate the cxti and reset the reerences
     *
     * @return the name of the original boundsheet this record is associated with.
     */
    /**
     * returns the bound sheet name - must be called after populateForTransfer
     *
     * @return String
     */ // 20080116 KSC: Changed name for clarity
    var boundName: String? = ""
        private set
    private var origSheetName = ""
    /**
     * returns the sheet reference index - must call after populateForTransfer
     *
     * @return
     */
    var boundXti = -1
        private set

    // XLSRecord x = getRecord(0);
    // if(x instanceof SeriesText)return ((SeriesText)x
    // ).toString();
    // TODO: figure out why this doesn't find the title -- see
    // "reportS01Template.xls"
    val text: String
        get() {
            try {
                if (st != null)
                    return st!!.toString()
            } catch (e: Exception) {
                if (DEBUGLEVEL > 0)
                    Logger.logWarn("Error getting Chart String value: $e")
                return definition
            }

            return "undefined"
        }

    /**
     * get the display name
     */
    internal val name: String
        get() = "Chart Ai"

    /**
     * get the definition text
     * the definition is stored
     * in Excel parsed format
     */
    // 20091019 KSC: if complex
    // series, will have a PtgParen
    // after PtgMemFunc;
    /*
                 * if (ep[t] instanceof PtgMemFunc)
                 * sb.append("(" + ep[t].getString() + ")");
                 * else
                 */ val definition: String
        get() {
            val sb = StringBuffer()
            var ep = arrayOfNulls<Ptg>(expression!!.size)
            ep = expression!!.toTypedArray<Ptg>()
            for (t in ep.indices) {
                if (ep[t] !is PtgParen)
                    sb.append(ep[t].string)
            }
            return sb.toString()
        }

    /*
     * Returns an array of ptgs that represent any BiffRec
     * ranges in the formula.
     * Ranges can either be in the format "C5" or "Sheet1!C4:D9"
     */
    val cellRangePtgs: Array<Ptg>
        @Throws(FormulaNotFoundException::class)
        get() = ExpressionParser.getCellRangePtgs(expression!!)

    /**
     * Sets the boundsheet name for the data referenced in this AI.
     * + sets the xti of the boundsheet reference (necessary upon sheet copy/move)
     * Does not currently support multi-boundsheet references.
     *
     * @see updateSheetRef
     */
    fun populateForTransfer(origSheetName: String) {
        if ("" == boundName) {
            this.origSheetName = origSheetName // 20080708 KSC: trap original
            // sheet name for
            // updateSheetRefs comparison
            for (t in expression!!.indices) {
                val p = expression!![t]
                if (p is PtgArea3d) {
                    try {
                        boundName = p.sheetName // original boundname +
                        // xti sheet ref
                        boundXti = p.getIxti().toInt()
                    } catch (e: Exception) {
                        Logger.logErr("Ai.populateForTransfer: Chart contains links to other data sources")
                    }

                } else if (p is PtgRef3d) {
                    try {
                        boundName = p.sheetName // original boundname +
                        // xti sheet ref
                        boundXti = p.getIxti().toInt()
                    } catch (e: Exception) {
                        Logger.logErr("Ai.populateForTransfer: Chart contains links to other data sources")
                    }

                } else if (p is PtgMemFunc) { // 20091015 KSC: for
                    // non-contiguous ranges
                    val ptg = p.firstloc
                    if (ptg is PtgRef3d) {
                        try {
                            boundName = ptg.sheetName // original boundname
                            // + xti sheet ref
                            boundXti = ptg.getIxti().toInt()
                        } catch (e: Exception) {
                            Logger.logErr("Ai.populateForTransfer: Chart contains links to other data sources")
                        }

                    } else { // should be a PtgArea3d
                        val pr = ptg as PtgArea3d
                        try {
                            boundName = pr.sheetName // original boundname
                            // + xti sheet ref
                            boundXti = pr.getIxti().toInt()
                        } catch (e: Exception) {
                            Logger.logErr("Ai.populateForTransfer: Chart contains links to other data sources")
                        }

                    }
                }
            }
        }
    }

    fun setLegend(newLegend: String) {
        this.setRt(1)
        st!!.setText(newLegend)
        expression = null // if setting to a string, no expression!
    }

    /**
     * Refer to the associated Series Text.
     *
     * @param txt
     * @return
     */
    fun setText(txt: String): Boolean {
        try {
            // XLSRecord x = getRecord(0);
            // ((SeriesText)x ).setText(txt);
            st!!.setText(txt)
            return true
        } catch (e: ClassCastException) {
            // Logger.logInfo("Error getting Chart String value: " + e);
            return false
        }

    }

    override fun toString(): String {
        when (type) {
            Ai.TYPE_TEXT -> return text
            Ai.TYPE_VALS -> return definition
            Ai.TYPE_CATEGORIES -> return definition
            Ai.TYPE_BUBBLES -> return definition
        }
        return super.toString()
    }

    // public void addRecord(BiffRec rec){
    // xlsrecs.add(rec);
    // }
    fun setSeriesText(s: SeriesText) {
        st = s
    }

    // protected XLSRecord getRecord(int i){
    // return (XLSRecord) xlsrecs.get(i);
    // }

    /**
     * set the Externsheet reference
     * for any associated PtgArea3d's
     */
    @Throws(WorkSheetNotFoundException::class)
    fun setExternsheetRef(x: Int) {
        val dt = this.data
        val pos = 8

        for (t in expression!!.indices) {
            val p = expression!![t]
            if (p is PtgArea3d) {
                p.setIxti(x.toShort())
                p.addToRefTracker()
                if (DEBUGLEVEL > 3)
                    Logger.logInfo("Setting sheet reference for: "
                            + p.toString() + "  in Ai record.")
                // register with the Externsheet reference
                this.workBook!!.externSheet!!.addPtgListener(p)
                updateRecord()
            } else if (p is PtgRef3d) { // 20091015 KSC: Added
                p.setIxti(x.toShort())
                p.addToRefTracker()
                if (DEBUGLEVEL > 3)
                    Logger.logInfo("Setting sheet reference for: "
                            + p.toString() + "  in Ai record.")
                // register with the Externsheet reference
                this.workBook!!.externSheet!!.addPtgListener(p)
                updateRecord()
            } else if (p is PtgMemFunc) { // 20091015 KSC: Added
                val pr = p.firstloc
                if (pr is PtgRef3d) {
                    pr.setIxti(x.toShort())
                    pr.addToRefTracker()
                    this.workBook!!.externSheet!!
                            .addPtgListener(pr)
                } else { // should be a PtgArea3d
                    (pr as PtgArea3d).setIxti(x.toShort())
                    pr.addToRefTracker()
                    this.workBook!!.externSheet!!
                            .addPtgListener(pr)
                }
                if (DEBUGLEVEL > 3)
                    Logger.logInfo("Setting sheet reference for: "
                            + pr.toString() + "  in Ai record.")
                // register with the Externsheet reference
                updateRecord()
            } else {
                Logger.logInfo("Ai.setExternsheetRef: unknown Ptg")
            }
        }
    }

    /**
     * set the Externsheet reference
     * for any associated PtgArea3d's that match the old reference.
     *
     *
     * invaluble for modifying only one sheets worth of references (ie a move sheet situation)
     */
    @Throws(WorkSheetNotFoundException::class)
    fun setExternsheetRef(oldRef: Int, newRef: Int) {
        for (t in expression!!.indices) {
            val p = expression!![t]
            if (p is PtgArea3d) {
                val oRef = p.getIxti().toInt()
                if (oRef == oldRef) { // got the one to update
                    p.removeFromRefTracker() // 20100506 KSC: added
                    p.sheetName = this.sheet!!.sheetName // 20100415
                    // KSC:
                    // added
                    p.setIxti(newRef.toShort())
                    p.addToRefTracker() // 20080709 KSC
                    if (DEBUGLEVEL > 3)
                        Logger.logInfo("Setting sheet reference for: "
                                + p.toString() + "  in Ai record.")
                    // register with the Externsheet reference
                    this.workBook!!.externSheet!!.addPtgListener(p)
                    updateRecord()
                }
            } else if (p is PtgRef3d) {
                val oRef = p.getIxti().toInt()
                if (oRef == oldRef) {
                    p.removeFromRefTracker()
                    p.sheetName = this.sheet!!.sheetName // 20100415
                    // KSC:
                    // added
                    p.setIxti(newRef.toShort())
                    if (!p.isRefErr)
                        p.addToRefTracker()
                    if (DEBUGLEVEL > 3)
                        Logger.logInfo("Setting sheet reference for: "
                                + p.toString() + "  in Ai record.")
                    // register with the Externsheet reference
                    this.workBook!!.externSheet!!.addPtgListener(p)
                    updateRecord()
                }
            } else if (p is PtgMemFunc) { // 20091015 KSC: Added
                val pr = p.firstloc
                if (pr is PtgRef3d) {
                    val oRef = pr.getIxti().toInt()
                    if (oRef == oldRef) {
                        pr.removeFromRefTracker() // 20100506 KSC:
                        // added
                        (pr as PtgArea3d)
                                .sheetName = this.sheet!!.sheetName // 20100415
                        // KSC:
                        // added
                        (pr as PtgRef3d).setIxti(newRef.toShort())
                        (pr as PtgRef3d).addToRefTracker()
                        this.workBook!!.externSheet!!
                                .addPtgListener(pr as PtgRef3d)
                    }
                } else { // should be a PtgArea3d
                    val oRef = (pr as PtgRef3d).getIxti().toInt()
                    if (oRef == oldRef) {
                        pr.removeFromRefTracker() // 20100506 KSC:
                        // added
                        (pr as PtgArea3d)
                                .sheetName = this.sheet!!.sheetName // 20100415
                        // KSC:
                        // added
                        (pr as PtgArea3d).setIxti(newRef.toShort())
                        (pr as PtgArea3d).addToRefTracker()
                        this.workBook!!.externSheet!!
                                .addPtgListener(pr as PtgArea3d)
                    }
                }
                if (DEBUGLEVEL > 3)
                    Logger.logInfo("Setting sheet reference for: "
                            + pr.toString() + "  in Ai record.")
                // register with the Externsheet reference
                updateRecord()
            } else if (p is PtgMystery) {
                // TODO: do what???
            }
        }
    }

    /**
     * take the original boundSheet + boundSheet xti reference and update it to
     * the sheet reference in the new workbook (or same workbook but different ixti reference)
     */
    fun updateSheetRef(newSheetName: String, origWorkBookName: String) {
        try {
            // 20080630/20080708 KSC: Fixes for [BugTracker 1799] +
            // [BugTracker 1434]
            if (boundXti > -1) { // has populate for transfer been called -
                // should!
                var newSheetNum = -1
                try {
                    if (!boundName!!.equals(origSheetName, ignoreCase = true)) { // Ai
                        // reference
                        // is on
                        // a
                        // dfferent
                        // sheet,
                        // see
                        // if it
                        // exists
                        // in
                        // new
                        // workbook
                        newSheetNum = this.workBook!!
                                .getWorkSheetByName(boundName).sheetNum
                    } else
                    // Ai reference is on same sheet, point now to new
                    // sheet
                        newSheetNum = this.workBook!!
                                .getWorkSheetByName(newSheetName).sheetNum
                } catch (e: Exception) { // 20080123 KSC: if links arent there,
                    // fix == try to make an External ref
                    for (t in expression!!.indices) {
                        if (expression!![t] is PtgArea3d) {
                            val p = expression!![t] as PtgArea3d
                            Logger.logWarn("External References are unsupported: External reference found in Chart: " + p.sheetName!!)
                            p.sheetName = boundName // set external reference
                            // to original
                            // boundsheet name
                            p.setExternalReference(origWorkBookName)
                            this.setExternsheetRef(p.getIxti().toInt())
                        } else {
                            Logger.logInfo("Ai.updateSheetRef:")
                        }
                    }
                }

                if (newSheetNum != -1) {
                    this.setSheet(this.workBook!!
                            .getWorkSheetByName(newSheetName)) // 20100415 KSC:
                    // set Ai sheet
                    // ref to new
                    // sheet
                    val xsht = this.workBook!!.getExternSheet(true) // create
                    // if
                    // necessary
                    val newXRef = xsht!!.insertLocation(newSheetNum, newSheetNum)
                    this.setExternsheetRef(boundXti, newXRef)
                    boundXti = newXRef // 20100506 KSC: reset
                    boundName = newSheetName // ""
                }
            } else {// debugging 20100415
                // Logger.logErr("Ai.updateSheetRef: boundxti is -1 for AI "
                // + this.toString());
            }
        } catch (e: Exception) {
            Logger.logErr("Ai.updateSheetRef: $e")
        }

    }

    /*
     * Returns the ptg that matches the string location sent to
     * it.
     * this can either be in the format "C5" or a range, such as
     * "C4:D9"
     */
    fun getPtgsByLocation(loc: String): List<Ptg>? {
        try {
            return ExpressionParser.getPtgsByLocation(loc, expression!!)
        } catch (e: FormulaNotFoundException) {
            Logger.logWarn("failed to update Chart Series Location: $e")
        }

        return null
    }

    /**
     * locks the Ptg at the specified location
     */

    fun setLocationPolicy(loc: String, l: Int): Boolean {
        val dx = this.getPtgsByLocation(loc)
        val lx = dx!!.iterator()
        while (lx.hasNext()) {
            val d = lx.next()
            if (d != null) {
                d.locationPolicy = l
            }
        }
        return true
    }

    /**
     * simplified version of changeAiLocation which locates current Ptg and updates expression
     * NOTE: newLoc is expected to be a valid reference
     *
     * @param p
     * @param newLoc
     * @return
     */
    fun changeAiLocation(p: Ptg, newLoc: String): Boolean {
        val aiLocs = StringTool.splitString(newLoc, ",")
        for (i in aiLocs.indices) {
            try { // NOTE: Ai has only 1 expression OR 2: 1= PtgMemFunc
                // 2=PtgParen
                // looking up via getExpressionLocByPtg errors if loc is the
                // same but Ptg's are different objects eg. PtgRef3d vs
                // PtgArea3d ...
                if (expression!![0] is PtgMemFunc) {
                    try { // must find particular ptg in PtgMemFunc's
                        // subexression to reset
                        // APPARENTLY WE DO NOT NEED TO UDPATE PTGMEMFUNC
                        // SUBEXPRESSION EXPLICITLY
                        // int z= ExpressionParser.getExpressionLocByPtg(p,
                        // ((PtgMemFunc)
                        // expression.get(0)).getSubExpression());
                        // Stack subexp= ((PtgMemFunc)
                        // expression.get(0)).getSubExpression();
                        // for (int z= 0; z < subexp.size(); z++) {
                        // if (p.equals(subexp.get(z))) {
                        p.location = aiLocs[i] // updates ref. tracker
                        // ((PtgMemFunc)
                        // expression.get(0)).getSubExpression().set(z, p); //
                        // update expression with new Ptg
                        // }
                        // }
                    } catch (ex: Exception) {
                        Logger.logErr("Ai.changeAiLocation: Error updating Location in non-contiguous range $ex")
                    }

                } else {
                    p.location = aiLocs[i] // updates ref. tracker
                    expression!![0] = p // update expression with new Ptg
                    if (this.type == Ai.TYPE_TEXT) { // must reset text for
                        // SeriesText as
                        // well
                        try {
                            val o = p.value
                            this.setText(o.toString())
                        } catch (e: Exception) {
                        }

                    }
                }
            } catch (e: Exception) {
                Logger.logErr("Ai.changeAiLocation: Error updating Location to "
                        + newLoc + ":" + e.toString())
                return false
            }

        }
        updateRecord()
        return true
    }

    /**
     * Takes a string as a current formula location, and changes
     * that pointer in the formula to the new string that is sent.
     * This can take single cells "A5" and cell ranges, "A3:d4"
     * Returns true if the cell range specified in formulaLoc exists & can be changed
     * else false.  This also cannot change a cell pointer to a cell range or vice
     * versa.
     *
     * @param String - range of Cells within Formula to modify
     * @param String - new range of Cells within Formula
     */
    fun changeAiLocation(loc: String, newLoc: String): Boolean {
        // TODO: Implement formula policy!! -jm
        var ptg: Ptg? = null
        var z = -1
        try {
            if (expression!!.size > 0) {
                z = ExpressionParser
                        .getExpressionLocByLocation(loc, expression!!)
                ptg = expression!![z] // 20090917 KSC: since creating
                // new ptgs below, must remove
                // original from reftracker "by
                // hand"
            }
            if (ptg != null)
                (ptg as PtgRef).removeFromRefTracker()
        } catch (e: Exception) {
        }

        if (z == -1 && newLoc == "") {// no reference -- happens on
            // legends, category ai's ...
            this.data[1] = 1 // text reference rather than worksheet
            // reference
            return false
        }
        ptg = PtgRef.createPtgRefFromString(newLoc, this)
        if (z != -1)
        // then must change original
            expression!![z] = ptg // update expression with new Ptg
        else
            expression!!.add(ptg)
        updateRecord()
        return true
    }

    /*
     * Update the record byte array with the modified ptg
     * records
     */
    fun updateRecord() {
        var offy = 8 // the start of the parsed expression
        val rkdata = this.data
        var updated = ByteArray(rkdata!!.size)
        System.arraycopy(rkdata, 0, updated, 0, offy)
        for (i in expression!!.indices) {
            val o = expression!!.elementAt(i)
            val ptg = o as Ptg
            val b = ptg.record
            // must inc. size if Ptgs have inc.'d ... see
            // changeAiLocation
            val len = b.size
            if (updated.size - offy < len) {
                val newArr = ByteArray(offy + len)
                System.arraycopy(updated, 0, newArr, 0, updated.size)
                // update cce in array as well ...
                cce += newArr.size - updated.size
                updated = newArr
                val ix = ByteTools.shortToLEBytes(cce.toShort())
                System.arraycopy(ix, 0, updated, 6, 2)
            }
            System.arraycopy(b, 0, updated, offy, len)
            offy = offy + len
        }
        this.data = updated
    }

    override fun init() {
        super.init()
        type = this.getByteAt(0).toInt()
        // index id: (0=title or text, 1=series vals, 2=series cats,
        // 3= bubbles
        rt = this.getByteAt(1).toShort()
        // reference type(0=default,1=text in formula bar,
        // 2=worksheet, 4=error)
        grbit = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        setfCustomIfmt(grbit and 0x1 == 0x1)
        // flags
        ifmt = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()
        // Index to number format
        cce = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt()).toInt()
        // size of rgce (in bytes)
        var pos = 8
        // Parsed formula of link

        // get the parsed expression
        val expressionbytes = this.getBytesAt(pos, cce)
        expression = ExpressionParser.parseExpression(expressionbytes, this)
        pos += cce
        if (DEBUGLEVEL > 10)
            Logger.logInfo(this.name + ":" + this.definition)
    }

    /**
     * set reference type(0=default,1=text in formula bar, 2=worksheet, 4=error)
     *
     * @param i
     */
    fun setRt(i: Int) {
        rt = i.toShort()
        this.data[1] = rt.toByte()
    }

    fun getExpression(): Stack<*>? {
        return expression
    }

    override fun close() {
        if (expression != null) {
            while (!expression!!.isEmpty()) {
                var p: GenericPtg? = expression!!.pop() as GenericPtg
                if (p is PtgRef)
                    p.close()
                else
                    p!!.close()
                p = null
            }
        }
        super.close()
    }

    protected fun finalize() {
        this.close()
    }

    /**
     * @return the fCustomIfmt
     */
    fun isfCustomIfmt(): Boolean {
        return fCustomIfmt
    }

    /**
     * @param fCustomIfmt the fCustomIfmt to set
     */
    fun setfCustomIfmt(fCustomIfmt: Boolean) {
        this.fCustomIfmt = fCustomIfmt
    }

    companion object {
        /**
         *
         */
        private val serialVersionUID = -6647823755603289012L
        // define the type of Ai record
        val TYPE_TEXT = 0
        val TYPE_VALS = 1
        val TYPE_CATEGORIES = 2
        val TYPE_BUBBLES = 3

        /**
         * Get a prototype with the specified ai types.
         *
         *
         * 0 = Legend AI
         * 1=  Series Value Ai
         * 2 = Category Ai
         * 3 = Unknown, undocumented, but neccesarry AI
         * 4 = Blank Legend AI with no reference.
         */
        fun getPrototype(aiType: ByteArray): ChartObject {
            val ai = Ai()
            ai.opcode = XLSConstants.AI
            ai.data = aiType
            ai.init()
            return ai
        }

        // 20070801 KSC: since changeAiLocation now allows addition
        // of new expression bytes, alter
        // default prototype bytes here to not include any
        // expression bytes ..
        var AI_TYPE_LEGEND = byteArrayOf(0, 2, 0, 0, 0, 0, 0, 0)                                                                // ,
        // 7,
        // 0,
        // 58,
        // 0,
        // 0,
        // 0,
        // 0,
        // 0,
        // 0};
        var AI_TYPE_SERIES = byteArrayOf(1, 2, 0, 0, 0, 0, 0, 0)                                                                // ,
        // 11,
        // 0,
        // 59,
        // 0,
        // 0,
        // 1,
        // 0,
        // 1,
        // 0,
        // 1,
        // 0,
        // 3,
        // 0};
        var AI_TYPE_CATEGORY = byteArrayOf(2, 2, 0, 0, 0, 0, 0, 0)                                                                // 11,
        // 0,
        // 59,
        // 0,
        // 0,
        // 0,
        // 0,
        // 0,
        // 0,
        // 1,
        // 0,
        // 3,
        // 0};
        var AI_TYPE_BUBBLE = byteArrayOf(3, 1, 0, 0, 0, 0, 0, 0)
        var AI_TYPE_NULL_LEGEND = byteArrayOf(0, 1, 0, 0, 0, 0, 0, 0)
    }
}