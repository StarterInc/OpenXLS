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
package io.starter.formats.XLS.formulas

import io.starter.OpenXLS.CellHandle
import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList


/**
 * ptgArea3d is a reference to an area (rectangle) of cells.
 * Essentially it is a collection of two ptgRef's, so it will be
 * treated that way in the code...
 * implies external sheet ref (rather than ptgarea)
 *
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           ixti        2       index into the Externsheet record
 * 2           rwFirst     2       The First row of the reference
 * 4           rwLast      2       The Last row of the reference
 * 6           grbitColFirst   2       (see following table)
 * 8           grbitColLast    2       (see following table)
 *
 * Only the low-order 14 bits specify the Col, the other bits specify
 * relative vs absolute for both the col or the row.
 *
 * Bits        Mask        Name    Contents
 * -----------------------------------------------------
 * 15          8000h       fRwRel  =1 if row offset relative,
 * =0 if otherwise
 * 14          4000h       fColRel =1 if row offset relative,
 * =0 if otherwise
 * 13-0        3FFFh       col     Ordinal column offset or number
</pre> *
 *
 *
 * For 3D references, the tokens contain a negative EXTERNSHEET index, indicating a reference into the own workbook.
 * The absolute value is the one-based index of the EXTERNSHEET record that contains the name of the first sheet. The
 * tokens additionally contain absolute indexes of the first and last referenced sheet. These indexes are independent of the
 * EXTERNSHEET record list. If the referenced sheets do not exist anymore, these indexes contain the value FFFFH (3D
 * reference to a deleted sheet), and an EXTERNSHEET record with the special name <04H> (own document) is used.
 * Each external reference contains the positive one-based index to an EXTERNSHEET record containing the URL of the
 * external document and the name of the sheet used. The sheet index fields of the tokens are not used.
 *
 *
 * is the above correct?? Documentation sez different!!!!!
 *
 * @see Ptg
 *
 * @see Formula
 */
open class PtgArea3d
/**
 * create new PtgArea3d
 */
() : PtgArea(), Ptg, IxtiListener {

    override val isOperand: Boolean
        get() = true

    override val isReference: Boolean
        get() = true

    internal var quoted = false
    var ixti: Short = 0
    /**
     * return true if this PtgArea3d is an external reference
     * i.e. defined in another, external workbook
     *
     * @return
     */
    var isExternalRef = false
        private set    // true if this ptg area is a reference in another workbook
    private var comps: Array<Ptg>? = null

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    override// handle non-standard ranges i.e. $B:$C or $1:$3
    // sheet, addr
    // sheet, addr
    // otherwise,
    // is ok if new ptg...
    val string: String?
        get() {
            try {
                if (this.isWholeCol || this.isWholeRow) {
                    val s = firstPtg!!.location
                    val y = lastPtg!!.location

                    val loc1 = ExcelTools.stripSheetNameFromRange(s!!)
                    val loc2 = ExcelTools.stripSheetNameFromRange(y!!)
                    if (this.isWholeCol) {
                        var i = loc1[1].length
                        if (Character.isDigit(loc1[1][i - 1]))
                            while (Character.isDigit(loc1[1][--i]));
                        loc1[1] = loc1[1].substring(0, i)
                        i = loc2[1].length
                        if (Character.isDigit(loc2[1][i - 1]))
                            while (Character.isDigit(loc2[1][--i]));
                        loc2[1] = loc2[1].substring(0, i)
                    } else if (this.isWholeRow) {
                        var i = 0
                        while (!Character.isDigit(loc1[1][i++]));
                        loc1[1] = "$" + loc1[1].substring(i - 1)
                        i = 0
                        while (!Character.isDigit(loc2[1][i++]));
                        loc2[1] = "$" + loc2[1].substring(i - 1)
                    }
                    sheetname = GenericPtg.qualifySheetname(sheetname)

                    return sheetname + "!" + loc1[1] + ":" + loc2[1]
                }
                return location
            } catch (e: Exception) {
                return null
            }

        }

    /**
     * return the first sheet referenced
     *
     * @return
     */
    // fall thru -- see sheet copy operations -- appears correct
    val firstSheet: Boundsheet?
        get() {
            if (parentRec != null) {
                val wb = parentRec.workBook
                if (sheetname != null)
                    try {
                        return wb!!.getWorkSheetByName(sheetname)
                    } catch (e: WorkSheetNotFoundException) {
                    }

                if (wb != null && wb.externSheet != null) {
                    val bsa = wb.externSheet!!.getBoundSheets(this.ixti.toInt())
                    if (bsa != null)
                        return bsa[0]
                }
            }
            return null
        }

    /**
     * return the last sheet referenced
     *
     * @return
     */
    val lastSheet: Boundsheet?
        get() {
            if (parentRec != null) {
                val wb = parentRec.workBook
                if (wb != null && wb.externSheet != null) {
                    val bsa = wb.externSheet!!.getBoundSheets(this.ixti.toInt())
                    if (bsa != null) {
                        return if (bsa.size > 1) bsa[bsa.size - 1] else bsa[0]
                    }

                }
            }
            return null
        }

    override var parentRec: XLSRecord
        get
        set(rec) {
            super.parentRec = rec
            if (firstPtg != null)
                firstPtg!!.parentRec = parentRec
            if (lastPtg != null)
                lastPtg!!.parentRec = parentRec
        }

    /**
     * return the first sheet referenced
     *
     * @return
     */
    val sheet: Boundsheet?
        get() = firstSheet

    /**
     * get the sheet name from the 1st 3d reference
     */
    override// try this:
    var sheetName: String?
        get() {
            if (sheetname == null) {
                if (parentRec != null) {
                    val wb = parentRec.workBook
                    if (wb != null && wb.externSheet != null) {
                        val sheets = wb.externSheet!!.getBoundSheetNames(this.ixti.toInt())
                        if (sheets != null && sheets.size > 0) {
                            sheetname = sheets[0]
                            sheetname = GenericPtg.qualifySheetname(sheetname)
                        }
                    }

                    if (sheetname == null && parentRec != null && parentRec.sheet != null) {
                        sheetname = parentRec.sheet!!.sheetName
                        sheetname = GenericPtg.qualifySheetname(sheetname)
                    }
                }
            }
            return sheetname
        }
        set(value: String?) {
            super.sheetName = value
        }

    /**
     * return the name of the last sheet referenced if it's an external ref
     *
     * @return
     */
    // 20100217 KSC: default to 1st sheet
    val lastSheetName: String?
        get() {
            var sheetname = this.sheetname
            if (parentRec != null) {
                val wb = parentRec.workBook
                if (wb != null && wb.externSheet != null) {
                    val sheets = wb.externSheet!!.getBoundSheetNames(this.ixti.toInt())
                    if (sheets != null && sheets.size > 0)
                        sheetname = sheets[sheets.size - 1]
                }
            }
            return sheetname
        }

    /**
     * Set the location of this PtgRef.  This takes a location
     * such as "a14:b15"
     *
     *
     * NOTE: the reference stays on the same sheet!
     */
    override var location: String
        get
        set(address) {
            val s = ExcelTools.stripSheetNameFromRange(address)
            setLocation(s)
        }


    /**
     * returns the location of the ptg as an array of shorts.
     * [0] = firstrow
     * [1] = firstcol
     * [2] = lastrow
     * [3] = lastcol
     */
    override val intLocation: IntArray?
        get() {
            val first = firstPtg!!.intLocation
            val last = lastPtg!!.intLocation
            val returning = IntArray(4)
            System.arraycopy(first!!, 0, returning, 0, 2)
            System.arraycopy(last!!, 0, returning, 2, 2)
            return returning
        }

    /**
     * return all of the Ptg values represented in this array
     *
     *
     * will have to reference the workbook cells as well as
     * any upstream formulas...
     *
     * @return
     */
    val allVals: Array<Any>?
        get() = null

    override val length: Int
        get() = Ptg.PTG_AREA3D_LENGTH


    /**
     * Returns all of the cells of this range as PtgRef's.
     * This includes empty cells, values, formulas, etc.
     * Note the setting of parent-rec requires finding the cell
     * the PtgRef refer's to.  If that is null then the PtgRef
     * will exist, just with a null value.  This could cause issues when
     * programatically populating cells.
     */
    override// loop through the cols
    // normal case
    //				 TODO: check rc sanity here
    // Get Actual Coordinates
    // Get Actual Coordinates
    // like $1:$1
    // can happens when Name record is being init'd and sheet records are not set yet
    // like $J:$J
    // Get Actual Coordinates
    // can happens when Name record is being init'd and sheet records are not set yet
    // loop through the rows inside
    // cache these suckers!
    // must set parentrec before setLocation
    val components: Array<Ptg>?
        get() {
            if (comps != null)
                return comps

            val components = ArrayList<Ptg>()
            try {
                var sht: String? = ""
                if (this.toString()!!.indexOf("!") > -1) {
                    sht = this.toString()
                    sht = sht!!.substring(0, sht.indexOf("!")) + "!"
                }
                var startrow = 0
                var startcol = 0
                var endrow = 0
                var endcol = 0
                if (!this.isWholeCol && !this.isWholeRow) {
                    val startloc = firstPtg!!.realIntLocation
                    startcol = startloc[1]
                    startrow = startloc[0]
                    val endloc = lastPtg!!.realIntLocation
                    endcol = endloc[1]
                    endrow = endloc[0]
                } else if (this.isWholeRow) {
                    startcol = 0
                    try {
                        endcol = this.sheet!!.maxCol
                    } catch (ne: NullPointerException) {
                        return null
                    }

                    endrow = firstPtg!!.rw
                    startrow = endrow
                } else if (this.isWholeCol) {
                    startrow = 0
                    endcol = firstPtg!!.col
                    startcol = endcol
                    try {
                        endrow = this.sheet!!.maxRow
                    } catch (ne: NullPointerException) {
                        return null
                    }

                }
                while (startcol <= endcol) {
                    var rowholder = startrow
                    while (rowholder <= endrow) {
                        val displaycol = ExcelTools.getAlphaVal(startcol)
                        val displayrow = rowholder + 1
                        val loc = sht + displaycol + displayrow
                        val pref = PtgRef3d(false)
                        pref.parentRec = parentRec
                        pref.location = loc
                        components.add(pref)
                        rowholder++
                    }
                    startcol++
                }
            } catch (e: Exception) {
                Logger.logErr("calculating range value in PtgArea3d failed.", e)
            }

            val pref = arrayOfNulls<PtgRef>(components.size)
            components.toTypedArray<PtgRef>()
            comps = pref
            return comps
        }


    /**
     * sets the column to be relative (relative is true) or absolute (relative is false)
     * <br></br>absolute references do not shift upon column inserts or deletes
     * <br></br>NOTE: DOES NOT handle asymmetrical ranges i.e. 1st ref is absolute, 2nd is relative
     *
     * @param boolean relative
     */
    override var isColRel: Boolean
        get() = super.isColRel
        set(relative) {
            this.fColRel = relative
            firstPtg!!.isColRel = relative
            lastPtg!!.isColRel = relative
            updateRecord()
        }

    /**
     * sets the row to be relative (relative is true) or absolute (relative is false)
     * <br></br>absolute references do not shift upon row inserts or deletes
     * <br></br>NOTE: DOES NOT handle asymmetrical ranges i.e. 1st ref is absolute, 2nd is relative
     *
     * @param boolean relative
     */
    override var isRowRel: Boolean
        get() = super.isRowRel
        set(relative) {
            if (this.fRwRel != relative) {
                this.fRwRel = relative
                firstPtg!!.isRowRel = relative
                lastPtg!!.isRowRel = relative
                updateRecord()
            }
        }

    /**
     * link to the externsheet to be automatically updated upon removals
     */
    override fun addListener() {
        try {
            parentRec.workBook!!.externSheet!!.addPtgListener(this)
        } catch (e: Exception) {
            // no need to output here.  NullPointer occurs when a ref has an invalid ixti, such as when a sheet was removed  Worksheet exception could never really happen.
        }

    }

    override fun toString(): String? {
        return string
    }

    init {
        opcode = 0x3b
        record = ByteArray(11)
        record[0] = opcode // ""
        this.is3dRef = true
    }

    constructor(useReferenceTracker: Boolean) : this() {
        this.useReferenceTracker = useReferenceTracker
    }

    /**
     * set the Ptg Id type to one of:
     * VALUE, REFERENCE or Array
     * 25H (tAreaR), 45H (tAreaV), 65H (tAreaA)
     * <br></br>The Ptg type is important for certain
     * functions which require a specific type of operand
     */
    override fun setPtgType(type: Short) {
        when (type) {
            Ptg.VALUE -> opcode = 0x5B
            Ptg.REFERENCE -> opcode = 0x3B
            Ptg.ARRAY -> opcode = 0x7B
        }
        record[0] = opcode
    }


    /**
     * @return Returns the ixti.
     */
    override fun getIxti(): Short {    // only valid for 3d refs!!!!
        return ixti
    }


    /**
     * set the pointer into the Externsheet Rec.
     * this is only valid for 3d refs
     */
    override fun setIxti(ixf: Short) {
        if (ixti != ixf) {
            ixti = ixf
            // this seems to be only one byte...
            if (record != null) {
                record[1] = ixf.toByte()
                populateVals()    // add listener is done here
            }
        }
    }


    /**
     * get the worksheet that this ref is on
     */
    fun getSheets(b: WorkBook): Array<Boundsheet> {
        val bsa = b.externSheet!!.getBoundSheets(this.ixti.toInt())
        if (bsa!![0] == null)
        // 20080303 KSC: Catch Unresolved External refs
            Logger.logErr("PtgArea3d.getSheet: Unresolved External Worksheet")
        return bsa
    }


    /**
     * constructor, takes the array of the ptgRef, including
     * the identifier so we do not need to figure it out again later...
     * also takes the parent rec -- needed to init the sub-ptgs
     *
     * @param b
     * @param parent
     */
    fun init(b: ByteArray, parent: XLSRecord) {
        ixti = ByteTools.readShort(b[1].toInt(), b[2].toInt())
        record = b
        this.parentRec = parent
        populateVals()
    }

    /**
     * Throw this data into two ptgref's
     */
    override fun populateVals() {
        val temp1 = ByteArray(7)    // PtgRef3d is 7 bytes
        val temp2 = ByteArray(7)
        // Encoded Cell Range Address:
        // 0-2= first row
        // 2-4= last row
        // 4-6= first col
        // 6-8= last col
        // Encoded Cell Address:
        // 0-2=	row index
        // 2-4= col index + relative flags
        try {
            temp1[0] = 0x3a
            temp1[1] = record[1]    // ixti
            temp1[2] = record[2]    // ""
            temp1[3] = record[3]    // first row
            temp1[4] = record[4]    // ""
            temp1[5] = record[7]    // first col
            temp1[6] = record[8]    // ""

            temp2[0] = 0x3a
            temp2[1] = record[1]        // ixti
            temp2[2] = record[2]        // ""
            temp2[3] = record[5]        // last row
            temp2[4] = record[6]        // ""
            temp2[5] = record[9]        // last col
            temp2[6] = record[10]        //	""
        } catch (e: Exception) {
            //should never happen!
            return
        }

        // pass in parent_rec so can properly set formulaRow/formulaCol
        firstPtg = PtgRef3d(false)

        // the following method registers the Ptg with the ReferenceTracker
        firstPtg!!.parentRec = parentRec
        firstPtg!!.sheetName = this.sheetName
        firstPtg!!.init(temp1)

        lastPtg = PtgRef3d(false)
        lastPtg!!.parentRec = parentRec
        lastPtg!!.sheetName = this.lastSheetName
        lastPtg!!.init(temp2)
        // flag if it's an external reference

        isExternalRef = (firstPtg as PtgRef3d).isExternalLink || (lastPtg as PtgRef3d).isExternalLink

        setWholeRowCol()

        // take 1st Ptg as sample for relative state
        this.fColRel = firstPtg!!.isColRel
        this.fRwRel = firstPtg!!.isRowRel
        //init sets formula row to 1st row for a shared formula; adjust here
        if (parentRec != null && parentRec is Shrfmla) {
            lastPtg!!.formulaRow = (parentRec as Shrfmla).lastRow
            lastPtg!!.formulaCol = (parentRec as Shrfmla).lastCol
        }
        this.hashcode = super.hashCode
    }


    /**
     * set Ptg Location to parsed location
     *
     * @param loc String[] sheet1, range, sheet2, exref1, exref2
     */
    override fun setLocation(s: Array<String>) {
        try {
            if (useReferenceTracker && locax != null)
            // if in tracker already, remove
                this.removeFromRefTracker()
        } catch (e: Exception) {
            // will happen if this is not in tracker yet
        }

        var sheetname2: String? = null
        var range = ""
        range = s[1]
        if (s[0] != null) {    // has a sheet in the address
            sheetname = s[0]
            sheetname2 = s[2]
            if (sheetname2 == null)
                sheetname2 = sheetname
            // revised so can set ixti on error'd references
            var b: WorkBook? = null
            var xsht: Externsheet? = null
            if (parentRec != null) {
                b = parentRec.workBook
                if (b == null)
                    b = parentRec.sheet!!.workBook
            }
            try {
                xsht = b!!.externSheet
                val boundnum = b.getWorkSheetByName(sheetname).sheetNum
                var boundnum2 = boundnum // it could possibly be a 3d ref - check
                if (sheetname != sheetname2 && sheetname2 != null)
                    boundnum2 = b.getWorkSheetByName(sheetname2).sheetNum
                this.setIxti(xsht!!.insertLocation(boundnum, boundnum2).toShort())
            } catch (e: WorkSheetNotFoundException) {
                try {
                    // try to link to external sheet, if possible
                    val boundnum = xsht!!.getXtiReference(s[0], s[0])
                    if (boundnum == -1) {    // can't resolve
                        this.setIxti(xsht.insertLocation(boundnum, boundnum).toShort())
                    } else {
                        this.setIxti(boundnum.toShort())
                        this.isExternalRef = true
                    }
                } catch (ex: Exception) {
                }

            }

        } else if (parentRec != null) {
            sheetname2 = parentRec.sheet!!.sheetName
            sheetname = sheetname2    // use parent rec's sheet
        }
        var i = range.indexOf(":")
        if (i < 0) {
            range = "$range:$range"
            i = range.indexOf(":")
        }

        var firstcell = range.substring(0, i)
        var lastcell = range.substring(i + 1)
        if (sheetname != null)
            firstcell = "$sheetname!$firstcell"
        if (sheetname2 != null)
            lastcell = "$sheetname2!$lastcell"
        if (s[3] != null)
        // store OOXML External References
            firstcell = s[3] + firstcell
        if (s[4] != null)
            lastcell = s[4] + lastcell

        if (firstPtg == null) {
            firstPtg = PtgRef3d(false)
            firstPtg!!.parentRec = this.parentRec
        }
        firstPtg!!.sheetname = sheetname
        firstPtg!!.location = firstcell
        (firstPtg as PtgRef3d).setIxti(this.ixti)

        if (lastPtg == null) {
            lastPtg = PtgRef3d(false)
            lastPtg!!.parentRec = this.parentRec
        }
        lastPtg!!.sheetname = sheetname2
        lastPtg!!.location = lastcell
        (lastPtg as PtgRef3d).setIxti(this.ixti)

        this.setWholeRowCol()
        this.updateRecord()
        // TODO: must deal with non-symmetrical absolute i.e. if first and last ptgs don't match
        this.fRwRel = firstPtg!!.fRwRel
        this.fColRel = firstPtg!!.fColRel
        hashcode = hashCode
        if (useReferenceTracker) {
            if (!this.isWholeCol && !this.isWholeRow)
                this.addToRefTracker()
            else
                useReferenceTracker = false
        }
    }


    /**
     * returns whether this CellRange Contains a Cell
     *
     * @param the cell to test
     * @return whether the cell is in the range
     */
    override fun contains(ch: CellHandle): Boolean {
        val chsheet = ch.workSheetName
        sheetName
        if (!chsheet.equals(sheetname!!, ignoreCase = true)) return false
        val adr = ch.cellAddress
        //      FIX broken COLROW
        val rc = ExcelTools.getRowColFromString(adr)
        return contains(rc)
    }

    /**
     * Switches the two internal ptgref3ds to a new
     * sheet.
     */
    open fun setReferencedSheet(b: Boundsheet) {
        (firstPtg as PtgRef3d).setReferencedSheet(b)
        (lastPtg as PtgRef3d).setReferencedSheet(b)
        val boundnum = b.sheetNum
        val xsht = b.workBook!!.getExternSheet(true)
        //TODO: add handling for multi-sheet reference.  Already handled in externsheet
        try {
            this.sheetname = null    // 20100218 KSC: RESET
            val xloc = xsht!!.insertLocation(boundnum, boundnum)
            setIxti(xloc.toShort())
        } catch (e: WorkSheetNotFoundException) {
            Logger.logErr("Unable to set referenced sheet in PtgRef3d $e")
        }

    }

    /**
     * Updates the record bytes so it can be pulled back out.
     */
    override fun updateRecord() {
        comps = null
        val first = firstPtg!!.record
        val last = lastPtg!!.record
        // KSC: this apparently is what excel wants:
        if (isWholeRow)
            first[5] = 0
        if (isWholeCol) {
            first[3] = 0
            first[4] = 0
        }
        // the last record has an extra identifier on it.
        val newrecord = ByteArray(Ptg.PTG_AREA3D_LENGTH)
        newrecord[0] = 0x3B
        System.arraycopy(first, 1, newrecord, 1, 2)
        System.arraycopy(first, 3, newrecord, 3, 2)
        System.arraycopy(last, 3, newrecord, 5, 2)
        System.arraycopy(first, 5, newrecord, 7, 2)
        System.arraycopy(last, 5, newrecord, 9, 2)
        record = newrecord
        if (parentRec != null) {
            if (this.parentRec is Formula)
                (this.parentRec as Formula).updateRecord()
            else if (this.parentRec is Name)
                (this.parentRec as Name).updatePtgs()
        }
    }

    override fun close() {
        super.close()
        if (comps != null)
            for (i in comps!!.indices) {
                comps!![i].close()
                comps[i] = null
            }
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1176168076050592292L
    }

}