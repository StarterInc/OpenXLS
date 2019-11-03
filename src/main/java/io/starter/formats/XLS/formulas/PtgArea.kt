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
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger

import java.util.Vector


/**
 * ptgArea is a reference to an area (rectangle) of cells.
 * Essentially it is a collection of two ptgRef's, so it will be
 * treated that way in the code...
 *
 * <pre>
 * Offset      Size    Contents
 * ----------------------------------------------------
 * 0			2 		Index to first row (065535) or offset of first row (method [B], -3276832767)
 * 2 			2 		Index to last row (065535) or offset of last row (method [B], -3276832767)
 * 4 			2 		Index to first column or offset of first column, with relative flags (see table above)
 * 6 			2 		Index to last column or offset of last column, with relative flags (see table above)
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
 * @see Ptg
 *
 * @see Formula
 */
open class PtgArea : PtgRef, Ptg {

    override val isOperand: Boolean
        get() = true

    override val isReference: Boolean
        get() = true

    var firstPtg: PtgRef? = null
    var lastPtg: PtgRef? = null

    /**
     * Returns all of the cells of this range as PtgRef's.
     * This includes empty cells, values, formulas, etc.
     * Note the setting of parent-rec requires finding the cell
     * the PtgRef refer's to.  If that is null then the PtgRef
     * will exist, just with a null value.  This could cause issues when
     * programatically populating cells.
     */
    override//       TODO: check rc sanity here
            /*if (this.wholeRow) {
			startcol= 0;
			endcol= this.getSheet().getMaxCol();
			startrow= endrow= firstPtg.rw;
        } if (this.wholeCol) {
		    startrow= 0;	// Get Actual Coordinates
			startcol= endcol= firstPtg.col;
			endrow= this.getSheet().getMaxRow();
        } */// usually don't need to set sheet on setlocation becuase uses parent_rec's sheet
    // cases of named range or if location sheet does not = parent_rec sheet, set sheet explicitly
    // usual case, don't need to set sheet
    // loop through the cols
    // loop through the rows inside
    val components: Array<Ptg>?
        get() {
            val v = Vector()
            try {
                var startcol = -1
                var startrow = -1
                var endrow = -1
                var endcol = -1
                var startloc: IntArray? = null
                var endloc: IntArray? = null
                if (firstPtg != null) {
                    startloc = firstPtg!!.realIntLocation
                    startcol = startloc!![1]
                    startrow = startloc[0]
                } else {
                    startloc = ExcelTools.getRangeRowCol(locax!!)
                    startcol = startloc!![1]
                    startrow = startloc[0]
                }

                if (lastPtg != null) {
                    endloc = lastPtg!!.realIntLocation
                    endcol = endloc!![1]
                    endrow = endloc[0]
                } else {
                    endloc = ExcelTools.getRangeRowCol(locax!!)
                    endcol = endloc!![3]
                    endrow = endloc[2]
                }
                var sht: String? = null
                val sh = parentRec.sheet
                if (sh == null || this.sheetname != null && this.sheetname != sh.sheetName) {
                    if (sh == null || GenericPtg.qualifySheetname(this.sheetname) != GenericPtg.qualifySheetname(sh.sheetName))
                        sht = this.sheetname!! + "!"
                }
                while (startcol <= endcol) {
                    var rowholder = startrow
                    while (rowholder <= endrow) {
                        val displaycol = ExcelTools.getAlphaVal(startcol)
                        val displayrow = rowholder + 1
                        val pref: PtgRef
                        if (sht == null)
                            pref = PtgRef(displaycol + displayrow, parentRec, false)
                        else
                            pref = PtgRef(sht + displaycol + displayrow, parentRec, false)
                        v.add(pref)
                        rowholder++
                    }
                    startcol++
                }
            } catch (e: Exception) {
                Logger.logErr("calculating formula range value failed.", e)
            }

            val pref = arrayOfNulls<PtgRef>(v.size)
            v.toTypedArray()
            return pref
        }


    /**
     * returns the row/col ints for the ref
     *
     *
     * Format is FirstRow,FirstCol,LastRow,LastCol
     *
     * @return
     */
    override val rowCol: IntArray?
        get() {
            if (firstPtg == null) {
                return null
            }
            if (lastPtg == null && firstPtg != null) {
                val rc1 = firstPtg!!.rowCol
                return intArrayOf(rc1[0], rc1[1], rc1[0], rc1[1])
            }
            val rc1 = firstPtg!!.rowCol
            val rc2 = lastPtg!!.rowCol
            return intArrayOf(rc1[0], rc1[1], rc2[0], rc2[1])
        }


    // private byte[] PROTOTYPE_BYTES = {0x25, 0, 0, 0, 0, 0, 0, 0, 0};

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    override val string: String?
        get() = this.location

    override// 20080221 KSC: just set parent_rec super.setParentRec(rec);
    var parentRec: XLSRecord
        get
        set(rec) {
            super.parentRec = rec
            if (firstPtg != null)
                firstPtg!!.parentRec = parentRec
            if (lastPtg != null)
                lastPtg!!.parentRec = parentRec
        }

    /**
     * returns the location of the ptg as an array of ints.
     * [0] = firstRow
     * [1] = firstCol
     * [2] = lastRow
     * [3] = lastCol
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

    /*
        Returns the location of the Ptg as a string
    */
    /* Set the location of this PtgRef.  This takes a location
       such as "a14:b15"
    */
    override var location: String?
        get() {
            val lc = locationHelper
            locax = lc
            return lc
        }
        set(address) {
            val s = ExcelTools.stripSheetNameFromRange(address)
            setLocation(s)
            this.hashcode = hashCode
        }

    private//String loc= null;
    // we tried!
    // sheet, addr
    // sheet, addr
    //if (addr1.equals(addr2))	// this is proper but makes so many assertions fail, revert for now
    //return addr2;
    // no sheetname avail
    // handle OOXML external references
    // have sheetname
    // range is in one sheet
    // both sheets in sub-ptgs are null
    // 20090325 KSC: handle OOXML external references
    // 20081215 KSC:
    // only 1 sheetnaame specified
    // 20081215 KSC:
    // otherwise, include both sheets in return string
    val locationHelper: String
        get() {
            if (firstPtg == null || lastPtg == null) {
                this.populateVals()
                if (firstPtg == null || lastPtg == null)
                    throw AssertionError("PtgArea.getLocationHelper null ptgs")
            }
            val s = firstPtg!!.location
            val y = lastPtg!!.location

            val loc1 = ExcelTools.stripSheetNameFromRange(s!!)
            val loc2 = ExcelTools.stripSheetNameFromRange(y!!)

            var sh1: String? = loc1[0]
            var sh2: String? = loc2[0]
            val addr1 = loc1[1]
            val addr2 = loc2[1]

            if (this !is PtgArea3d) {
                return "$addr1:$addr2"
            }

            if (sh1 == null) sh1 = sheetname
            if (sh1 == null)
                return "$addr1:$addr2"
            if (externalLink1 > 0)
                sh1 = "[$externalLink1]$sh1"
            if (externalLink2 > 0 && sh2 != null)
                sh2 = "[$externalLink2]$sh2"

            sh1 = GenericPtg.qualifySheetname(sh1)
            if (sh1 == sh2) {
                if (sh1 != "") {
                    return if (addr1 != addr2)
                        "$sh1!$addr1:$addr2"
                    else
                        "$sh1!$addr1"
                } else if (sheetname != null) {
                    sh1 = sheetname
                    if (externalLink1 > 0)
                        sh1 = "[$externalLink1]$sh1"
                    sh1 = GenericPtg.qualifySheetname(sh1)
                    return if (addr1 != addr2) "$sh1!$addr1:$addr2" else "$sh1!$addr1"
                }
            } else if (sh2 == null) {
                return if (addr1 != addr2) "$sh1!$addr1:$addr2" else "$sh1!$addr1"
            }
            sh2 = GenericPtg.qualifySheetname(sh2)
            return "$sh1:$sh2!$addr1:$addr2"
        }

    override val length: Int
        get() = Ptg.PTG_AREA_LENGTH

    /*
        returns the sum of all fields within this range
        ** this may need to be modified, as we might not always want sum's
        *20080730 KSC: Excel does *NOT* sum these values except in array formulas

        TODO:  Calculate cell values that are result of formula - take care of
        in cell?
        From Excel File Format Documentation:
			Value class tokens will be changed dependent on further conditions. In array type functions and name type
		  	functions, or if the forced array class state is set, it is changed to array class. In all other cases (cell type formula
		  	without forced array class), value class is retained.
    */
    override// 20080214 KSC: underlying cells may have changed ...if(refCell==null)
    // 20090203 KSC
    //Double d = new Double(s);
    // 20090202 KSC: was +=
    //DblVal();	// 20090202 KSC: was +=
    // keep going???
    //DblVal();		// 20090202 KSC: was +=
    // 20080730 KSC: if not an array, retrieve only 1st referenced cell value
    //
    // 20090817 KSC:  [BugTracker 2683]
    //new Double(returnval);
    val value: Any?
        get() {
            refCell = this.refCells
            var returnval: Any = 0
            var retstr: String? = null
            var array: String? = ""
            val isArray = this.parentRec is Array
            for (t in refCell!!.indices) {
                val cel = refCell!![t] ?: continue

                try {
                    val f = cel.formulaRec
                    if (f != null) {
                        val oby = f.calculateFormula()
                        val s = oby.toString()
                        try {
                            returnval = Double(s)
                        } catch (ex: NumberFormatException) {
                            retstr = s
                        }

                    } else {
                        returnval = cel.internalVal
                    }
                } catch (e: FunctionNotSupportedException) {
                } catch (e: Exception) {
                    returnval = cel.internalVal
                }

                if (!isArray)
                    break
                if (retstr != null)
                    array = "$array$retstr,"
                else
                    array = "$array$returnval,"
                retstr = null
            }
            if (isArray && array != null && array.length > 1) {
                array = "{" + array.substring(0, array.length - 1) + "}"
                val pa = PtgArray()
                pa.setVal(array)
                return pa
            }
            return retstr ?: returnval

        }


    /**
     * @return Returns the refCell.
     */
    override// this.parent_rec.getWorkBook().getWorkSheetByName(this.getSheetName());
    // handle misc sheets
    // guard against NPEs
    // 20080212 KSC
    //			 TODO: check rc sanity here
    // loop through the cols
    // 20090521 KSC: may have range switched so that firstPtg>lastPtg (example in named ranges in tcr_formatted_2007.xlsm)
    // 20090521 KSC: may have range switched so that firstPtg>lastPtg (example in named ranges in tcr_formatted_2007.xlsm)
    // 20090521 KSC: try to handle both cases i.e. ranges such that first<last or first>last
            /*
            for (;startcol<=endcol;startcol++){ // loop the cols
                int rowpos = startrow;
                for (;rowpos<=endrow;rowpos++){ // loop through the rows
                    Row r = bs.getRowByNumber(rowpos);
                    if (r!=null)
                        refCell[rowctr] = (BiffRec)r.getCell((short)(startcol));
                    rowctr++;
                }
            }
            */ val refCells: Array<BiffRec>?
        get() {
            val returnval = 0.0
            try {
                var bs: Boundsheet? = null
                sheetName
                if (sheetname != null) {
                    try {
                        bs = this.parentRec.workBook!!.getWorkSheetByName(sheetname)
                    } catch (ex: Exception) {
                        bs = parentRec.sheet
                    }

                } else {
                    bs = parentRec.sheet
                    sheetname = bs!!.sheetName
                }
                val startloc = firstPtg!!.intLocation
                var startcol = startloc!![1]
                val startrow = startloc[0]
                val endloc = lastPtg!!.realIntLocation
                var endcol = endloc[1]
                var endrow = endloc[0]
                var numcols = endcol - startcol
                if (numcols < 0)
                    numcols = startcol - endcol
                numcols++
                var numrows = endrow - startrow
                if (numrows < 0)
                    numrows = startrow - endrow
                numrows++
                var totcell = numcols * numrows
                if (totcell == 0) totcell++
                if (totcell < 0) {
                    Logger.logErr("PtgArea.getRefCells.  Error in Ptg locations: " + firstPtg!!.toString() + ":" + lastPtg!!.toString())
                    totcell = 0
                }
                refCell = arrayOfNulls(totcell)
                var rowctr = 0
                if (startcol < endcol)
                    endcol++
                else
                    endcol--
                if (startrow < endrow)
                    endrow++
                else
                    endrow--
                while (startcol != endcol) {
                    var rowpos = startrow
                    while (rowpos != endrow) {
                        val r = bs!!.getRowByNumber(rowpos)
                        if (r != null)
                            refCell[rowctr] = r.getCell(startcol.toShort())
                        rowctr++
                        if (rowpos < endrow)
                            rowpos++
                        else
                            rowpos--
                    }
                    if (startcol < endcol)
                        startcol++
                    else
                        startcol--
                }
            } catch (ex: Exception) {
                Logger.logErr("PtgArea.getRefCells failed.", ex)
            }

            return refCell
        }

    protected override val hashCode: Long
        get() = lastPtg!!.hashcode + firstPtg!!.hashcode * (XLSConstants.MAXCOLS.toLong() + XLSConstants.MAXROWS.toLong() * XLSConstants.MAXCOLS)

    /* constructor, takes the array of the ptgRef, including
    the identifier so we do not need to figure it out again later...
    */
    override fun init(b: ByteArray) {
        locax = null // cache reset
        opcode = b[0]
        record = b
        this.populateVals()
    }

    /*
     Throw this data into two ptgref's
    */
    public override fun populateVals() {
        val temp1 = ByteArray(5)
        val temp2 = ByteArray(5)
        temp1[0] = 0x24
        temp2[0] = 0x24
        System.arraycopy(record, 1, temp1, 1, 2)
        System.arraycopy(record, 5, temp1, 3, 2)
        System.arraycopy(record, 3, temp2, 1, 2)
        System.arraycopy(record, 7, temp2, 3, 2)
        try {
            sheetName // 20080212 KSC:
        } catch (we: WorkSheetNotFoundException) {
            Logger.logErr(we)
        }

        firstPtg = PtgRef(temp1, parentRec, false)    // don't add to ref tracker as it's part of area
        firstPtg!!.sheetname = sheetname

        lastPtg = PtgRef(temp2, parentRec, false)        // don't add to ref tracker as it's part of area
        lastPtg!!.sheetname = sheetname
        setWholeRowCol()
        this.hashcode = hashCode
    }

    /**
     * returns whether this CellRange Contains a Cell
     *
     * @param the cell to test
     * @return whether the cell is in the range
     */
    open operator fun contains(ch: CellHandle): Boolean {
        val chsheet = ch.workSheetName
        var mysheet = ""
        if (this.parentRec != null) {
            val b = this.parentRec
            if (b.sheet != null) {
                mysheet = b.sheet.sheetName
            }
        }
        if (!chsheet.equals(mysheet, ignoreCase = true))
            return false
        val adr = ch.cellAddress
        //      FIX broken COLROW
        val rc = ExcelTools.getRowColFromString(adr)
        return contains(rc)
    }

    /**
     * check to see if the sheet and row/col are contained
     * in this ref
     *
     * @param sheetname
     * @param rc
     * @return
     */
    fun contains(sn: String, rc: IntArray): Boolean {
        if (sheetname == null) {
            try {
                sheetname = this.sheetName
            } catch (e: Exception) {
            }

        }
        return if (!sn.equals(sheetname!!, ignoreCase = true)) false else contains(rc)
    }

    /**
     * returns whether this PtgArea Contains the specified row/col coordinate
     *
     *
     *
     *
     * [0] = firstrow
     * [1] = firstcol
     * [2] = lastrow
     * [3] = lastcol
     *
     * @param the rc coordinates to test
     * @return whether the coordinates are in the range
     */
    operator fun contains(rc: IntArray): Boolean {
        val thisRange = this.intLocation
        // test the first rc
        if (rc[0] < thisRange!![0]) return false // row above the first ref row?
        if (rc[0] > thisRange[2]) return false // row after the last ref row?

        return if (rc[1] < thisRange[1]) false else rc[1] <= thisRange[3] // col before the first ref col?
// col after the last ref col?
    }

    /**
     * Creates a new PtgArea from 2 component ptgs, used by shared formula
     * to create ptgareas.  ptg1 should be upperleft corner, ptg2 bottomright
     */
    constructor(ptg1: PtgRef, ptg2: PtgRef, parent: XLSRecord) : this() {
        firstPtg = ptg1
        lastPtg = ptg2
        parentRec = parent
        this.hashcode = hashCode
        this.updateRecord()
    }

    /*
     * Creates a new PtgArea.  The parent rec is needed
     * as getting a value goes to the boundsheet to determine
     * values.  The parent rec *must* be on the same sheet
     * as the PtgArea referenced!
     *
     * relativeRefs = true is excel default
     */
    @JvmOverloads
    constructor(range: String, parent: XLSRecord, relativeRefs: Boolean = true) : this() {
        val loc = ExcelTools.getRangeRowCol(range)
        val temp = IntArray(2)
        temp[0] = loc[0]
        temp[1] = loc[1]
        var res = ExcelTools.formatLocation(temp, relativeRefs, relativeRefs)
        firstPtg = PtgRef(res, parent, false)
        temp[0] = loc[2]
        temp[1] = loc[3]
        res = ExcelTools.formatLocation(temp, relativeRefs, relativeRefs)
        lastPtg = PtgRef(res, parent, false)
        setWholeRowCol()
        parentRec = parent
        this.hashcode = hashCode
        this.updateRecord()
    }

    /*
     * Creates a new PtgArea using an int array as [r,c,r1,c1].
     * The parent rec is needed
     * as getting a value goes to the boundsheet to determine
     * values.  The parent rec *must* be on the same sheet
     * as the PtgArea referenced!
     *
     * relativeRefs = true is excel default
     */
    constructor(loc: IntArray, parent: XLSRecord, relativeRefs: Boolean) {
        val temp = IntArray(2)
        temp[0] = loc[0]
        temp[1] = loc[1]
        var res = ExcelTools.formatLocation(temp, relativeRefs, relativeRefs)
        firstPtg = PtgRef(res, parent, false)
        temp[0] = loc[2]
        temp[1] = loc[3]
        res = ExcelTools.formatLocation(temp, relativeRefs, relativeRefs)
        lastPtg = PtgRef(res, parent, false)
        setWholeRowCol()
        parentRec = parent
        this.hashcode = hashCode
        this.updateRecord()
    }

    /**
     * set the wholeRow and/or wholeCol flag for this PtgArea
     * for ranges such as:
     * $B:$B and $5:%9
     */
    fun setWholeRowCol() {
        if (firstPtg!!.rw <= 1 && lastPtg!!.isWholeCol)
        // TODO: inconsistencies in 0-based or 1-based rows
            this.isWholeCol = true
        this.isWholeRow = lastPtg!!.isWholeRow
        if (this.isWholeCol)
            useReferenceTracker = false
    }

    /*
     * Default constructor
     */
    constructor() {
        record = ByteArray(9)
        opcode = 0x25
        record[0] = 0x25
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
            Ptg.VALUE -> opcode = 0x45
            Ptg.REFERENCE -> opcode = 0x25
            Ptg.ARRAY -> opcode = 0x65
        }
        record[0] = opcode
    }


    override fun toString(): String? {
        var ret = string

        if (parentRec != null)
            if (ret!!.indexOf("!") < 0)
                try { // Catch WorkSheetNotFoundException to handle Unresolved External refs
                    sheetName
                    if (sheetname != null)
                        ret = "$sheetname!$ret"
                } catch (we: WorkSheetNotFoundException) {
                    Logger.logErr(we)
                }

        return ret
    }

    /**
     * set Ptg to parsed location
     *
     * @param loc String[] sheet1, range, sheet2, exref1, exref2
     */
    override fun setLocation(loc: Array<String>) {
        locax = null // cache reset
        if (firstPtg == null) {
            this.record = byteArrayOf(0x25, 0, 0, 0, 0, 0, 0, 0, 0)

            if (this.parentRec != null) this.populateVals()
        } else if (this.useReferenceTracker)
            this.removeFromRefTracker()
        var i = loc[1].indexOf(":")
        // handle single cell addresses as:  A1:A1
        if (i == -1) {
            loc[1] = loc[1] + ":" + loc[1]
            i = loc[1].indexOf(":")
        }
        var firstloc = loc[1].substring(0, i)
        var lastloc = loc[1].substring(i + 1)
        if (loc[0] != null)
            firstloc = loc[0] + "!" + firstloc
        if (loc[2] != null)
            lastloc = loc[2] + "!" + lastloc
        if (loc[3] != null)
        // 20090325 KSC: store OOXML External References
            firstloc = loc[3] + firstloc
        if (loc[4] != null)
        // 20090325 KSC: store OOXML External References
            lastloc = loc[4] + lastloc

        // TODO: do we need to remove refs from tracker?
        firstPtg!!.parentRec = this.parentRec
        lastPtg!!.parentRec = this.parentRec

        firstPtg!!.useReferenceTracker = false
        lastPtg!!.useReferenceTracker = false
        firstPtg!!.location = firstloc
        lastPtg!!.location = lastloc
        setWholeRowCol()
        this.hashcode = hashCode
        this.updateRecord()
        if (this.useReferenceTracker) {// check of boolean useReferenceTracker
            if (!this.isWholeCol && !this.isWholeRow)
                this.addToRefTracker()
            else
                useReferenceTracker = false
        }
    }


    /**
     * Set the location of this PtgArea.  This takes a location
     * such as {1,2,3,4}
     */
    override fun setLocation(rowcol: IntArray) {
        locax = null // cache reset
        if (firstPtg == null) {
            //this.record = new byte[] {0x25, 0, 0, 0, 0, 0, 0, 0, 0}; -- don't as can be called from PtgArea3d
            if (this.parentRec != null) this.populateVals()
        } else if (this.useReferenceTracker)
            this.removeFromRefTracker()

        // TODO: do we need to remove refs from tracker?
        firstPtg!!.parentRec = this.parentRec
        firstPtg!!.sheetName = sheetname
        lastPtg!!.parentRec = this.parentRec
        lastPtg!!.sheetName = sheetname

        firstPtg!!.useReferenceTracker = false
        lastPtg!!.useReferenceTracker = false
        firstPtg!!.setLocation(rowcol)
        val rc = IntArray(2)
        rc[0] = rowcol[2]
        rc[1] = rowcol[3]
        lastPtg!!.setLocation(rc)

        this.hashcode = hashCode
        this.updateRecord()
        if (this.useReferenceTracker)
        // check of boolean useReferenceTracker
            this.addToRefTracker()
    }

    /* Updates the record bytes so it can be pulled back out.
     */
    override fun updateRecord() {
        locax = null // cache reset
        val pols = intArrayOf(firstPtg!!.locationPolicy, lastPtg!!.locationPolicy)
        val first = firstPtg!!.record
        val last = lastPtg!!.record
        // the last record has an extra identifier on it.
        val newrecord = ByteArray(9)
        newrecord[0] = record[0]
        System.arraycopy(first, 1, newrecord, 1, 2)
        System.arraycopy(last, 1, newrecord, 3, 2)
        System.arraycopy(first, 3, newrecord, 5, 2)
        System.arraycopy(last, 3, newrecord, 7, 2)
        record = newrecord
        //        this.populateVals();
        if (parentRec != null) {
            if (this.parentRec is Formula)
                (this.parentRec as Formula).updateRecord()
            else if (this.parentRec is Name)
                (this.parentRec as Name).updatePtgs()
        }
        firstPtg!!.locationPolicy = pols[0]
        lastPtg!!.locationPolicy = pols[1]
    }


    /*
	    Returns all of the cells of this range as PtgRef's.
	    This includes empty cells, values, formulas, etc.
	    Note the setting of parent-rec requires finding the cell
	    the PtgRef refer's to.  If that is null then the PtgRef
	    will exist, just with a null value.  This could cause issues when
	    programatically populating cells.
	*/
    fun getColComponents(colNum: Int): Array<Ptg>? {
        if (colNum < 0) return null
        val lu = this.toString()
        val p = parentRec.workBook!!.refTracker!!.vlookups[lu]

        if (p != null) {
            val par = p as PtgArea?
            val ret = par!!.parentRec.workBook!!
                    .refTracker!!.lookupColCache[lu + ":" + Integer.valueOf(colNum)] as Array<Ptg>
            if (ret != null)
                return ret
        }

        var v: Array<PtgRef>? = null
        try {
            //			 TODO: check rc sanity here
            val startloc = firstPtg!!.realIntLocation
            val startrow = startloc[0]
            val endloc = lastPtg!!.realIntLocation
            var endrow = endloc[0]
            // error trap
            if (endrow < startrow)
            // can happen if wholerow/wholecol, getMaxRow may be less than startRow
                endrow = startrow
            var sz = endrow - startrow
            sz++
            v = arrayOfNulls(sz)
            // loop through the cols
            // loop through the rows inside
            var rowholder = startrow
            var pos = 0
            var sht = this.toString()
            if (sht!!.indexOf("!") > -1) sht = sht.substring(0, sht.indexOf("!"))
            while (rowholder <= endrow) {
                val displaycol = ExcelTools.getAlphaVal(colNum)
                val displayrow = rowholder + 1
                val loc = "$sht!$displaycol$displayrow"

                val pref = PtgRef(loc, parentRec, this.useReferenceTracker)

                v[pos++] = pref
                rowholder++
            }
        } catch (e: Exception) {
            Logger.logErr("Getting column range in PtgArea failed.", e)
        }

        // cache
        parentRec.workBook!!.refTracker!!.vlookups[this.toString()] = this
        parentRec.workBook!!.refTracker!!.lookupColCache[lu + ":" + Integer.valueOf(colNum)] = v
        return v
    }

    /**
     * return the ptg components for a certain column within a ptgArea()
     *
     * @param rowNum
     * @return all Ptg's within colNum
     */
    fun getRowComponents(rowNum: Int): Array<Ptg> {
        val v = FastAddVector()
        val allComponents = this.components
        for (i in allComponents!!.indices) {
            val p = allComponents[i] as PtgRef
            //			 TODO: check rc sanity here
            val x = p.realIntLocation
            if (x[0] == rowNum) v.add(p)
        }
        val pref = arrayOfNulls<PtgRef>(v.size)
        v.toTypedArray()
        return pref
    }

    /**
     * @return
     */
    fun getFirstPtg(): PtgRef? {
        return firstPtg
    }

    /**
     * @return
     */
    fun getLastPtg(): PtgRef? {
        return lastPtg
    }

    /**
     * @param ref
     */
    fun setFirstPtg(ref: PtgRef) {
        locax = null // cache reset
        firstPtg = ref
    }

    /**
     * @param ref
     */
    fun setLastPtg(ref: PtgRef) {
        locax = null // cache reset
        lastPtg = ref
    }

    companion object {

        val serialVersionUID = 666555444333222L
    }
}/*
     * Creates a new PtgArea.  The parent rec is needed
     * as getting a value goes to the boundsheet to determine
     * values.  The parent rec *must* be on the same sheet
     * as the PtgArea referenced!
     */