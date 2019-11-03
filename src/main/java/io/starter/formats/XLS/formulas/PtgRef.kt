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

import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.*
import io.starter.formats.cellformat.CellFormatFactory
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool


/**
 * ptgRef is a reference to a single cell.  It contains row and
 * column information, plus a grbit to determine whether these
 * values are relative or absolute.  This grbit is, stupidly, but expectedly,
 * encoded within the column value.
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rw          2       the row
 * 2           grbitCol    2       (see following table)
 *
 * Only the low-order 14 bits specify the Col, the other bits specify
 * relative vs absolute for both the col or the row.
 *
 * Bits        Mask        Name    Contents
 * -----------------------------------------------------
 * 15          8000h       fRwRel  =1 if row offset relative,
 * =0 if otherwise
 * 14          4000h       fColRel =1 if col offset relative,
 * =0 if otherwise
 * 13-0        3FFFh       col     Ordinal column offset or number
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Dbcell
 *
 * @see Row
 *
 * @see Cell
 *
 * @see XLSRecord
 */
open class PtgRef : GenericPtg, Ptg {
    var rw: Int = 0
    // TODO: We actually are talking about 2 different notions of relativity:
    // 1- Relativity based on shared formula parent formula row/col
    // 2- Relative/Absolute in terms of row movement ($'s mean reference is ABSOLUTE)
    // we are combining the two concepts erroneously
    var fRwRel = true  //true is relative, false is absolute (=$'s)
    var fColRel = true
    var col: Int = 0
    var is3dRef = false
        protected set
    private var cachedLocation: String? = null

    var formulaRow: Int = 0
    var formulaCol: Int = 0
    internal var refCell: Array<BiffRec>? = arrayOfNulls(1)
    var sheetname: String? = null

    protected var externalLink1 = 0
    protected var externalLink2 = 0
    /**
     * Ptgs upkeep their mapping in reference tracker, however, some ptgs
     * are components of other Ptgs, such as individual ptg cells in a PtgArea.  These
     * should not be stored in the RT.
     */
    var useReferenceTracker = true

    internal var locax: String? = null
    //    String locstrax = null;
    var hashcode = -1L

    open val isRefErr: Boolean
        get() = false

    var isWholeRow = false
    var isWholeCol = false    // denotes a range which spans the entire row or column, a shorthand for checking end col or row # as this will vary between excel versions

    override val isOperand: Boolean
        get() = true

    override val isReference: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    override val string: String?
        get() = this.location

    /**
     * returns the String address of this ptg including sheet reference
     *
     * @return
     */
    // AI PtgRefs do not have location info
    val locationWithSheet: String
        get() {
            var ret = string
            if (ret == null && parentRec.opcode == XLSRecord.AI)
                return parentRec.toString()

            if (ret == null)
                return ""

            if (ret.indexOf("!") > -1)
                return ret

            ret = "$sheetname!$ret"

            return ret
        }


    /**
     * returns the row/col ints for the ref
     *
     * @return
     */
    open// if row truly references MAXROWS_BIFF8 comes out -
    val rowCol: IntArray
        get() {
            val ret = intArrayOf(this.rw, this.col)
            if (this.rw < 0) {
                ret[0] = XLSConstants.MAXROWS_BIFF8
                this.isWholeCol = true
            }
            return ret
        }


    /**
     * Returns the location of the Ptg as a string (ie c4)
     *
     * @see io.starter.formats.XLS.formulas.GenericPtg.getLocation
     */
    /**
     * Set the location of this PtgRef.  This takes a location
     * such as "a14",   also can take a absolute location, such as $A14
     */
    override//cache
    var location: String?
        get() {
            if (locax != null)
                return locax

            val adjusted = this.intLocation
            val s: String
            if (this.isWholeCol) {
                s = (if (fColRel) "" else "$") + ExcelTools.getAlphaVal(adjusted!![1])
            } else if (this.isWholeRow) {
                s = (if (fRwRel) "" else "$") + (adjusted!![0] + 1)
            } else {
                if (rw < 0 || col < 0)
                    return PtgRefErr().toString()

                s = (if (fColRel) "" else "$") + ExcelTools.getAlphaVal(adjusted!![1]) +
                        (if (fRwRel) "" else "$") + (adjusted[0] + 1)
            }
            locax = s
            return locax
        }
        set(address) {
            locax = null
            refCell = null
            if (record != null) {
                val s = ExcelTools.stripSheetNameFromRange(address)
                setLocation(s)
                locax = s[1]
            } else {
                Logger.logWarn("PtgRef.setLocation() failed: NO record data: $address")
            }
        }

    /**
     * Get the location of this ptgRef as an int array {row, col}.  0 based
     */
    override// the row is a relative location
    // the column is a relative location
    val intLocation: IntArray?
        get() {

            this.setIsWholeRowCol()
            var rowNew = rw
            var colNew = col
            try {
                val isExcel2007 = this.parentRec.workBook!!.isExcel2007
                if (fRwRel) {
                    rowNew += formulaRow
                }
                if (fColRel) {
                    colNew += formulaCol
                }
                if (isWholeRow) {
                    if (!isExcel2007)
                        colNew = XLSConstants.MAXCOLS_BIFF8
                    else
                        colNew = XLSConstants.MAXCOLS
                }
                if (isWholeCol) {
                    if (isExcel2007)
                        rowNew = XLSConstants.MAXROWS - 1
                    else
                        rowNew = XLSConstants.MAXROWS_BIFF8 - 1
                }

            } catch (e: NullPointerException) {
            }

            return intArrayOf(rowNew, colNew)
        }

    /**
     * Get the location of this ptgRef as an int array {row, col}.  0 based
     * NOTE: this version of getIntLocation returns the actual or real coordinates
     * This may be different from getIntLocation when rw designates MAXROWS - in these cases,
     * this method will return real max rows
     */
    // the row is a relative location
    // the column is a relative location
    val realIntLocation: IntArray
        get() {
            var rowNew = rw
            var colNew = col
            if (fRwRel) {
                rowNew += formulaRow
            }
            if (fColRel) {
                colNew += formulaCol
            }

            if (isWholeCol || rowNew < 0)
                try {
                    if (rowNew < 0)
                        isWholeCol = true
                    rowNew = this.parentRec.sheet!!.maxRow
                } catch (e: Exception) {
                }

            if (isWholeRow || colNew >= XLSConstants.MAXCOLS)
                try {
                    colNew = this.parentRec.sheet!!.maxCol
                } catch (e: Exception) {
                }

            return intArrayOf(rowNew, colNew)
        }

    /**
     * Get the worksheet name this ptgref refers to
     *
     * @throws WorkSheetNotFoundException
     */
    /**
     * sets the sheetname for this
     *
     * @param sheetname
     */
    open// reference on different sheet than parent
    // no sheetname
    //handle external references (OOXML-specific)
    var sheetName: String?
        @Throws(WorkSheetNotFoundException::class)
        get() {

            if (locax != null) {
                if (locax!!.indexOf("!") > -1) {
                    sheetname = locax!!.substring(0, locax!!.indexOf("!"))
                }
            }
            if (sheetname == null && parentRec != null) {
                if (parentRec.sheet != null) {
                    sheetname = parentRec.sheet!!.sheetName
                }
            }

            if (sheetname == null) {
                return ""
            } else {
                if (externalLink1 > 0) {
                    if (sheetname!![0] == '\'')
                        sheetname = sheetname!!.substring(1, sheetname!!.length - 1)
                    sheetname = "[$externalLink1]$sheetname"
                }
                sheetname = GenericPtg.qualifySheetname(sheetname)

            }
            return sheetname
        }
        set(sheetname) {
            this.sheetname = sheetname
        }

    override val length: Int
        get() = Ptg.PTG_REF_LENGTH

    /**
     * return truth of "reference is blank"
     *
     * @return
     */
    override//getOpcode()==BLANK);
    val isBlank: Boolean
        get() {
            refCells
            return refCell!![0] == null || (refCell!![0] as XLSRecord).isBlank
        }

    /**
     * returns the value of the cell refereced by the PtgRef
     */
    override// assume zero, which the vast majority of cases are
    val value: Any?
        get() {
            refCells
            var retValue: Any? = null
            if (refCell!![0] != null) {
                if (refCell!![0].formulaRec != null) {
                    val f = refCell!![0].formulaRec
                    retValue = f.calculateFormula()
                    return retValue
                } else {
                    if (refCell!![0].dataType == "Float") {
                        retValue = refCell!![0].dblVal
                        return retValue
                    } else {
                        retValue = refCell!![0].internalVal
                        return retValue
                    }
                }
            } else {
                try {
                    if (!this.parentRec.sheet!!.window2!!.showZeroValues)
                        return null
                } catch (e: NullPointerException) {
                }

                return Integer.valueOf(0)
            }
        }

    /**
     * returns the value of the ptg formatted via the underlying cell's number format
     *
     * @return String underlying cell value formatted via cell's format pattern
     */
    // assume zero, which the vast majority of cases are
    val formattedValue: String
        get() {
            refCells
            var retValue: Any? = null
            val cell = refCell!![0]

            if (cell != null) {
                if (cell.formulaRec != null) {
                    val f = cell.formulaRec
                    retValue = f.calculateFormula()
                } else {
                    if (cell.dataType == "Float") {
                        retValue = cell.dblVal
                    } else {
                        retValue = cell.internalVal
                    }
                }
                return CellFormatFactory.fromPatternString(
                        cell.xfRec.formatPattern).format(retValue)
            } else {
                try {
                    if (!this.parentRec.sheet!!.window2!!.showZeroValues)
                        return ""
                } catch (e: NullPointerException) {
                }

                return "0"
            }
        }

    /**
     * @return Returns the refCell.
     */
    open val refCells: Array<BiffRec>?
        get() {
            refCell = arrayOfNulls(1)
            try {
                var bs: Boundsheet? = null
                if (sheetname != null && parentRec != null) {
                    bs = this.parentRec.workBook!!.getWorkSheetByName(sheetname)
                } else if (parentRec != null) {
                    bs = parentRec.sheet
                }
                refCell[0] = bs!!.getCell(rw, col)
            } catch (ex: Exception) {
            }

            return refCell
        }

    override// trap formulaRow, Col for relative PtgRefs
    var parentRec: XLSRecord
        get
        set(f) {
            parentRec = f
            setRelativeRowCol()
        }

    /**
     * sets the row to be relative (relative is true) or absolute (relative is false)
     * <br></br>absolute references do not shift upon row inserts or deletes
     *
     * @param boolean relative
     */
    open var isRowRel: Boolean
        get() = fRwRel
        set(relative) {
            if (fRwRel != relative) {
                locax = null
                fRwRel = relative
                updateRecord()
            }
        }

    /**
     * sets the column to be relative (relative is true) or absolute (relative is false)
     * <br></br>absolute references do not shift upon column inserts or deletes
     *
     * @param boolean relative
     */
    open var isColRel: Boolean
        get() = fColRel
        set(relative) {
            if (fColRel != relative) {
                locax = null
                fColRel = relative
                updateRecord()
            }
        }

    /**
     * uniquely identifies a row/col
     * to unencrypt:
     * col= hashcode%maxcols
     * row= hashcode/maxcols -1
     */
    protected open val hashCode: Long
        get() = if (rw >= 0)
            (col + (rw + 1) * XLSConstants.MAXCOLS).toLong()
        else
            (col + (XLSConstants.MAXROWS - rw + 1) * XLSConstants.MAXCOLS).toLong()


    override fun equals(ob: Any?): Boolean {
        return ob!!.hashCode() == this.hashCode()
    }

    constructor(rowcol: IntArray, x: XLSRecord, useRefTracker: Boolean) : this() {
        parentRec = x
        this.useReferenceTracker = useRefTracker
        setLocation(rowcol)
        updateRecord()
    }


    /**
     * This constructor is for programmatic creation of Ptg's
     * in this case we do not have the ptgid, just the refereced location
     */
    constructor(location: String, x: XLSRecord, utilizeRefTracker: Boolean) {
        this.useReferenceTracker = utilizeRefTracker
        opcode = 0x44  //0x24; defaulting to value operand
        record = ByteArray(5)
        record[0] = opcode
        parentRec = x    // MUST set before setLocation also sets formulaRow ...
        this.location = location
        this.setIsWholeRowCol()
        if (useReferenceTracker)
            addToRefTracker()
    }


    /**
     *
     * @param id
     */

    /**
     * set the Ptg Id type to one of:
     * VALUE, REFERENCE or Array
     * <br></br>The Ptg type is important for certain
     * functions which require a specific type of operand
     */
    open fun setPtgType(type: Short) {
        when (type) {
            Ptg.VALUE -> opcode = 0x44
            Ptg.REFERENCE -> opcode = 0x24
            Ptg.ARRAY -> opcode = 0x64
        }
        record[0] = opcode
    }


    /**
     * This constructor is for programmatic creation of Ptg's
     * in this case we do not have the ptgid, just the refereced location
     *
     *
     * this version sets the value of useReferenceTracker to avoid multiple entries due to area parent
     */
    constructor(bin: ByteArray, x: XLSRecord, utilizeRefTracker: Boolean) : this() {
        this.useReferenceTracker = utilizeRefTracker
        parentRec = x    //MUST DO BEFORE INIT ... also sets formulaRow ...
        init(bin)
        if (useReferenceTracker)
            addToRefTracker() // TODO: check subreference issue (if it's not a 'real' ptg)
    }

    /**
     * default constructor
     */
    constructor() {
        // 24H (tRefR), 44H (tRefV), 64H (tRefA)
        opcode = 0x44  // default to value operand
        record = ByteArray(5)
        record[0] = opcode
    }


    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    /**
     * parse all the values out of the byte array and
     * populate the classes values
     */
    protected open fun populateVals() {
        rw = readRow(record[1], record[2])
        val column = ByteTools.readShort(record[3].toInt(), record[4].toInt())
        // is the Row relative?
        fRwRel = column and 0x8000 == 0x8000
        // is the Column relative?
        fColRel = column and 0x4000 == 0x4000
        col = (column and 0x3fff).toShort().toInt()
        setRelativeRowCol()  // set formulaRow/Col for relative references if necessary
        this.setIsWholeRowCol()
        hashcode = hashCode
    }


    override fun toString(): String? {
        return string
    }

    /**
     * Clears the location cache when needed
     */
    fun clearLocationCache() {
        locax = null
    }

    /**
     * Does this ref reference an entire row (ie $1);
     *
     * @return
     */
    private fun referencesEntireRow(): Boolean {
        val isExcel2007 = this.parentRec.workBook!!.isExcel2007
        var colNew = col
        if (fColRel) {  // the row is a relative location
            colNew += formulaRow
        }
        if (colNew < 0) {   // have to assume that it's a wholeRow even if 2007
            return true
        } else if (colNew >= XLSConstants.MAXCOLS_BIFF8 - 1 && !isExcel2007) {
            return true
        }
        return if (this.cachedLocation != null && isExcel2007) {
            this.locationStringReferencesEntireRow()
        } else colNew == XLSConstants.MAXCOLS_BIFF8 - 1 && isExcel2007
        // This is unfortunately a bit of a hack due to biff 8 incompatibilies

    }

    /**
     * Check if the cached string location referrs to a full row
     *
     * @return
     */
    private fun locationStringReferencesEntireRow(): Boolean {
        if (this.cachedLocation != null) {
            val res = ExcelTools.getRowColFromString(cachedLocation!!)
            return res[1] < 0
        }
        return false
    }

    /**
     * Does this ref reference an entire col (ie $A);
     *
     * @return
     */
    private fun referencesEntireCol(): Boolean {
        var rowNew = rw
        val isExcel2007 = this.parentRec.workBook!!.isExcel2007
        if (fRwRel) {  // the row is a relative location
            rowNew += formulaRow
        }
        if (rowNew < 0) {
            return true
        } else if (rowNew >= XLSConstants.MAXROWS_BIFF8 - 1 && !isExcel2007) {
            rowNew = -1
            return true
        }
        return false
    }

    /**
     * Inspects the record to determin if it references whole
     * rows or columns and sets the values as required.
     *
     * @return
     */
    protected fun setIsWholeRowCol() {
        this.isWholeCol = referencesEntireCol()
        this.isWholeRow = referencesEntireRow()
    }

    /**
     * set Ptg to parsed location
     *
     * @param loc String[] sheet1, range, sheet2, exref1, exref2
     */
    open fun setLocation(loc: Array<String>) {
        if (useReferenceTracker)
            removeFromRefTracker()
        locax = null
        sheetname = loc[0]
        var addr = loc[1]
        cachedLocation = addr
        fRwRel = true
        fColRel = true
        if (addr.indexOf("$") == -1) {    // both row and col are relative refs, meaning moves/copies will change ref
            // relative link
            if (addr != "#REF!" && addr != "") {
                val res = ExcelTools.getRowColFromString(addr)
                col = res[1]
                rw = res[0]
            } else {
                col = -1
                rw = -1
            }
        } else {
            // absolute reference
            if (addr.substring(0, 1).equals("$", ignoreCase = true)) {
                fColRel = false
                addr = addr.substring(1)
            }
            if (addr.indexOf("$") != -1) {
                fRwRel = false
                addr = StringTool.strip(addr, "$")
            }
            var res: IntArray? = null
            try {
                res = ExcelTools.getRowColFromString(addr)
                col = res!![1]
                rw = res[0]
                if (col == -1 || rw == -1) {    // if wholerow or wholecol, must be absolute
                    fColRel = false
                    fRwRel = false
                }
            } catch (ie: IllegalArgumentException) {    //is it a wholerow/wholecol issue?
                if (Character.isDigit(addr[0])) { //assume wholecol ref
                    col = XLSConstants.MAXCOLS_BIFF8 - 1
                    rw = Integer.valueOf(addr).toInt() - 1
                    fColRel = false
                    fRwRel = false
                } else { //wholerow ref?
                    rw = -1
                    col = ExcelTools.getIntVal(addr)
                    fColRel = false
                    fRwRel = false
                }
            }

        }
        if (col == -1)
            isWholeRow = true
        if (rw == -1)
            isWholeCol = true
        this.setIsWholeRowCol()
        this.updateRecord()
        hashcode = hashCode
        // trap OOXML external reference link, if any
        if (loc[3] != null)
            externalLink1 = Integer.valueOf(loc[3].substring(1, loc[3].length - 1)).toInt()
        if (loc[4] != null)
            externalLink2 = Integer.valueOf(loc[4].substring(1, loc[4].length - 1)).toInt()
        if (useReferenceTracker) {
            if (!isRefErr && !this.isWholeCol && !this.isWholeRow)
                addToRefTracker()
        }
    }


    /**
     * Set the location of this PtgRef.  This takes a location
     * such as {1,2}
     */
    open fun setLocation(rowcol: IntArray) {
        locax = null
        cachedLocation = null
        if (record != null) {
            if (useReferenceTracker)
                removeFromRefTracker()
            rw = rowcol[0]
            col = rowcol[1]
            fRwRel = true    // default
            fColRel = true
            this.updateRecord()
            hashcode = hashCode
            if (useReferenceTracker)
                addToRefTracker()
        } else {
            Logger.logWarn("PtgRef.setLocation() failed: NO record data: $rowcol")
        }
    }

    /**
     * set the location of this PtgRef
     *
     * @param rowcol  int[] rowcol
     * @param bRowRel true if row is relative (i.e. A1 not A$1)
     * @param bColRel true if col is relative (i.e. A1 not $A1)
     */
    fun setLocation(rowcol: IntArray, bRowRel: Boolean, bColRel: Boolean) {
        locax = null
        cachedLocation = null
        if (record != null) {
            if (useReferenceTracker) removeFromRefTracker()
            rw = rowcol[0]
            col = rowcol[1]
            fRwRel = bRowRel
            fColRel = bColRel
            this.updateRecord()
            if (useReferenceTracker)
                addToRefTracker()
        } else {
            Logger.logWarn("PtgRef.setLocation() failed: NO record data: $rowcol")
        }
    }

    /**
     * Updates the record bytes so it can be pulled back out.
     */
    override fun updateRecord() {
        val tmp = ByteArray(5)
        tmp[0] = record[0]
        val brow = ByteTools.cLongToLEBytes(rw)
        System.arraycopy(brow, 0, tmp, 1, 2)
        if (fRwRel) {
            col = (0x8000 or col).toShort().toInt()
        }
        if (fColRel) {
            col = (0x4000 or col).toShort().toInt()
        }
        val bcol = ByteTools.cLongToLEBytes(col)
        if (col == -1) {    // KSC: what excel expects
            bcol[1] = 0
        }
        System.arraycopy(bcol, 0, tmp, 3, 2)

        record = tmp
        if (parentRec != null) {
            if (this.parentRec is Formula)
                (this.parentRec as Formula).updateRecord()
            else if (this.parentRec is Name)
                (this.parentRec as Name).updatePtgs()
        }

        col = col.toShort() and 0x3FFF    //get lower 14 bits which represent the actual column;
    }

    fun changeLocation(newLoc: String, f: Formula): Boolean {
        var newLoc = newLoc
        locax = null
        var ptg: Ptg? = null
        var z = -1
        try {
            z = ExpressionParser.getExpressionLocByPtg(this, f.expression!!)
            ptg = f.expression!![z] as Ptg
        } catch (e: Exception) {

        }

        val unstripped = newLoc

        if (newLoc.indexOf("!") > -1) newLoc = newLoc.substring(newLoc.indexOf("!") + 1)
        if (unstripped.indexOf(":") > 0) { // then either PtgRef3d or PtgArea
            if (unstripped.indexOf("!") > unstripped.indexOf(":")) { // than it's a PtgRef3d or PtgArea3d
                if (unstripped.indexOf(":") != unstripped.lastIndexOf(":")) {    // it's a PtgArea3d ala Sheet1:Sheet3!A1:D1
                    val pta3 = PtgArea3d()
                    pta3.location = unstripped
                    ptg = pta3
                } else { // it's a PtgRef3d ala Sheet1!Sheet3:A1
                    val prd = PtgRef3d()
                    prd.parentRec = f
                    prd.location = unstripped
                    ptg = prd
                }
                // no sheet ref ...
            } else {    // it's a PtgArea3d (according to Excel's Ai recs !!)
                val pta = PtgArea()
                pta.parentRec = f
                pta.location = unstripped
                ptg = pta
            }
        } else if (ptg != null) {
            // it's a single location
            if (unstripped != "") {
                ptg.parentRec = f
                ptg.location = unstripped
            } else {
                ptg = PtgRef3d()
                ptg.parentRec = f
            }
        } else { // ptg is null, create a new one
            ptg = PtgRef()
            ptg.parentRec = f
            ptg.location = unstripped
        }
        if (z != -1)
            f.expression!![z] = ptg    // update expression with new Ptg
        else
            f.expression!!.add(ptg)
        return true
    }

    /**
     * removes this reference from the tracker...
     *
     *
     * used mostly when we've updated the ref and want
     * to re-register it.
     */
    open fun removeFromRefTracker() {
        try {
            if (parentRec != null) {
                parentRec.workBook!!.refTracker!!.removeCellRange(this)
                if (parentRec.opcode == XLSConstants.FORMULA) (parentRec as Formula).setCachedValue(null)
            }
        } catch (ex: Exception) {
            // no need to error here, sometimes this is called before its in Logger.logErr("PtgRef.removeFromRefTracker() failed.", ex);
        }

    }

    /**
     * add this reference to the ReferenceTracker... this
     * is crucial if we are to update this Ptg when cells
     * are changed or added...
     */
    open fun addToRefTracker() {
        //Logger.logInfo("Adding :" + this.toString() + " to tracker");
        try {
            if (parentRec != null)
                parentRec.workBook!!.refTracker!!.addCellRange(this)
        } catch (ex: Exception) {
            Logger.logErr("PtgRef.addToRefTracker() failed.", ex)
        }

    }

    /**
     * update existing tracked ptg with new parent in reference tracker
     *
     * @param parent
     */
    fun updateInRefTracker(parent: XLSRecord?) {
        try {
            if (parent != null)
                parent.workBook!!.refTracker!!.updateInRefTracker(this, parent)
        } catch (ex: Exception) {
            Logger.logErr("updateInRefTracker() failed.", ex)
        }

    }

    /**
     * set the formulaRow and formulaCol for relatively-referenced PtgRefs
     */
    fun setRelativeRowCol() {
        if (fRwRel || fColRel) {
            var opc: Short = 0
            if (parentRec != null)
                opc = parentRec.opcode
            // protocol for shared formulas, conditional formatting, data validity and defined names only (type B cell addresses!)
            if (opc == XLSConstants.SHRFMLA || opc == XLSConstants.DVAL) {
                this.formulaRow = parentRec.rowNumber
                this.formulaCol = parentRec.colNumber.toInt()
            }
        }
    }

    /**
     * set this Ptg to an External Location - used when copying a sheet from another workbook
     *
     * @param f parent formula rec
     */
    fun setExternalReference(externalWorkbook: String) {
        if (this is PtgArea3d) {
            val ptg = this
            var b = parentRec.workBook
            if (b == null)
                b = parentRec.sheet!!.workBook
            val ixti = b!!.externSheet!!.addExternalSheetRef(externalWorkbook, ptg.sheetName)        //20080714 KSC: May not reflect external reference!  this.sheetname);
            ptg.setIxti(ixti)
            if (ptg.firstPtg != null) { // it's not a Ref3d
                ptg.firstPtg!!.updateRecord()
                ptg.lastPtg!!.updateRecord()
            }
            ptg.updateRecord()
        } else if (this is PtgRef3d) {
            var b = parentRec.workBook
            val pr = this
            if (b == null)
                b = parentRec.sheet!!.workBook
            val ixti = b!!.externSheet!!.addExternalSheetRef(externalWorkbook, pr.sheetName)        //20080714 KSC: May not reflect external reference!  this.sheetname);
            pr.setIxti(ixti)
        } else { // TODO: convert to ref3d?
            Logger.logWarn("PtgRef.setExternalReference: unable to convert ref")
        }
    }

    fun setArrayTypeRef() {
        val b = (record[0] or 0x60).toByte()
        record[0] = b
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        if (useReferenceTracker) removeFromRefTracker()
        useReferenceTracker = false
        super.close()
        if (refCell != null && refCell!!.size > 0 && refCell!![0] != null)
        // clear out object references
            (refCell!![0] as XLSRecord).close()
        refCell = null
    }

    companion object {
        /**
         *
         */
        private val serialVersionUID = -7776520933300730470L

        /**
         * given an address string, parse and assign to the appropriate PtgRef-type object
         * <br></br>#REF! 's return either PtgRefErr or PtgRefErr3d
         * <br></br>Ranges return either PtgArea or PtgArea3d
         * <br></br>Single addresses return either PtgRef or PtgRef3d
         * <br></br>NOTE: This method does not extract names embedded within the address string
         *
         * @param address
         * @param parent  parent record to assign the ptg to
         * @return
         */
        fun createPtgRefFromString(address: String, parent: XLSRecord): Ptg {
            try {
                val s = ExcelTools.stripSheetNameFromRange(address)
                val sh1 = s[0]
                val range = s[1]
                val ptg: Ptg
                if (range == null || range == "#REF!" || sh1 != null && sh1 == "#REF") {
                    if (sh1 != null) {
                        val pe3 = PtgRefErr3d()
                        pe3.parentRec = parent
                        pe3.setLocation(s)
                        return pe3
                    } else {
                        val pe = PtgRefErr()
                        pe.parentRec = parent
                        pe.setLocation(s)
                        return pe
                    }
                }
                val bk = parent.workBook


                val sht = "((?:\\\\?+.)*?!)?+"
                val rangeMatch = "(.*(:).*){2,}?"    //matches 2 or more range ops (:'s)
                val opMatch = "(.*([ ,]).*)+"        //matches union or isect op	( " " or ,)
                val m = "$sht(($opMatch)|($rangeMatch))"
                // is address a complex range??
                if (address.matches(m.toRegex()) || range.indexOf("(") > -1) {
                    //NOTE: this can be a MemFunc OR a MemArea --
                    // PtgMemFunc= a NON-CONSTANT cell address, cell range address or cell range list
                    // Whenever one operand of the reference subexpression is a function, a defined name, a 3D
                    // reference, or an external reference (and no error occurs), a PtgMemFunc token is used.
                    // PtgMemArea= constant cell address, cell range address, or cell range list on the same sheet
                    val pmf = PtgMemFunc()
                    pmf.parentRec = parent
                    pmf.location = address    // TODO HANDLE FUNCTION MEMFUNCS ALA OFFSET(x,y,0):OFFSET(x,y,0)
                    ptg = pmf
                } else if (range.indexOf(":") > 0) { // it's a range, either PtgRef3d or PtgArea3d
                    val ops = StringTool.getTokensUsingDelim(range, ":")
                    if (bk!!.getName(ops[0]) != null || bk.getName(ops[1]) != null) {
                        val pmf = PtgMemFunc()
                        pmf.parentRec = parent
                        pmf.location = address
                        ptg = pmf
                    } else if (sh1 != null) {
                        val rc = ExcelTools.getRowColFromString(ops[0])    // see if a wholerow/wholecol ref
                        if (!(ops[0] == ops[1] && rc[0] != -1 && rc[1] != -1)) {
                            val pta = PtgArea3d()
                            pta.parentRec = parent
                            pta.setLocation(s)
                            ptg = pta
                        } else {
                            ptg = PtgRef3d()
                            ptg.setPtgType(Ptg.REFERENCE)
                            ptg.parentRec = parent
                            (ptg as PtgRef).useReferenceTracker = false
                            ptg.setLocation(s)
                            (ptg as PtgRef).useReferenceTracker = true
                            ptg.addToRefTracker()
                        }
                    } else {
                        val pa = PtgArea()
                        pa.parentRec = parent
                        pa.useReferenceTracker = false
                        pa.setLocation(s)
                        pa.useReferenceTracker = true
                        pa.addToRefTracker()
                        ptg = pa
                    }
                } else { // it's a single ref NOT a range e.g. Sheet1!A1
                    if (sh1 != null) {
                        ptg = PtgRef3d()
                        ptg.setPtgType(Ptg.REFERENCE)
                        ptg.parentRec = parent
                        (ptg as PtgRef).useReferenceTracker = false
                        ptg.setLocation(s)
                        (ptg as PtgRef).useReferenceTracker = true
                        ptg.addToRefTracker()
                    } else {
                        val pr = PtgRef()
                        pr.parentRec = parent
                        pr.useReferenceTracker = false
                        pr.setLocation(s)
                        pr.useReferenceTracker = true
                        pr.addToRefTracker()
                        ptg = pr
                    }
                }
                return ptg
            } catch (e: Exception) {    // any error in parsing return a referr -- makes sense!!!
                val pe3 = PtgRefErr3d()
                pe3.parentRec = parent
                return pe3
            }

        }

        fun getHashCode(row: Int, col: Int): Long {
            return (col + (row + 1) * XLSConstants.MAXCOLS).toLong()
        }
    }

}