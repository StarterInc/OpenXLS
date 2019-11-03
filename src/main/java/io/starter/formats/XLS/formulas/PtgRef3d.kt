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
import io.starter.toolkit.ByteTools
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger


/**
 * A BiffRec range spanning 3rd dimension of WorkSheets.
 * `
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 0       ixti       2        Index to Externsheet Sheet Record
 * 2       row         2       The row
 * 4       grCol       2       The col, or the col offset (see next table)
 *
 * the low-order 8 bytes store the col numbers.  The 2 MSBs specify whether the row
 * and col refs are relative or absolute.
 *
 * bits    mask        name    content
 * ---
 * 15      8000h       fRwRel  =1 if the row is relative, 0 if absolute
 * 14      4000h       fColRel =1 if the col is relative, 0 if absolute
 * 13-8    3F00h       (reserved)
 * 7-0     00FFh       col     the col number or col offset (0-based)
 *
 * For 3D references, the tokens contain a negative EXTERNSHEET index, indicating a reference into the own workbook.
 * The absolute value is the one-based index of the EXTERNSHEET record that contains the name of the first sheet. The
 * tokens additionally contain absolute indexes of the first and last referenced sheet. These indexes are independent of the
 * EXTERNSHEET record list. If the referenced sheets do not exist anymore, these indexes contain the value FFFFH (3D
 * reference to a deleted sheet), and an EXTERNSHEET record with the special name <04H> (own document) is used.
 * Each external reference contains the positive one-based index to an EXTERNSHEET record containing the URL of the
 * external document and the name of the sheet used. The sheet index fields of the tokens are not used.
 *
 * @see Ptg
 *
 * @see Formula
</pre> */
open class PtgRef3d : PtgRef, Ptg, IxtiListener {

    internal var quoted = false
    var ixti: Short = 0


    override var parentRec: XLSRecord
        get
        set(r) {
            super.parentRec = r
        }

    /**
     * returns true if this PtgRef3d's ixti refers to an external sheet reference
     *
     * @return
     */
    val isExternalLink: Boolean
        get() {
            try {
                return parentRec.workBook!!.externSheet!!.getIsExternalLink(ixti.toInt())
            } catch (e: Exception) {
                return false
            }

        }


    override val length: Int
        get() = Ptg.PTG_REF3D_LENGTH

    override val isOperand: Boolean
        get() = true

    override val isReference: Boolean
        get() = true

    /**
     * Returns the location of the Ptg as a string, including sheet name
     */
    /**
     * Set Location can take either a local page address (ie A54) or
     * a reference to a page and location(ie Sheet2!A22).  It then changes
     * the location reference of the Ptg.
     *
     * @see io.starter.formats.XLS.formulas.Ptg.setLocation
     */
    override// doesn't have a sheet ref
    // NOTE: Our tests error when PtgRefs have fully qualified range syntax
    // PtgRef does not have ixti
    var location: String?
        get() {
            val ret = super.location
            if (ret!!.indexOf("!") == -1) {
                if (sheetname == null)
                    sheetname = this.sheetName
                if (this.sheetname != null) {
                    if (sheetname == "#REF!")
                        return sheetname!! + ret!!
                    sheetname = GenericPtg.qualifySheetname(sheetname)
                    return "$sheetname!$ret"
                }
            }
            return ret
        }
        set(address) {
            val s = ExcelTools.stripSheetNameFromRange(address)
            setLocation(s)
        }


    // 20080303 KSC: Catch Unresolved External refs
    val sheet: Boundsheet?
        get() {
            if (parentRec != null) {
                val wb = parentRec.workBook
                if (wb != null && wb.externSheet != null) {
                    val bsa = wb.externSheet!!.getBoundSheets(this.ixti.toInt())
                    if (bsa == null || bsa[0] == null) {
                        if (parentRec is Formula)
                            Logger.logErr("PtgRef3d.getSheet: Unresolved External Worksheet in Formula " + parentRec.cellAddressWithSheet)
                        else if (parentRec is Name)
                            Logger.logErr("PtgRef3d.getSheet: Unresolved External Worksheet in Name " + (parentRec as Name).name)
                        else
                            Logger.logErr("PtgRef3d.getSheet: Unresolved External Worksheet for " + parentRec.cellAddressWithSheet)
                        return null
                    }
                    return bsa[0]
                }
            }
            return null
        }

    /**
     * return the sheet name for this 3d reference
     */
    override// 20080306 KSC: new way is to get sheet names rather than sheets as can be external refs
    var sheetName: String?
        get() {
            if (this.sheetname == null) {
                if (parentRec != null) {
                    val wb = parentRec.workBook
                    if (wb != null && wb.externSheet != null) {
                        val sheets = wb.externSheet!!.getBoundSheetNames(this.ixti.toInt())
                        if (sheets != null && sheets[0] != null)
                            sheetname = sheets[0]
                    }
                }
            }
            return sheetname
        }
        set(value: String?) {
            super.sheetName = value
        }


    /**
     * @return Returns the refCell.
     */
    override val refCells: Array<BiffRec>?
        get() {
            if (sheetname == null) sheetname = this.sheetName
            refCell = super.refCells
            return refCell
        }

    /**
     * PtgRef's have no sub-compnents
     */
    override// only one
    val components: Array<Ptg>?
        get() = null

    override fun addListener() {
        try {
            parentRec.workBook!!.externSheet!!.addPtgListener(this)
        } catch (e: Exception) {
            // no need to output here.  NullPointer occurs when a ref has an invalid ixti, such as when a sheet was removed  Worksheet exception could never really happen.
        }

    }


    /**
     * @return Returns the ixti.
     */
    override fun getIxti(): Short {    // only valid for 3d refs
        return ixti
    }
    //true is relative, false is absolute

    /**
     * 0x3A Reference class token: The reference address itself, independent of the cell contents.
     * • 0x5A Value class token: A value (a constant, a function result, or one specific value from a dereferenced cell range).
     * • 0x7A Array class token: An array of values (array of constant values, an array function result, or all values of a cell range).
     */
    constructor() {
        record = ByteArray(Ptg.PTG_REF3D_LENGTH)
        opcode = 0x5A  // id varies with type of token see above and setPtgType below
        record[0] = opcode // ""
        this.is3dRef = true
    }

    /**
     * set the Ptg Id type to one of:
     * VALUE, REFERENCE or Array
     * <br></br>The Ptg type is important for certain
     * functions which require a specific type of operand
     */
    override fun setPtgType(type: Short) {
        when (type) {
            Ptg.VALUE -> opcode = 0x5A
            Ptg.REFERENCE -> opcode = 0x3A
            Ptg.ARRAY -> opcode = 0x7A
        }
        record[0] = opcode
    }

    constructor(addToRefTracker: Boolean) {
        this.useReferenceTracker = addToRefTracker
        opcode = 0x5A  // TODO: id varies with type of token see above
        record[0] = opcode // ""
        this.is3dRef = true

    }

    constructor(addr: String, _ixti: Short) : this() {
        location = addr
        this.is3dRef = true
    }


    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        populateVals()
    }

    /**
     * get the worksheet that this ref is on
     * for some reason this seems to be backwards in Ref3d
     */
    fun getSheet(b: WorkBook): Boundsheet? {
        val bsa = b.externSheet!!.getBoundSheets(ixti.toInt())
        if (bsa != null && bsa[0] == null) { // 20080303 KSC: catch error
            // try harder...
            if (parentRec.sheet != null) {
                return parentRec.sheet // sheetless names belong to parent rec
            } else {
                if (b.factory!!.debugLevel > 1)
                // 20080925 KSC
                    Logger.logErr("PtgRef3d.getSheet: Unresolved External or Deleted Sheet Reference Found") // [BUGTRACKER 1836] Claritas extenXLS22677.rec (Deleted Sheet/Named Range causes errant value in B3)
                return null    //20080805 KSC: Don't just return the 1st sheet, may be wrong, deleted, etc!
            }
        } else if (bsa == null)
            return null
        return bsa[0]
    }

    /**
     * set Ptg to parsed location
     *
     * @param loc String[] sheet1, range, sheet2, exref1, exref2
     */
    override fun setLocation(s: Array<String>) {
        if (useReferenceTracker && !isRefErr)
            this.parentRec.workBook!!.refTracker!!.removeCellRange(this)
        sheetname = null
        if (s[0] != null) {
            sheetname = s[0]
        } else {
            try {    // if not provided, assume that parent rec sheet is correct
                sheetname = this.parentRec.sheet!!.sheetName
            } catch (e: NullPointerException) {
            }

        }
        var loc = s[1]
        if (sheetname != null) {
            loc = "$sheetname!$loc"        // loc uses quoted vers of sheet
            if (sheetname!!.indexOf("'") == 0) {
                sheetname = sheetname!!.substring(1, sheetname!!.length - 1)
                quoted = true
            }
        }
        if (sheetname != null) {
            var xsht: Externsheet? = null
            var b = parentRec.workBook
            if (b == null)
                b = parentRec.sheet!!.workBook
            try {

                val boundnum = b!!.getWorkSheetByName(sheetname).sheetNum
                xsht = b.externSheet
                try {
                    val xloc = xsht!!.insertLocation(boundnum, boundnum)
                    setIxti(xloc.toShort())
                } catch (e: Exception) {
                    Logger.logWarn("PtgRef3d.setLocation could not update Externsheet:$e")
                }

            } catch (e: WorkSheetNotFoundException) {
                try {
                    xsht = b!!.externSheet
                    val boundnum = xsht!!.getXtiReference(s[0], s[0])
                    if (boundnum == -1) {    // can't resolve
                        this.setIxti(xsht.insertLocation(boundnum, boundnum).toShort())
                    } else {
                        this.setIxti(boundnum.toShort())
                    }
                } catch (ex: Exception) {
                }

            }

        }
        super.setLocation(s)
    }

    /**
     * Throw this data into a ptgref's
     */
    public override fun populateVals() {
        ixti = ByteTools.readShort(record[1].toInt(), record[2].toInt())
        this.sheetname = this.sheetName

        rw = readRow(record[3], record[4])
        val column = ByteTools.readShort(record[5].toInt(), record[6].toInt())
        // is the Row relative?
        fRwRel = column and 0x8000 == 0x8000
        // is the Column relative?
        fColRel = column and 0x4000 == 0x4000
        col = (column and 0x3fff).toShort().toInt()
        setRelativeRowCol()  // set formulaRow/Col for relative references if necessary
        this.intLocation    // sets the wholeRow and/or wholeCol flag for certain refs
        this.hashcode = super.hashCode
    }


    /**
     * Set the location of this PtgRef.  This takes a location
     * such as "a14"
     */
    fun setLocation(address: String, ix: Short) {
        ixti = ix
        val s = ExcelTools.stripSheetNameFromRange(address)
        this.setLocation(s)
    }

    override fun toString(): String? {
        var ret: String? = ""
        try {
            ret = location
            if (ret!!.indexOf("!") == -1 && sheetname != null) { // prepend sheetname
                if (sheetname!!.indexOf(' ') == -1 && sheetname!![0] != '\'')
                // 20081211 KSC: Sheet names with spaces must have surrounding quotes
                    ret = "$sheetname!$ret"
                else
                    ret = "'$sheetname'!$ret"
            }
        } catch (ex: Exception) {
            Logger.logErr("PtgRef3d.toString() failed", ex)
        }

        return ret
    }

    override fun setIxti(ixf: Short) {
        if (ixti != ixf) {
            ixti = ixf
            // this seems to be only one byte...
            if (record != null) {
                record[1] = ixf.toByte()
            }
            updateRecord()
        }
    }

    /**
     * Change the sheet reference to the passed in boundsheet
     *
     * @see io.starter.formats.XLS.formulas.PtgArea3d.setReferencedSheet
     */
    fun setReferencedSheet(b: Boundsheet) {
        val boundnum = b.sheetNum
        val xsht = b.workBook!!.getExternSheet(true)
        //TODO: add handling for multi-sheet reference.  Already handled in externsheet
        try {
            val xloc = xsht!!.insertLocation(boundnum, boundnum)
            setIxti(xloc.toShort())
            this.sheetname = null    // 20100218 KSC: RESET
            this.sheetName
            locax = null
        } catch (e: WorkSheetNotFoundException) {
            Logger.logErr("Unable to set referenced sheet in PtgRef3d $e")
        }

    }

    /**
     * Updates the record bytes so it can be pulled back out.
     */
    override fun updateRecord() {
        val tmp = ByteArray(Ptg.PTG_REF3D_LENGTH)
        tmp[0] = record[0]
        val ix = ByteTools.shortToLEBytes(ixti)
        System.arraycopy(ix, 0, tmp, 1, 2)
        val brow = ByteTools.cLongToLEBytes(rw)
        System.arraycopy(brow, 0, tmp, 3, 2)
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
        System.arraycopy(bcol, 0, tmp, 5, 2)
        record = tmp
        if (parentRec != null) {
            if (this.parentRec is Formula)
                (this.parentRec as Formula).updateRecord()
            else if (this.parentRec is Name)
                (this.parentRec as Name).updatePtgs()
        }

        col = col.toShort() and 0x3FFF    //get lower 14 bits which represent the actual column;
    }

    /**
     * return the ptg components for a certain column within a ptgArea()
     *
     * @param colNum
     * @return all Ptg's within colNum
     */
    fun getColComponents(colNum: Int): Array<Ptg> {
        val v = FastAddVector()
        val x = this.intLocation
        if (x!![1] == colNum) v.add(this)
        val pref = arrayOfNulls<PtgRef>(v.size)
        v.toTypedArray()
        return pref
    }

    companion object {

        private val serialVersionUID = -441121385905948168L
    }
}