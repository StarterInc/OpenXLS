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

import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.Stack


/**
 * The Array class describes a formula that was Array-entered into a range of cells.
 *
 *
 * The range of Cells in which the array is entered is defined by the rwFirst, last, colFirst and last fields.
 *
 *
 * The Array record occurs directly after the Formula record for the cell in the upper-left corner of the array -- that is, the cell
 * defined by the rwFirst and colFirst fields.
 *
 *
 * The parsed expression is the array formula -- consisting of Ptgs.
 *
 *
 * You should ignore the chn field when you read the file, it must be 0 if written.
 *
 *
 *
 *
 * OFFSET		NAME			SIZE		CONTENTS
 * -----
 * 4				rwFirst			2			FirstRow of the array
 * 6				fwLast			2			Last Row of the array
 * 8				colFirst		1			First Column of the array
 * 9				colLast			1			Last Column of the array
 * 10			grbit			2			Option Flags
 * 12			chn				4			set to 0, ignore
 * 16			cce				2			Length of the parsed expression
 * 18			rgce			var		Parsed Expression
 *
 *
 * grbit fields
 * bit 1 = fAlwaysCalc - always calc the formula
 * bit 2 = fCalcOnLoad - calc formula on load
 */
class Array : io.starter.formats.XLS.XLSRecord() {
    private var rwFirst: Short = 0
    private var rwLast: Short = 0
    private var colFirst: Short = 0
    private var colLast: Short = 0
    private var cce: Short = 0
    private var grbit: Short = 0
    private var rgce: ByteArray? = null
    /**
     * return parent formula
     *
     * @return
     */
    /**
     * link this shared formula to it's parent formula
     */
    var parentRec: Formula? = null

    var firstRow: Int
        get() = rwFirst.toInt()
        set(i) {
            rwFirst = i.toShort()
        }

    var lastRow: Int
        get() = rwLast.toInt()
        set(i) {
            rwLast = i.toShort()
        }

    var firstCol: Int
        get() = colFirst.toInt()
        set(i) {
            colFirst = i.toByte().toShort()
        }

    var lastCol: Int
        get() = colLast.toInt()
        set(i) {
            colLast = i.toByte().toShort()
        }

    /*
     * For getRow() and getCol() we are going to return the upper-right hand
     * location of the sharedformula.  This should be the same location referenced
     * in the PTGExp associated with these formulas as well
     */
    override var rowNumber: Int
        get() = firstRow
        set

    val col: Int
        get() = firstCol

    /**
     * allow access to expression
     *
     * @return
     */
    var expression: Stack<*>? = null
        private set

    /**
     * Returns the top left location of the array, used for identifying which array goes with what formula.
     */
    val parentLocation: String
        get() {
            val `in` = IntArray(2)
            `in`[0] = colFirst.toInt()
            `in`[1] = rwFirst.toInt()
            return ExcelTools.formatLocation(`in`)
        }

    /**
     * return the string representation of the array formula
     */
    val formulaString: String
        get() {
            val expressString = FormulaParser.getExpressionString(expression!!)
            return if ("" != expressString) expressString.substring(1) else ""
        }

    /**
     * return the cells referenced by this array in string range form
     *
     * @return
     */
    val arrayRefs: String
        get() {
            val rc = IntArray(2)
            rc[0] = rwFirst.toInt()
            rc[1] = colFirst.toInt()
            val rowcol1 = ExcelTools.formatLocation(rc)
            rc[0] = rwLast.toInt()
            rc[1] = colLast.toInt()
            val rowcol2 = ExcelTools.formatLocation(rc)
            return "$rowcol1:$rowcol2"
        }

    override fun init() {
        super.init()
        this.opcode = XLSConstants.ARRAY
        // NO NOT TRUE!!!!!        super.setIsValueForCell(true); // this ensures that it gets picked up by the Formula...
        // will get "picked up" by formula due to WorkBook.addRecord code
        // isValueForCell will DELETE THE CELL if remove/change NOT WHAT WE WANT HERE
        rwFirst = io.starter.toolkit.ByteTools.readInt(this.getByteAt(0), this.getByteAt(1), 0.toByte(), 0.toByte()).toShort()
        rwLast = io.starter.toolkit.ByteTools.readInt(this.getByteAt(2), this.getByteAt(3), 0.toByte(), 0.toByte()).toShort()

        colFirst = ByteTools.readUnsignedShort(this.getByteAt(4), 0.toByte()).toShort()
        colLast = ByteTools.readUnsignedShort(this.getByteAt(5), 0.toByte()).toShort()
        grbit = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        cce = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        rgce = this.getBytesAt(14, this.cce.toInt())

        expression = ExpressionParser.parseExpression(rgce, this)

        // for PtgArray and PtgMemAreas, have Constant Array Tokens following expression
        // this length is NOT included in cce
        var posExtraData = cce + 14
        val len = this.getData()!!.size
        for (i in expression!!.indices) {
            if (expression!![i] is PtgArray) {
                try {
                    val b = ByteArray(len - posExtraData)
                    System.arraycopy(this.getData()!!, posExtraData, b, 0, len - posExtraData)
                    val pa = expression!![i] as PtgArray
                    posExtraData += pa.setArrVals(b)
                } catch (e: Exception) {
                    Logger.logErr("Array: error getting Constants " + e.localizedMessage)
                }

            } else if (expression!![i] is PtgMemArea) {
                try {
                    val pm = expression!![i] as PtgMemArea
                    val b = ByteArray(pm.getnTokens() + 8)
                    System.arraycopy(pm.record, 0, b, 0, 7)
                    System.arraycopy(this.getData()!!, posExtraData, b, 7, pm.getnTokens())
                    pm.init(b)
                    posExtraData += pm.getnTokens()
                } catch (e: Exception) {
                    Logger.logErr("Array: error getting memarea constants $e")
                }

            }
        }
        if (DEBUGLEVEL > XLSConstants.DEBUG_LOW)
            Logger.logInfo("Array encountered at: " + this.wkbook!!.lastbound!!.sheetName + "!" + ExcelTools.getAlphaVal(colFirst.toInt()) + (rwFirst + 1) + ":" + ExcelTools.getAlphaVal(colLast.toInt()) + (rwLast + 1))
    }

    /**
     * Associate this record with a worksheet.
     * init array refs as well
     */
    override fun setSheet(b: Sheet?) {
        this.worksheet = b
        // add to array formula references since this is the parent
        if (expression != null) { // it's been initted
            val loc = ExcelTools.formatLocation(intArrayOf(rwFirst.toInt(), colFirst.toInt()))    // this formula address == array formula references for OOXML usage
            val ref = ExcelTools.formatRangeRowCol(intArrayOf(rwFirst.toInt(), colFirst.toInt(), rwLast.toInt(), colLast.toInt()))
            (b as Boundsheet).addParentArrayRef(loc, ref)        // formula address, array formula references OOXML usage
        }

    }

    /**
     * init Array record from formula string
     *
     * @param fmla
     */
    fun init(fmla: String, rw: Int, col: Int) {
        var fmla = fmla
        // NO NOT TRUE!!!!!        super.setIsValueForCell(true); // this ensures that it gets picked up by the Formula...
        // will get "picked up" by formula due to WorkBook.addRecord code
        // isValueForCell will DELETE THE CELL if remove/change NOT WHAT WE WANT HERE
        // TODO: ever a case of rwFirst!=rwLast, colFirst!=colLast ?????
        this.opcode = XLSConstants.ARRAY
        rwFirst = rw.toShort()
        rwLast = rw.toShort()
        colFirst = col.toByte().toShort()
        colLast = col.toByte().toShort()
        grbit = 0x2    // calc on load = default	 20090824 KSC	[BugTracker 2683]
        fmla = fmla.substring(1, fmla.length - 1)     // parse formula string and add stack to Array record
        val newptgs = FormulaParser.getPtgsFromFormulaString(this, fmla)
        expression = newptgs
        this.updateRecord()

    }

    /*  Update the record byte array with the modified ptg records
     */
    fun updateRecord() {
        // first, get expression length
        cce = 0        // sum up expression (Ptg) records
        for (i in expression!!.indices) {
            val ptg = expression!!.elementAt(i) as Ptg
            cce += ptg.record.size.toShort()
            if (ptg is PtgRef) {
                ptg.setArrayTypeRef()
            }
        }

        var newdata = ByteArray(14 + cce)        // total record data (not including extra data, if any)

        var pos = 0
        // 20090824 KSC: [BugTracker 2683]
        // use setOpcode rather than setting 1st byte as it's a record not a ptg		newdata[0]= (byte) XLSConstants.ARRAY;			pos++;
        this.opcode = XLSConstants.ARRAY
        System.arraycopy(ByteTools.shortToLEBytes(rwFirst), 0, newdata, pos, 2)
        pos += 2
        System.arraycopy(ByteTools.shortToLEBytes(rwLast), 0, newdata, pos, 2)
        pos += 2
        newdata[pos++] = colFirst.toByte()
        newdata[pos++] = colLast.toByte()
        System.arraycopy(ByteTools.shortToLEBytes(grbit), 0, newdata, pos, 2)    // 20090824 KSC: Added [BugTracker 2683]

        pos = 12
        System.arraycopy(ByteTools.shortToLEBytes(cce), 0, newdata, pos, 2)
        pos += 2

        // expression
        rgce = ByteArray(cce)                    // expression record data
        pos = 0
        var arraybytes = ByteArray(0)
        for (i in expression!!.indices) {
            val ptg = expression!!.elementAt(i) as Ptg
            // trap extra data after expression (not included in cce count)
            if (ptg is PtgArray) {
                arraybytes = ByteTools.append(ptg.postRecord, arraybytes)
            } else if (ptg is PtgMemArea) {
                arraybytes = ByteTools.append(ptg.postRecord, arraybytes)
            }
            val b = ptg.record
            System.arraycopy(b, 0, rgce!!, pos, ptg.length)
            pos = pos + ptg.length
        }
        System.arraycopy(rgce!!, 0, newdata, 14, cce.toInt())
        newdata = ByteTools.append(arraybytes, newdata)
        this.setData(newdata)
    }

    /**
     * Determines whether the address is part of the array Formula range
     */
    internal fun isInRange(addr: String): Boolean {
        return io.starter.OpenXLS.ExcelTools.isInRange(addr, rwFirst.toInt(), rwLast.toInt(), colFirst.toInt(), colLast.toInt())
    }


    fun getValue(pxp: PtgExp): Any {
        //try{
        return FormulaCalculator.calculateFormula(expression!!)
        //}catch(FunctionNotSupportedException e){
        //Logger.logWarn("Array.getValue() failed: " + e);
        //return null;
        //	}
    }

    companion object {

        private val serialVersionUID = -7316545663448065447L
    }
}