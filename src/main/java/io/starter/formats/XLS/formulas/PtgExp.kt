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

import java.util.Stack

/**
 * ptgExp indicates an Array Formula or Shared Formula
 *
 *
 * When ptgExp occurs in a formula, it's the only token in the formula.
 * this indicates that the cell containing the formula
 * is part of an array or opartof a shared formula.
 * The actual formula is found in an array record.
 *
 *
 * The value for ptgExp consists of the row and the column of the
 * upper-left corner of the array formula.
 *
 * @see Ptg
 *
 * @see Formula
 *
 * @see Array
 *
 * @see Shrfmla
 */
class PtgExp : GenericPtg(), Ptg {
    var rwFirst: Int = 0
        internal set
    var colFirst: Int = 0

    override val isControl: Boolean
        get() = true

    override val isStandAloneOperator: Boolean
        get() = true

    override val length: Int
        get() = Ptg.PTG_EXP_LENGTH

    /**
     * Returns the location this PtgExp points to.
     */
    val referent: String
        get() = ExcelTools.formatLocation(intArrayOf(rwFirst, colFirst))

    /**
     * Looks up into it's parent shared formula, and returns
     * the expression as if it were a regular formula.
     *
     * @return converted Calculation Expression
     */
    //			throw new UnsupportedOperationException (
    //					"Shared formulas must be instantiated for calculation");
    // if it's an array formula, return ptg's as well
    // this.getParentRec().getSheet().getArrayFormula(getParentLocation());
    val convertedExpression: Array<Ptg>
        get() {
            val f = this.parentRec as Formula
            if (f.isSharedFormula) {
                val expression = f.shared!!.instantiate(f)
                val retPtg = arrayOfNulls<Ptg>(expression.size)
                for (i in expression.indices) {
                    val p = expression[i] as Ptg
                    retPtg[i] = p
                }
                return retPtg
            } else {
                val a = f.internalRecords[0] as Array
                val calcStack = a.expression
                val retPtg = arrayOfNulls<Ptg>(calcStack!!.size)
                for (i in calcStack!!.indices) {
                    val p = calcStack!!.get(i) as Ptg
                    retPtg[i] = p
                }
                return retPtg
            }
        }

    override//			throw new UnsupportedOperationException (
    //					"Shared formulas must be instantiated for calculation");
    // it's part of an array formula but not the parent
    val value: Any?
        get() {
            var o: Any? = null
            val f = this.parentRec as Formula
            if (f.isSharedFormula) {
                o = FormulaCalculator.calculateFormula(f.shared!!.instantiate(f))
            } else {
                var r: Any? = null
                if (f.internalRecords.size > 0)
                    r = f.internalRecords[0]
                else {
                    r = this.parentRec!!.sheet!!.getArrayFormula(referent)
                }
                if (r is Array) {
                    val arr = r as Array?
                    o = arr!!.getValue(this)
                } else if (r is StringRec) {
                    o = r.stringVal
                }
            }
            return o
        }

    /**
     * return the location of this PtgExp
     * 20060302 KSC
     */
    /**
     * setLocation vars from address string
     *
     * @param s String address
     */
    override var location: String?
        @Throws(FormulaNotFoundException::class)
        get() {
            var s = ""
            try {
                s = this.parentRec!!.cellAddress
                s = this.parentRec!!.sheet!!.sheetName + "!" + s
            } catch (e: Exception) {

            }

            return s
        }
        set(s) {
            val rc = ExcelTools.getRowColFromString(s)
            rwFirst = rc[0]
            colFirst = rc[1]
            updateRecord()
        }

    /**
     * return the human-readable String representation of the linked shared formula
     */
    override// Object o= ((Formula) this.getParentRec()).getInternalRecords().get(0);	PARENT REC of ARRAY or SHRFMLA is determined by referent (record) NOT necessarily same as actual Parent Rec
    //	    		Formula pr= (Formula) sht.getCell(this.getReferent());
    //if this is a shared formula the attached string is the RESULT, not the formula string itself
    // subsequent formulas use same shared formula rec so find
    val string: String?
        get() {
            try {
                try {
                    val sht = this.parentRec!!.sheet
                    val f = this.parentRec as Formula
                    val o: Any?
                    if (f.isSharedFormula) {
                        o = FormulaParser.getExpressionString(f.shared!!.instantiate(f))
                        return if (o != null && o.toString().startsWith("=")) o.toString().substring(1) else o!!.toString()
                    }
                    val pr = sht!!.getCell(this.referent) as Formula
                    o = pr.internalRecords[0]
                    if (o is Array) {
                        val a = o as Array?
                        return a!!.formulaString
                    } else if (o is StringRec) {
                        if ((this.parentRec as Formula).isSharedFormula) {
                            throw IndexOutOfBoundsException("parse it")
                        }
                        val s = o as StringRec?
                        return s!!.stringVal
                    }
                } catch (e: IndexOutOfBoundsException) {
                    throw UnsupportedOperationException(
                            "Shared formulas must be instantiated for calculation")
                }

            } catch (e: Exception) {
                return "Array-Entered or Shared Formula"
            }

            return "Array-Entered or Shared Formula"
        }

    /**
     * init from row, col
     *
     * @param row
     * @param col
     */
    fun init(row: Int, col: Int) {
        val r = ByteTools.shortToLEBytes(row.toShort())
        val c = ByteTools.shortToLEBytes(col.toShort())
        record = byteArrayOf(0x1, r[0], r[1], c[0], c[1])
        opcode = record[0]
        this.populateVals()
        //this.addToReferenceTracker();
    }

    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
        //this.addToReferenceTracker();
    }

    private fun populateVals() {
        rwFirst = readRow(record[1], record[2])
        colFirst = ByteTools.readShort(record[3].toInt(), record[4].toInt()).toInt()
    }

    override fun calculatePtg(parsething: Array<Ptg>): Ptg? {
        var o: Any? = null
        val f = this.parentRec as Formula
        if (f.isSharedFormula) {
            o = FormulaCalculator.calculateFormula(f.shared!!.instantiate(f))
        } else {
            var r: Any? = null
            if (f.internalRecords.size > 0)
                r = f.internalRecords[0]
            else {    // it's part of an array formula but not the parent
                r = this.parentRec!!.sheet!!.getArrayFormula(referent)
            }
            if (r is Array) {
                val arr = r as Array?
                o = arr!!.getValue(this)
            } else if (r is StringRec) {
                o = r.stringVal
            } else
            // should never happen
                throw UnsupportedOperationException(
                        "Expected records parsing Formula were not present")
        }
        var p: Ptg? = null
        // conversion isn't necessary
        //		try{
        if (o is Int)
            return PtgInt(o.toInt())
        else if (o is Double)
            return PtgNumber(o.toDouble())//			Double d = new Double(o.toString());
        //p = new PtgNumber(d.doubleValue());
        //		}catch(NumberFormatException e){
        if (o!!.toString().equals("true", ignoreCase = true) || o.toString().equals("false", ignoreCase = true))
            p = PtgBool(o.toString().equals("true", ignoreCase = true))
        else
            p = PtgStr(o.toString())
        //		}
        return p
    }

    /**
     * updateRecord from local rwFirst and colFirst values
     *
     * @see
     */
    override fun updateRecord() {
        System.arraycopy(ByteTools.shortToLEBytes(rwFirst.toShort()), 0, record, 1, 2)
        System.arraycopy(ByteTools.shortToLEBytes(colFirst.toShort()), 0, record, 3, 2)
    }

    fun setRowFirst(r: Int) {
        this.rwFirst = r
    }

    override fun toString(): String {
        return "PtgExp: Parent Formula at [$rwFirst,$colFirst]"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2150560716287810448L
    }

}