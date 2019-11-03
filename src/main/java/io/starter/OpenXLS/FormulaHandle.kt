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
package io.starter.OpenXLS

import io.starter.formats.XLS.*
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.Logger


/**
 * Formula Handle allows for manipulation of Formulas within a WorkBook.
 *
 * @see WorkBookHandle
 *
 * @see WorkSheetHandle
 *
 * @see CellHandle
 */
class FormulaHandle
/**
 * Create a new FormulaHandle from an Excel Formula
 *
 * @param Formula - the formula to create a handle for.
 */
(f: Formula, private val bk: WorkBook) {

    // 20090120 KSC: should be detected auto public void setIsArrayFormula(boolean b) { form.setIsArrayFormula(b); }

    /**
     * returns the low-level formula rec for this Formulahandle
     *
     * @return
     */
    var formulaRec: Formula? = null
        private set


    /**
     * Returns the cell Address of the formula
     */
    val cellAddress: String
        get() = formulaRec!!.cellAddress

    /**
     * Returns the Human-Readable Formula String
     *
     * @return String the Formula in Human-readable format
     */
    val formulaString: String
        get() = formulaRec!!.formulaString

    /**
     * If the Formula evaluates to a String, return
     * the value as a String.
     *
     * @return String - value of the Formula if stored as a String.
     */
    //this.form.init();
    val stringVal: String?
        @Throws(FunctionNotSupportedException::class)
        get() = formulaRec!!.stringVal

    /**
     * Return the value of the Formula
     *
     * @return Object - value of the Formula
     */
    val `val`: Any?
        @Throws(FunctionNotSupportedException::class)
        get() = sanitizeValue(formulaRec!!.calculateFormula())

    /**
     * If the Formula evaluates to a float, return
     * the value as an float.
     *
     *
     * If the workbook level flag CALCULATE_EXPLICIT is set
     * then the cached value of the formula (if available) will be returned,
     * otherwise the latest calculated value will be returned
     *
     * @return float - value of the Formula if available as a float.  If the
     * value cannot be returned as a float NaN will be returned.
     */
    val floatVal: Float
        @Throws(FunctionNotSupportedException::class)
        get() = formulaRec!!.floatVal

    /**
     * If the Formula evaluates to a double, return
     * the value as an double.
     *
     *
     * If the workbook level flag CALCULATE_EXPLICIT is set
     * then the cached value of the formula (if available) will be returned,
     * otherwise the latest calculated value will be returned
     *
     * @return double - value of the Formula if available as a double.  If the
     * value cannot be returned as a double NaN will be returned.
     */
    val doubleVal: Double
        @Throws(FunctionNotSupportedException::class)
        get() = formulaRec!!.dblVal

    /**
     * If the Formula evaluates to an int, return
     * the value as an int.
     *
     *
     * If the workbook level flag CALCULATE_EXPLICIT is set
     * then the cached value of the formula (if available) will be returned,
     * otherwise the latest calculated value will be returned
     *
     * @return int - value of the Formula if available as a int.  If the value returned can not be
     * represented by an int or is a float/double with a non-zero mantissa a runtime NumberFormatException
     * will be thrown
     */
    val intVal: Int
        @Throws(FunctionNotSupportedException::class)
        get() = formulaRec!!.intVal

    /**
     * get CellRange strings referenced by this formula
     *
     * @return
     * @throws FormulaNotFoundException
     */
    // need sheetname along with address; to ensure, must use explicit method:
    //			ret[x]=locptgs[x].getTextString();
    //avoid NumberFormatExceptions on parsing missing Named Ranges
    val ranges: Array<String>
        @Throws(FormulaNotFoundException::class)
        get() {
            val locptgs = formulaRec!!.cellRangePtgs
            val ret = arrayOfNulls<String>(locptgs.size)
            for (x in locptgs.indices) {
                try {
                    ret[x] = (locptgs[x] as PtgRef).locationWithSheet
                } catch (e: Exception) {
                    if (locptgs[x] is PtgName) {
                        ret[x] = locptgs[x].location
                    } else
                        ret[x] = locptgs[x].textString
                }

            }
            return ret
        }

    /**
     * Initialize CellRanges referenced by this formula
     *
     * @return
     * @throws FormulaNotFoundException
     */
    //
    val cellRanges: Array<CellRange>
        @Throws(FormulaNotFoundException::class)
        get() {
            val crstrs = this.ranges
            val crs = arrayOfNulls<CellRange>(crstrs.size)
            for (x in crs.indices) {
                crs[x] = CellRange(crstrs[x], bk, true)
                try {
                    crs[x].init()
                } catch (e: Exception) {
                }

            }
            return crs
        }

    /**
     * return truth of "this formula is shared"
     *
     * @return boolean
     */
    val isSharedFormula: Boolean
        get() = this.formulaRec!!.isSharedFormula

    val isArrayFormula: Boolean
        get() = formulaRec!!.isArrayFormula

    /**
     * Utility method to determine if the calculation works out to an error value.
     *
     *
     * The excel values that will cause this to be true are
     * #VALUE!, #N/A, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
     *
     * @return
     */
    val isErrorValue: Boolean
        get() = formulaRec!!.calculateFormula() is CalculationException

    /**
     * return the "Calculate Always" setting for this formula
     * used for formulas that always need calculating such as TODAY
     *
     * @return
     */
    /**
     * set the "Calculate Always setting for this formula
     * used for formulas that always need calculating such as TODAY
     *
     * @param fAlwaysCalc
     */
    var calcAlways: Boolean
        get() = formulaRec!!.calcAlways
        set(fAlwaysCalc) {
            formulaRec!!.calcAlways = fAlwaysCalc
        }

    /**
     * generate the OOXML necessary to describe this formula
     * OOXML element <f>
     *
     * @return
    </f> */
    // TODO: Deal with External References ... dataTables
    // common possible attributes:
    // aca= always calculate array, bx=name to assign formula to, ca=calculate cell, r1=data table cell1,
    // t=formula type shared, array, dataTable or normal=default
    // must have type of formula result
    // means cache was cleared (in all cases???) MUST recalc
    // Handle attributes for special cached values
    // TODO: how can we strip non-ascii? What about Japanese cell text?  ans:  has to be XML-compliant is all ...
    // handle non-standard xml chars -- ummm what about Japanese? -- it's all ok
    // ignore =
    // array formulas
    // it's the parent
    // remove "{= }"
    // it's part of a multi-cell array formula therefore DO NOT add array info here
    // only output value info
    // TODO: FINISH 00XML SHARED FORMULAS
    // TODO: need si= shared formula index; when referencing (after shared formula is defined) don't need to include "fs" just the si= ****
    // 20091022 KSC: Shared Formulas do not work 2003->2007
    //ooxml.append(" t=\"shared\" ref=\"" + this.getFormulaRec().getSharedFormula().getCellRange() + "\"");
    // can happen if not a parent array formula
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            var `val`: Any?
            try {
                `val` = this.`val`
                if (`val` == null) {
                    if (this.formulaRec!!.workBook!!.calcMode != WorkBookHandle.CALCULATE_EXPLICIT) {
                        this.calculate()
                        `val` = this.`val`
                    } else {
                        `val` = CalculationException(CalculationException.VALUE)
                    }
                } else if (`val` is String && `val`.startsWith("#"))
                    `val` = CalculationException(CalculationException.getErrorCode(`val` as String?))
            } catch (e: Exception) {
                `val` = CalculationException(CalculationException.VALUE)
            }

            if (`val` == null) {
                Logger.logErr("FormulaHandle.getOOXML:  unexpected null encountered when calculating formula: " + this.cellAddress)
            }
            if (`val` is String) {
                ooxml.append(" t=\"str\"")
                `val` = OOXMLAdapter.stripNonAscii(`val` as String?)
            } else if (`val` is Boolean) {
                ooxml.append(" t=\"b\"")
                if (`val`.booleanValue())
                    `val` = "1"
                else
                    `val` = "0"
            } else if (`val` is Double) {
                ooxml.append(" t=\"n\"")
            } else if (`val` is CalculationException) {
                ooxml.append(" t=\"e\"")
            }

            var fs: String? = "="
            try {
                fs = this.formulaString
            } catch (e: Exception) {
                Logger.logErr("FormulaHandle.getOOXML: error obtaining formula string: $e")
            }

            fs = OOXMLAdapter.stripNonAscii(fs).toString()
            if (!this.isArrayFormula) {
                ooxml.append("><f")
                fs = fs.substring(1)
            } else {
                if (formulaRec!!.sheet!!.isArrayFormulaParent(this.cellAddress)) {
                    ooxml.append("><f")
                    val refs = formulaRec!!.sheet!!.getArrayRef(this.cellAddress)
                    if (fs.startsWith("{=")) fs = fs.substring(2, fs.length - 1)
                    ooxml.append(" t=\"array\"")
                    ooxml.append(" ref=\"$refs\"")
                } else {
                    fs = null
                    ooxml.append(">")
                }
            }
            if (this.isSharedFormula) {
                try {
                } catch (e: Exception) {
                }

            }
            if (this.calcAlways) ooxml.append(" ca=\"1\"")
            if (fs != null)
                ooxml.append(">$fs</f>")
            ooxml.append("<v>$`val`</v>")
            return ooxml.toString()
        }

    /**
     * Sets the location lock on the Cell Reference at the
     * specified  location
     *
     *
     * Used to prevent updating of the Cell Reference when
     * Cells are moved.
     *
     * @param location of the Cell Reference to be locked/unlocked
     * @param lock     status setting
     * @return boolean whether the Cell Reference was found and modified
     */
    fun setLocationLocked(loc: String, l: Boolean): Boolean {
        var x = Ptg.PTG_LOCATION_POLICY_UNLOCKED
        if (l) x = Ptg.PTG_LOCATION_POLICY_LOCKED
        return formulaRec!!.setLocationPolicy(loc, x)
    }

    /**
     * Sets the location lock on the Cell Reference at the
     * specified  location
     *
     *
     * Used to prevent updating of the Cell Reference when
     * Cells are moved.
     *
     * @param location of the Cell Reference to be locked/unlocked
     * @param lock     status setting
     * @return boolean whether the Cell Reference was found and modified
     */
    fun setLocationPolicy(loc: String, l: Int): Boolean {
        return formulaRec!!.setLocationPolicy(loc, l)
    }

    init {
        formulaRec = f
    }

    /** Return the cached value of the Formula.
     *
     * This method returns the value as cached by OpenXLS or Excel of the formula.  Please note
     * that cases could exist where a cached value does not exist.  In this case getCachedVal will not try and calculate
     * the formula, it will return null.
     *
     * @return Object - cached value of the Formula as a String or a Double dependent on data type.
     *
     * public Object getCachedVal() {
     * return form.getCachedVal();
     * }
     */


    /**
     * Calculate the value of the formula and return it as an object
     *
     *
     * Calling calculate will ignore the WorkBook formula calculation flags
     * and forces calculation of the entire formula stack
     */
    @Throws(FunctionNotSupportedException::class)
    fun calculate(): Any? {
        formulaRec!!.clearCachedValue()

        return sanitizeValue(formulaRec!!.calculate())
    }


    /**
     * Sets the formula to a string passed in excel formula format.
     *
     * @param formulaString - String formatted as an excel formula, like Sum(A3+4)
     */

    @Throws(FunctionNotSupportedException::class)
    fun setFormula(formulaString: String) {
        formulaRec = FormulaParser.setFormula(formulaRec, formulaString, intArrayOf(formulaRec!!.rowNumber, formulaRec!!.colNumber.toInt()))
    }

    /**
     * If the Formula evaluates to a String, there
     * will be a Stringrec attached to the Formula
     * which contains the latest value.
     *
     * @return boolean whether this Formula evaluates to a String
     */
    fun evaluatesToString(): Boolean {
        return formulaRec!!.calculateFormula() is String
    }

    /**
     * Takes a string as a current formula location, and changes
     * that pointer in the formula to the new string that is sent.
     * This can take single cells"A5" and cell ranges,"A3:d4"
     * Returns true if the cell range specified in formulaLoc exists & can be changed
     * else false.  This also cannot change a cell pointer to a cell range or vice
     * versa.
     *
     * @param String - range of Cells within Formula to modify
     * @param String - new range of Cells within Formula
     */
    @Throws(FormulaNotFoundException::class)
    fun changeFormulaLocation(formulaLoc: String, newaddr: String): Boolean {
        val dx = formulaRec!!.getPtgsByLocation(formulaLoc)
        val lx = dx.iterator()
        while (lx.hasNext()) {
            try {
                val thisptg = lx.next() as Ptg
                ReferenceTracker.updateAddressPerPolicy(thisptg, newaddr)
                formulaRec!!.setCachedValue(null) // flag to recalculate
                return true
            } catch (e: Exception) {
                Logger.logInfo("updating Formula reference failed:$e")
                return false
            }

        }
        return true
    }

    /**
     * Changes a range in a formula to expand until it includes the
     * cell address from CellHandle.
     *
     *
     * Example:
     *
     *
     * CellHandle cell = new Cellhandle("D4")  Formula = SUM(A1:B2)
     * addCellToRange("A1:B2",cell); would change the formula to look like"SUM(A1:D4)"
     *
     *
     * Returns false if formula does not contain the formulaLoc range.
     *
     * @param String     - the Cell Range as a String to add the Cell to
     * @param CellHandle - the CellHandle to add to the range
     */
    @Throws(FormulaNotFoundException::class)
    fun addCellToRange(formulaLoc: String, handle: CellHandle): Boolean {
        val dx = formulaRec!!.getPtgsByLocation(formulaLoc)
        val lx = dx.iterator()
        var b = false
        while (lx.hasNext()) {
            val ptg = lx.next() as Ptg ?: return false
            val formulaaddr = ExcelTools.getRangeRowCol(formulaLoc)
            val handleaddr = handle.cellAddress
            val celladdr = ExcelTools.getRowColFromString(handleaddr)

            // check existing range and set new range vals if the new Cell is outside
            if (celladdr[0] > formulaaddr[2]) formulaaddr[2] = celladdr[0]
            if (celladdr[0] < formulaaddr[0]) formulaaddr[0] = celladdr[0]
            if (celladdr[1] > formulaaddr[3]) formulaaddr[3] = celladdr[1]
            if (celladdr[1] < formulaaddr[1]) formulaaddr[1] = celladdr[1]
            val newaddr = ExcelTools.formatRange(formulaaddr)
            b = this.changeFormulaLocation(formulaLoc, newaddr)

        }
        return b
    }


    override fun toString(): String {
        return this.formulaRec!!.cellAddress + ":" + this.formulaRec!!.formulaString
    }

    companion object {
        val supportedFunctions: Array<Array<String>>
            get() = FunctionConstants.recArr

        /**
         * Converts a cell value to a form suitable for the public API.
         * Currently this converts cached errors ([CalculationException]s)
         * to the corresponding error string.
         */
        internal fun sanitizeValue(`val`: Any?): Any? {
            return if (`val` is CalculationException) `val`.name else `val`
        }

        /**
         * Copy the formula references with offsets
         *
         * @param int[row,col] offsets to move the references
         * @return
         */
        @Throws(FormulaNotFoundException::class)
        fun moveCellRefs(fmh: FormulaHandle, offsets: IntArray) {
            // get the current offsets from the FMH references
            val celladdys = fmh.ranges

            // iterate
            for (x in celladdys.indices) {
                val s = ExcelTools.stripSheetNameFromRange(celladdys[x])
                val sh = s[0]        // sheet portion of address,if any
                var range = s[1]    // range or single address
                val rangeIdx = range.indexOf(":")
                var secondAddress: String? = null
                if (rangeIdx > -1) {    // separate out addresses within a range
                    secondAddress = range.substring(rangeIdx + 1)
                    range = range.substring(0, rangeIdx)
                }
                var orig = ExcelTools.getRowColFromString(range)
                var relCol = !range.startsWith("$")    // 20100603 KSC: handle relative refs
                var relRow = !(range.length > 0 && range.substring(1).indexOf('$') > -1)
                if (relRow)
                // only move if relative ref
                    orig[0] += offsets[0] //row
                if (relCol)
                // only move if relative ref
                    orig[1] += offsets[1] //col
                var newAddress = ExcelTools.formatLocation(orig, relRow, relCol)
                if (orig[0] < 0 || orig[1] < 0) newAddress = "#REF!"
                if (secondAddress != null) {
                    orig = ExcelTools.getRowColFromString(secondAddress)
                    relCol = !secondAddress.startsWith("$")    // handle relative refs
                    relRow = !(secondAddress.length > 0 && secondAddress.substring(1).indexOf('$') > -1)
                    if (orig[0] >= 0) {
                        if (relRow)
                        // only move if relative ref
                            orig[0] += offsets[0] //row
                    }
                    if (orig[1] >= 0) { //if not wholerow/wholecol ref
                        if (relCol)
                        // only move if relative ref
                            orig[1] += offsets[1] //col
                    }
                    val newAddress1 = ExcelTools.formatLocation(orig, relRow, relCol)
                    newAddress = "$newAddress:$newAddress1"
                }
                if (sh != null)
                // TODO: handle refs with multiple sheets
                    newAddress = "$sh!$newAddress"
                if (!fmh.changeFormulaLocation(celladdys[x], newAddress)) {
                    Logger.logErr("Could not change Formula Reference: " + celladdys[x] + " to: " + newAddress)
                }
            }
            return
        }
    }
}