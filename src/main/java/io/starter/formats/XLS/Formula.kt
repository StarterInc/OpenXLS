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
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.*

/**
 * FORMULA (0x406) describes a cell that contains a formula.
 * <pre>
 * offset  name        size    contents
 * 0       rw          2       Row
 * 2       col         2       Column
 * 4       ixfe        2       Index to the XF record
 * 6       num         8       Current value of the formula
 * 14      grbit       2       Option Flags
 * 16      chn         4       (Reserved, must be zero) A field that specifies an application-specific cache of information. This cache exists for performance reasons only, and can be rebuilt based on information stored elsewhere in the file without affecting calculation results.
 * 20      cce         2       Parsed Expression length
 * 22      rgce        cce     Parsed Expression
</pre> *
 * The grbit field contains the following flags:
 * <pre>
 * byte   bits   mask   name           asserted if
 * 0      0      0x01   fAlwaysCalc    the result must not be cached
 * 1      0x02   fCalcOnLoad    the cached value is incorrect
 * 2      0x04   (Reserved)
 * 3      0x08   fShrFmla       this is a reference to a shared formula
 * 7-4    0xF0   (Unused)
 * 1      7-0    0xFF   (Unused)
</pre> *
 * In most cases, formulas should have fAlwaysCalc asserted to ensure that
 * correct values are displayed upon opening of the file in Excel.
 */
class Formula : XLSCellRecord() {

    private var cachedValue: Any? = null
    private var expression: Stack<*>? = null


    /**
     * Whether the record data needs to be updated.
     */
    private var dirty = false

    /**
     * Whether this formula contains an indirect reference.
     */
    private var containsIndirectFunction = false

    /**
     * Whether this FORMULA record has an attached STRING record.
     */
    private var haveStringRec = false

    /**
     * Contains bitfield flags.
     */
    private var grbit = FCALCONLOAD

    /**
     * The attached STRING record, if one exists.
     */
    /**
     * Get the String that is attached to this formula (it has a result of a string..);
     */
    var attatchedString: StringRec? = null
        private set

    /**
     * The target ShrFmla record, if this is a shared formula reference.
     */
    var shared: Shrfmla? = null

    /**
     * List of records attached to this one.
     */
    private var internalRecords: MutableList<*>? = null

    /**
     * true if it's sub-ptgs are defined in other workbooks and therefore unable to be resolved
     */
    private var isExternalRef = false

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
        get() = grbit and FALWAYSCALC != 0
        set(fAlwaysCalc) = if (fAlwaysCalc)
            grbit = grbit or FALWAYSCALC
        else
            grbit = grbit and FALWAYSCALC.inv().toShort()

    val typeName: String
        get() = "formula"


    /**
     * Returns the Human-Readable String Representation of
     * this Formula.
     */
    val formulaString: String
        get() {
            populateExpression()
            return if (!this.isArrayFormula) {
                FormulaParser.getFormulaString(this)
            } else "{" + FormulaParser.getFormulaString(this) + "}"

        }

    /**
     * Returns an array of ptgs that represent any BiffRec ranges in the formula.
     * Ranges can either be in the format "C5" or "Sheet1!C4:D9"
     */
    val cellRangePtgs: Array<Ptg>
        @Throws(FormulaNotFoundException::class)
        get() = ExpressionParser.getCellRangePtgs(expression!!)

    /**
     * Get the value of the formula as an integer.
     * If the formula exceeds integer boundaries, or is a float with
     * a non-zero mantissa throw an exception
     *
     * @see io.starter.formats.XLS.XLSRecord.getIntVal
     */
    override// not back-compat return Integer.valueOf(new Long((long) tl).intValue()).intValue();
    // throw new NumberFormatException("Loss of precision converting " + tl + " to int.");
    // return a zero for empties
    // 20090514 KSC: bd.intValueExact(); is 1.6 compatible -- MAY CAUSE INFOTERIA REGRESSION ERROR
    var intVal: Int
        @Throws(RuntimeException::class)
        get() {
            val obx = calculateFormula()
            try {
                val tl = (obx as Double).toDouble()
                if (tl > Integer.MAX_VALUE) {
                    throw NumberFormatException("getIntVal: Formula value is larger than the maximum java signed int size")
                }
                if (tl < Integer.MIN_VALUE) {
                    throw NumberFormatException("getIntVal: Formula value is smaller than the minimum java signed int size")
                }
                val db = obx.toDouble()
                val ret = obx.toInt()
                if (db - ret > 0 && DEBUGLEVEL > this.DEBUG_LOW)
                    Logger.logWarn("Loss of precision converting $tl to int.")

                return ret
            } catch (e: ClassCastException) {
            }

            var l: Long = 0
            var s = obx.toString()
            if (s == "") s = "0"
            try {
                val t = ""
                val bd = java.math.BigDecimal(s)
                l = bd.toLong()
                if (l > Integer.MAX_VALUE) {
                    throw NumberFormatException("Formula value is larger than the maximum java signed int size")
                }
                if (l < Integer.MIN_VALUE) {
                    throw NumberFormatException("Formula value is smaller than the minimum java signed int size")
                }
                return Integer.valueOf(l.toString()).toInt()
            } catch (ne: NumberFormatException) {
                throw NumberFormatException("getIntVal: Formula is a non-numeric value")
            } catch (e: Exception) {
                throw NumberFormatException("getIntVal: $e")
            }

        }
        set


    override// return a zero for empties
    val floatVal: Float
        get() {
            val obx = calculateFormula()
            try {
                if (obx is Float) {
                    return obx.toFloat()
                }
            } catch (e: Exception) {
                Logger.logErr("Formula.getFloatVal failed for: $this", e)
            }

            try {
                var s = obx.toString()
                if (s == "") s = "0"
                return Float(s)
            } catch (ex: NumberFormatException) {
                return java.lang.Float.NaN
            } catch (e: Exception) {
                Logger.logWarn("Formula.getFloatVal() failed: $e")
            }

            return java.lang.Float.NaN
        }

    override var booleanVal: Boolean
        get() {
            val obx = calculateFormula()
            try {
                if (obx is Boolean) {
                    return obx.booleanValue()
                }
            } catch (e: Exception) {
                Logger.logErr("getBooleanVal failed for: $this", e)
            }

            try {
                val s = obx.toString()
                if (s.equals("true", ignoreCase = true) || s == "1") return true

            } catch (e: Exception) {
                Logger.logWarn("getBooleanVal() failed: $e")
            }

            return false
        }
        set

    override// return a zero for empties
    val dblVal: Double
        get() {
            val obx = calculateFormula()
            try {
                if (obx is Double) {
                    return obx.toDouble()
                }
            } catch (e: Exception) {
                Logger.logErr("Formula.getDblVal failed for: $this", e)
            }

            var s = obx.toString()
            if (s == "") s = "0"
            try {
                return Double(s)
            } catch (ex: NumberFormatException) {
                return java.lang.Double.NaN
            } catch (e: Exception) {
                Logger.logWarn("Formula.getDblVal() failed: $e")
            }

            return java.lang.Double.NaN
        }

    /**
     * return the String representation of the current Formula value
     */
    override// if null, return empty string
    // TODO: set the string value of the attached string?
    var stringVal: String?
        get() {
            val obx = calculateFormula()
            try {
                if (obx is Double) {
                    val d = obx.toDouble()
                    return if (!java.lang.Double.isNaN(d))
                        Formula.getDoubleAsFormattedString(d)
                    else
                        "NaN"
                }
            } catch (e: Exception) {
                Logger.logErr("Formula.getStringVal failed for: $this", e)
            }

            return obx?.toString() ?: ""

        }
        set(v) = throw CellTypeMismatchException("Attempting to set a string value on a formula")

    /**
     * Returns whether this is a reference to a shared formula.
     */
    /**
     * Sets whether this is a reference to a shared formula.
     */
    var isSharedFormula: Boolean
        get() = grbit and FSHRFMLA != 0
        private set(isSharedFormula) = if (isSharedFormula)
            grbit = grbit or FSHRFMLA
        else
            grbit = grbit and FSHRFMLA.inv().toShort()

    /**
     * return truth of "this is an array formula" i.e. contains an Array sub-record
     *
     * @return
     */
    val isArrayFormula: Boolean
        get() = if (internalRecords != null && internalRecords!!.size > 0) internalRecords!![0] is Array else false

    /**
     * fetches the internal Array record linked to this formula, if any,
     * or null if not an array formula
     *
     * @return array record
     * @see isArrayFormula
     */
    /*		if (expression.get(0) instanceof PtgExp) {
			// if it's the child of a parent array formula, obtain it's Array record
			// TODO: verify this is correct + finish
		}*/ val array: Array?
        get() {
            try {
                return internalRecords!![0] as Array
            } catch (e: Exception) {
            }

            return null
        }

    /**
     * OOXML-specific: set the range the Array references
     *
     * @param s
     */
    var arrayRefs: String
        get() {
            if (internalRecords != null && internalRecords!!.size > 0) {
                val o = internalRecords!![0]
                return (o as Array).arrayRefs
            }
            return ""
        }
        set(s) {
            if (internalRecords != null && internalRecords!!.size > 0) {
                val o = internalRecords!![0]
                if (o is Array) {
                    val rc = ExcelTools.getRangeRowCol(s)
                    o.firstRow = rc[0]
                    o.firstCol = rc[1]
                    o.lastRow = rc[2]
                    o.lastCol = rc[3]
                }
            }
        }

    /**
     * clear out object references in prep for closing workbook
     */
    private var closed = false

    /**
     * Default constructor
     */
    init {
        opcode = XLSConstants.FORMULA
        isValueForCell = true
        isFormula = true
    }

    /**
     * Parses the record bytes.
     * This method only needs to be called when the record is being constructed
     * from bytes. Calling it on a programmatically created formula is
     * unnecessary and will probably throw an exception.
     */
    override fun init() {
        // Prevent misuse of init
        if (expression != null)
            throw IllegalStateException(
                    "can't init a formula created from a string")

        super.init()
        if ((data = getData()) == null)
            throw IllegalStateException(
                    "can't init a formula without record bytes")
        super.initRowCol()
        ixfe = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()
        grbit = ByteTools.readShort(this.getByteAt(14).toInt(), this.getByteAt(15).toInt())

        // get the cached value bytes from the record
        val currVal = this.getBytesAt(6, 8)

        // Is this a non-numeric value?
        if (currVal!![6] == 0xFF.toByte() && currVal[7] == 0xFF.toByte()) {
            // String value
            if (currVal[0] == 0x00.toByte()) {
                haveStringRec = true        // bytes 1-5 are not used
                // Normally cachedValue will be set by StringRec's init.
                // Setting cachedValue null forces calculation in the rare
                // event that the STRING record is missing or fails to init.
                cachedValue = null
            } else if (currVal[0] == 0x03.toByte()) {
                // There is no attached STRING record. Set cachedValue directly.
                cachedValue = ""
            } else if (currVal[0] == 0x01.toByte()) {
                cachedValue = java.lang.Boolean.valueOf(currVal[2] != 0x00.toByte())
            } else if (currVal[0] == 0x02.toByte()) {
                cachedValue = CalculationException(currVal[2])
            } else {
                cachedValue = null
            }// Unknown value type
            // Error value
            // Boolean value
            // Empty string value

        } else {
            // do not cache NaN stored bytes
            val dbv = ByteTools.eightBytetoLEDouble(currVal)
            if (!java.lang.Double.isNaN(dbv)) {
                cachedValue = dbv
            }
        }

        if (this.sheet == null)
            this.setSheet(this.wkbook!!.lastbound)

        if (DEBUGLEVEL > 10) {
            try {
                Logger.logInfo("INFO: Formula " + this.cellAddress + this.formulaString)
            } catch (e: Exception) {
                Logger.logInfo("Debug output of Formula failed: $e")
            }

        }

        // The expression needs to be parsed on input in order to add it to
        // the reference tracker
        //TODO: Add a no calculation / read only mode without ref tracking
        this.populateExpression()
        // Perform some special handling for formulas with indirect references
        if (containsIndirectFunction) this.registerIndirectFunction()
        this.dirty = false
    }

    /**
     * Performs cleanup required before changing or removing this Formula.
     * This nulls out the expression, so save a copy if you need it.
     * Possible sub-records associated with Formula: array, shared string and/or shared formula
     */
    private fun clearExpression() {
        if (expression == null)
            return

        if (this.isArrayFormula) {
            val a = this.array
            if (a != null)
                this.sheet!!.removeRecFromVec(a)
        }
        if (this.hasAttachedString()) {    // remove that too
            if (attatchedString != null)
                this.sheet!!.removeRecFromVec(attatchedString!!)
            attatchedString = null
        }

        if (isSharedFormula) {
            shared!!.removeMember(this)
            shared = null
            isSharedFormula = false
        }

        val iter = expression!!.iterator()
        while (iter.hasNext()) {
            val ptg = iter.next() as Ptg
            if (ptg is PtgRef) ptg.removeFromRefTracker()
        }

        expression = null
    }

    /**
     * for creating a formula on the fly.  This is used currently by FormulaParser
     * to create a formula.  The location and XF fields are filled in by boundsheet.add();
     */
    fun setExpression(exp: Stack<*>) {
        if (expression != null)
            clearExpression()
        expression = exp
        updateRecord()
    }

    /**
     * get the expression stack
     *
     * @return
     */
    fun getExpression(): Stack<*>? {
        populateExpression()
        return expression
    }

    /**
     * store formula-associated records: one or more of:
     * ShrFmla
     * StringRec
     * Array
     *
     * @param b
     */
    fun addInternalRecord(b: BiffRec) {
        if (internalRecords == null) internalRecords = ArrayList(3)

        if (b is Shrfmla)
            internalRecords!!.add(0, b)
        else if (b is StringRec) {
            if (!haveStringRec || attatchedString == null) {// should ONLY have 1 StringRec
                internalRecords!!.add(b)
                haveStringRec = true
                attatchedString = b
                cachedValue = attatchedString!!.stringVal
            }/* TODO: this is such a rare occurrence - due possibly to OUR processing - keep if and wa
        		else
        		Logger.logErr("Formula.init:  Out of Spec Formula Encountered (Multiple String Recs Encountered)- Ignoring");
        		*/
        } else
        // array formula
            internalRecords!!.add(b)

    }

    fun removeInternalRecord(b: BiffRec) {
        if (internalRecords == null) return
        internalRecords!!.remove(b)
    }

    fun getInternalRecords(): List<*> {
        return if (internalRecords == null) emptyList<Any>() else internalRecords
    }


    fun hasAttachedString(): Boolean {
        return haveStringRec
    }

    /**
     * set if this formula refers to External References
     * (references defined in other workbooks)
     *
     * @param isExternalRef
     */
    fun setIsExternalRef(isExternalRef: Boolean) {
        this.isExternalRef = isExternalRef
    }

    /**
     * Adds an indirect function to the list of functions to be evaluated post load
     */
    fun registerIndirectFunction() {
        this.workBook!!.addIndirectFormula(this)
    }

    /**
     * If the method contains an indirect function then
     * register those ptgs into the reference tracker.
     *
     *
     * In order to do this it is necessary to calculate the formula
     * to retrieve the Ptg's
     */
    fun calculateIndirectFunction() {
        this.clearCachedValue()
        try {
            this.calculateFormula()
        } catch (e: FunctionNotSupportedException) {
            // If we do not support the function, calculation will throw a FNE anyway, so no need for logging
        } catch (e: Exception) {
            // problematic here.  As the calculation is happening on parse we dont really
            // want to throw an exception and crap out on the book loading
            // but the client should be informed in some way.  Also a generic exception is caught
            // because our code does not bubble a calc exception up.
            Logger.logErr("Error registering lookup for INDIRECT() function at cell: " + this.cellAddress + " : " + e)
        }

    }

    /**
     * Populates the expression in the formula.  This has been moved out of init for performance reasons.
     * The idea is that the processing is offloaded as a JIT for calculation/value retrieval.
     */
    //TODO: refactor external references and make private
    internal fun populateExpression() {
        if (expression != null || data == null) return

        try {
            val length = ByteTools.readShort(
                    this.getByteAt(20).toInt(), this.getByteAt(21).toInt())

            if (length + 22 > data!!.size)
                throw Exception(
                        "cce longer than record")

            expression = ExpressionParser.parseExpression(
                    this.getBytesAt(22, reclen - 22)!!, this, length.toInt())

            // If this is a shared formula reference, do some special init
            if (isSharedFormula)
                initSharedFormula(null)
        } catch (e: Exception) {
            if (DEBUGLEVEL > -10)
                Logger.logInfo("Formula.init:  Parsing Formula failed: $e")
        }

    }

    /**
     * Performs special initialization for shared formula references.
     *
     * @param target the target `SHRFMLA` record. If this is
     * `null` it will be retrieved from the cell pointed to
     * by the `PtgExp`.
     * @throws FormulaNotFoundException if the target shared formula is missing
     * @throws IllegalArgumentException if this is not a shared formula member
     */
    @Throws(FormulaNotFoundException::class)
    internal fun initSharedFormula(target: Shrfmla?) {
        if (!isSharedFormula) {
            isSharedFormula = true
        }

        // If we're already done, silently do nothing
        if (shared != null)
            return

        // If this is an instantiation instead of a reference
        if (expression!!.size != 1 || expression!![0] !is PtgExp) {
            //TODO: find which ShrFmla this is and convert to reference
            // For now, just clear fShrFmla
            isSharedFormula = false
            return
        }

        val pointer = expression!![0] as PtgExp

        if (target != null) {
            shared = target
        } else
            try {
                shared = (sheet!!.getCell(
                        pointer.rwFirst, pointer.colFirst) as Formula).shared    // find shared cell linked to host/first formula cell in shared formula series
                if (shared == null)
                    throw Exception()
            } catch (e: Exception) {
                // If this is the host cell, fail silently. This method will be
                // re-called by the ShrFmla record's init method.
                if (this.cellAddress == pointer.referent)
                    return

                // Otherwise, complain and clear fShrFmla
                throw FormulaNotFoundException(
                        "FORMULA at " + this.cellAddress
                                + " refers to missing SHRFMLA at "
                                + pointer.referent)
            }

        //Shared Formula Init Performance Changes:  do not instantiate until calculate
        //		expression = shared.instantiate( pointer );
        shared!!.addMember(this)

        if (shared!!.containsIndirectFunction) registerIndirectFunction()
    }

    /**
     * Converts a shared formula reference into a normal formula.
     *
     * @throws IllegalStateException if this is not a shared formula member
     */
    fun convertSharedFormula() {
        if (!isSharedFormula)
            throw IllegalStateException(
                    "not a shared formula reference")

        shared = null
        isSharedFormula = false
    }


    /**
     * Returns the ptg that matches the string location sent to it.
     * this can either be in the format "C5" or a range, such as "C4:D9"
     */
    @Throws(FormulaNotFoundException::class)
    fun getPtgsByLocation(loc: String): List<*> {
        var loc = loc
        populateExpression()
        if (loc.indexOf("!") == -1)
            loc = this.sheet!!.sheetName + "!" + loc

        return ExpressionParser.getPtgsByLocation(loc, expression!!)
    }

    /**
     * locks the Ptg at the specified location
     */

    fun setLocationPolicy(loc: String, l: Int): Boolean {
        populateExpression()
        try {
            val dx = this.getPtgsByLocation(loc)
            val lx = dx.iterator()
            while (lx.hasNext()) {
                val d = lx.next() as Ptg
                d.locationPolicy = l
                if (l == Ptg.PTG_LOCATION_POLICY_TRACK)
                // init the tracker cell right away
                    d.initTrackerCell()
            }
            return true
        } catch (e: FormulaNotFoundException) {
            Logger.logInfo("locking Formula Location failed:$loc: $e")
            return false
        }

    }

    /**
     * Called to indicate that the parsed expression has changed.
     * Does nothing if the expression doesn't exist yet.
     */
    fun updateRecord() {
        dirty = true
        if (this.data == null)
            setData(ByteArray(6))    // happens when newly init'ing a formula

        if (cachedValue is String && "" != cachedValue && !isErrorValue(cachedValue as String?)) {// if it's a string and not an error string
            if (!haveStringRec || attatchedString == null) {
                //					this.addInternalRecord( new StringRec( (String)cachedValue ) ); // will be added in workbook.addRecord
                attatchedString = StringRec(cachedValue as String?)
                attatchedString!!.setSheet(sheet)
                attatchedString!!.rowNumber = rowNumber
                attatchedString!!.setCol(colNumber)
                workBook!!.setLastFormula(this)    // for addRecord, sets appropriate formula internal record
                workBook!!.addRecord(attatchedString!!, true)
                haveStringRec = true
            } else {
                attatchedString!!.stringVal = cachedValue as String?
            }
        } else if (attatchedString != null) {
            attatchedString!!.remove(false)
            removeInternalRecord(attatchedString)
            haveStringRec = false
        }
    }

    /**
     * Updates the record data if necessary to prepare for streaming.
     */
    override fun preStream() {
        // If the record doesn't need to be updated, do nothing
        if (!dirty && !isSharedFormula && cachedValue != null &&
                workBook!!.calcMode != WorkBook.CALCULATE_EXPLICIT)
            return


        // If the formula needs calculation, do so
        try {
            if (cachedValue == null) calculateFormula()
        } catch (e: FunctionNotSupportedException) {
            // Fall through to null cachedValue handling below
        }

        // Sometimes we need to write a value other than the real cached value
        var writeValue = cachedValue

        // Handle formulas that can't be calculated for whatever reason
        if (cachedValue == null) {
            grbit = grbit or FCALCONLOAD
            if (writeValue == null) writeValue = java.lang.Double.NaN
        }

        // Handle CALCULATE_EXPLICIT mode
        if (workBook!!.calcMode == WorkBook.CALCULATE_EXPLICIT) {
            grbit = grbit or FCALCONLOAD
        }

        // If this is a shared formula, write a PtgExp
        var expr: Stack<*>? = expression
        if (isSharedFormula) {    /* ONLY need to do this if shared formula member(s) have changed-
    									a better choice is to trap and change in respective method */
            expr = Stack()
            expr!!.add(shared!!.pointer)
        }

        // Fetch the Ptg data and calculate the expression size
        val ptgdata = arrayOfNulls<ByteArray>(expr!!.size)
        var rgb: ByteArray? = null
        var cce: Short = 0
        var rgblen: Short = 0
        for (idx in expr.indices) {
            val ptg = expr[idx] as Ptg
            ptgdata[idx] = ptg.record
            cce += ptgdata[idx].size.toShort()

            if (ptg is PtgArray) {
                val extra = ptg.postRecord
                rgb = ByteTools.append(extra, rgb)
                rgblen += extra!!.size.toShort()
            } else if (ptg is PtgMemArea) {    // has PtgExtraMem structure appended
                val extra = ptg.postRecord
                rgb = ByteTools.append(extra, rgb)
                rgblen += extra.size.toShort()
            }
        }

        val newdata = ByteArray(22 + cce.toInt() + rgblen.toInt())
        // Cell Header (row, col, ixfe)
        System.arraycopy(data!!, 0, newdata, 0, 6)

        // Cached Value (num)
        val value: ByteArray
        if (writeValue is Number) {
            //TODO: Check infinity, NaN, etc.
            value = ByteTools.toBEByteArray(
                    writeValue.toDouble())
        } else {
            value = ByteArray(8)
            // byte 0 specifies the marker type
            value[1] = 0x00.toByte()
            // byte 2 is used by bool and boolerr
            value[3] = 0x00.toByte()
            value[4] = 0x00.toByte()
            value[5] = 0x00.toByte()
            value[6] = 0xFF.toByte()
            value[7] = 0xFF.toByte()

            if (writeValue is String) {
                if (!isErrorValue(writeValue as String?)) {
                    value[2] = 0x00.toByte()
                    val sval = writeValue as String?
                    if (sval == "" || attatchedString == null) {    // the latter can occur when input from XLSX; a cachedvalue is set without an associated StringRec
                        value[0] = 0x03.toByte() // means empty
                    } else {
                        value[0] = 0x00.toByte()
                        attatchedString!!.stringVal = sval
                    }
                } else {
                    value[0] = 0x02.toByte()    // error code
                    value[2] = CalculationException.getErrorCode(writeValue as String?)
                }
            } else if (writeValue is Boolean) {
                value[0] = 0x01.toByte()
                value[2] = if (writeValue.booleanValue())
                    0x01.toByte()
                else
                    0x00.toByte()
            } else if (writeValue is CalculationException) {
                value[0] = 0x02.toByte()
                value[2] = writeValue.errorCode
            } else
                throw Error("unknown value type " + (writeValue?.javaClass?.getName() ?: "null"))
        }
        System.arraycopy(value, 0, newdata, 6, 8)

        // Bit Flags (grbit)
        System.arraycopy(ByteTools.shortToLEBytes(grbit), 0, newdata, 14, 2)

        // chn - reserved zero
        Arrays.fill(newdata, 16, 19, 0x00.toByte())

        // Expression Length (cce)
        System.arraycopy(ByteTools.shortToLEBytes(cce), 0, newdata, 20, 2)

        // Expression Ptgs (rgce)
        var offset = 22
        for (idx in ptgdata.indices) {
            System.arraycopy(
                    ptgdata[idx], 0, newdata, offset, ptgdata[idx].size)
            offset += ptgdata[idx].size
        }

        // Expression Extra Data (rgb)
        if (rgblen > 0) System.arraycopy(rgb!!, 0, newdata, offset, rgblen.toInt())
        setData(newdata)
        dirty = false
    }

    override fun clone(): Any? {
        // Make the record bytes available to XLSRecord.clone
        preStream()
        return super.clone()
    }

    /**
     * Calculates the formula honoring calculation mode.
     *
     * @throws CalculationException
     */
    @Throws(FunctionNotSupportedException::class)
    fun calculateFormula(): Any? {
        // if this is calc explicit, we ALWAYS use cache
        if (workBook!!.calcMode == WorkBook.CALCULATE_EXPLICIT)
            return cachedValue
        // TODO: IF ALREADY RECALCED DONT SET TO null -- need flag?
        if (workBook!!.calcMode == WorkBook.CALCULATE_ALWAYS) {
            if (!isExternalRef)
            // if it's an external reference DONT CLEAR CACHE
                cachedValue = null // force calc
            else
                return cachedValue
        }
        return calculate()

    }

    /**
     * Calculate the formula if necessary.  This accessor resets the recurse count on the
     * formula.  Excel standard is to allow 100 recursions before throwing a circular reference.
     *
     * @throws CalculationException
     */
    fun calculate(): Any? {
        val depth = recurseCount.get()

        try {
            recurseCount.set(depth!! + 1)
            if (depth > WorkBookHandle.RECURSION_LEVELS_ALLOWED) {
                Logger.logWarn("Recursion levels reached in calculating formula "
                        + this.cellAddressWithSheet
                        + ". Possible circular reference.  Recursion levels can be set through WorkBookHandle.setFormulaRecursionLevels")
                cachedValue = CalculationException(
                        CalculationException.CIR_ERR)
                return cachedValue
            }
            return this.calculateInternal()
        } finally {
            recurseCount.set(depth)
        }
    }


    /**
     * Calculates the formula if necessary regardless of calculation mode.
     * If there is a cached value it will be returned. Otherwise, the formula
     * will be calculated and the result will be cached and returned. If you
     * need to force calculation call [.clearCachedValue] first.
     */
    private fun calculateInternal(): Any? {
        // If we have a cached value, return it instead of calculating
        if (cachedValue != null) return cachedValue
        populateExpression()
        try {
            cachedValue = FormulaCalculator.calculateFormula(this.expression!!)
        } catch (e: StackOverflowError) {
            Logger.logWarn("Stack overflow while calculating "
                    + this.cellAddressWithSheet
                    + ". Possible circular reference.")
            cachedValue = CalculationException(
                    CalculationException.CIR_ERR)
            return cachedValue
        }

        if (cachedValue == null)
            throw FunctionNotSupportedException("Unable to calculate Formula " + this.formulaString + " at: " + this.sheet!!.sheetName + "!" + this.cellAddress)

        if (cachedValue!!.toString() == "#CIR_ERR!") {
            return CircularReferenceException(CalculationException.CIR_ERR)
        }
        if (cachedValue!!.toString().length < 1) {
            // do something...?
        } else if (cachedValue!!.toString()[0] == '{') {
            // it's an array, we need to find the particular value that we want.
            // parse all array strings into rows, cols
            var arrStr: String = cachedValue as String?
            arrStr = arrStr.substring(1, arrStr.length - 1)
            var rows: Array<String>? = null
            var cols: Array<Array<String>>? = null
            // split rows
            rows = arrStr.split(";".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            cols = arrayOfNulls<Array<String>>(rows!!.size)
            for (i in rows!!.indices) {
                cols[i] = rows!![i].split(",".toRegex()).toTypedArray()    // include empty strings
            }
            val pxp: PtgExp
            var rowA: Int
            var colA: Int
            val p: Ptg
            try {
                pxp = expression!!.elementAt(0) as PtgExp
                rowA = this.rowNumber - pxp.rwFirst
                colA = this.colNumber - pxp.colFirst
                // now, if it's a 1-dimensional array e.g {1,2,3,4,5}, nr=1, nc= 5
                // if formula address is traversing rows then switch
                if (rows!!.size == 1 && rowA > 0 && colA == 0) {
                    colA = rowA
                    rowA = 0
                }
            } catch (e: ClassCastException) {
                // this is when we just calc'd a formula and have no exp reference.
                // assume it is the location of the formula
                // could be incorrect, may need to revisit
                rowA = 0
                colA = 0
            }

            cachedValue = cols!![rowA][colA]
            // try to cast
            try {
                cachedValue = Double((cachedValue as String?)!!)
            } catch (e: Exception) {
                // let it go
            }

            if (DEBUGLEVEL > 10) Logger.log(cachedValue!!)
        }
        if (this.attatchedString != null) {
            this.attatchedString!!.stringVal = cachedValue.toString()
        }

        updateRecord()
        return cachedValue
    }

    override fun toString(): String {
        populateExpression()
        return super.toString()
        //return this.worksheet.getSheetName() + "!" + this.getCellAddress() + ":" + this.getStringVal();
    }

    /**
     * Set the cached value of this formula,
     * in cases where the formula is null, set the cache to null,
     * as well as updating the attached string to null in order to force
     * recalc
     *
     * @see io.starter.formats.XLS.XLSRecord.setCachedValue
     */
    fun setCachedValue(newValue: Any?) {
        if (newValue == null) {
            this.clearCachedValue()
        } else {
            cachedValue = newValue    // TODO: need to check/validate StringRec ????
        }
    }


    /**
     * Set the cached value of this formula,
     * in cases where the formula is null, set the cache to null,
     * as welll as updating the attached string to null in order to force
     * recalc
     *
     * @see io.starter.formats.XLS.XLSRecord.setCachedValue
     */
    fun clearCachedValue() {
        cachedValue = null
        haveStringRec = false
        //         this.updateRecord(); no need; will be updated after recalc, which will automatically happen on write
    }

    /**
     * Set if the formula contains Indirect()
     *
     * @param containsIndirectFunction The containsIndirectFunction to set.
     */
    fun setContainsIndirectFunction(containsIndirectFunction: Boolean) {
        this.containsIndirectFunction = containsIndirectFunction
    }

    /**
     * Performs cleanup needed before removing the formula cell from the
     * work sheet. The formula will not behave correctly once this is called.
     */
    fun destroy() {
        clearExpression()
    }

    override fun close() {
        if (expression != null) {
            while (!expression!!.isEmpty()) {
                var p: GenericPtg? = expression!!.pop() as GenericPtg
                if (p is PtgRef)
                    p.close()
                else
                    p!!.close()/*	        	else if (p instanceof PtgExp ) {
	        		Ptg[] ptgs= ((PtgExp)p).getConvertedExpression();
	        		for (int i= 0; i < ptgs.length; i++) {
	    	        	if (ptgs[i] instanceof PtgRef)
	    	        		((PtgRef) ptgs[i]).close();
	    	        	else
	    	        		((GenericPtg)ptgs[i]).close();
	        		}
	        	} */
                p = null
            }
        }
        if (attatchedString != null) {
            attatchedString!!.close()
            attatchedString = null
        }
        if (shared != null) {
            if (shared!!.members == null || shared!!.members!!.size == 1)
            // last one
                shared!!.close()
            else
                shared!!.removeMember(this)
            shared = null
        }
        if (internalRecords != null)
            internalRecords!!.clear()
        super.close()
        closed = true
    }

    protected fun finalize() {
        if (!closed) {
            this.close()
        }
    }

    /**
     * Replaces a ptg in the active expression.  Useful for replacing a ptgRef with a ptgError after a bad movement.
     *
     * @param thisptg
     * @param ptgErr
     */
    fun replacePtg(thisptg: Ptg, ptgErr: Ptg) {
        ptgErr.parentRec = this
        val idx = this.expression!!.indexOf(thisptg)
        this.expression!!.removeAt(idx)
        this.expression!!.insertElementAt(ptgErr, idx)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 7563301825566021680L
        /**
         * Mask for the fAlwaysCalc grbit flag.
         */
        private val FALWAYSCALC: Short = 0x01
        /**
         * Mask for the fCalcOnLoad grbit flag.
         */
        private val FCALCONLOAD: Short = 0x02
        /**
         * Mask for the fShrFmla grbit flag.
         */
        private val FSHRFMLA: Short = 0x08


        /**
         * Returns the correct string representation of a double for excel.
         *
         *
         * Note this is for the standards that were determined with excel and io.starter.OpenXLS
         *
         * @param num
         * @return
         */
        private fun getDoubleAsFormattedString(theNum: Double): String {
            return ExcelTools.getNumberAsString(theNum)
        }

        private val recurseCount = object : ThreadLocal<Int>() {
            override fun initialValue(): Int {
                return 0
            }
        }

        /**
         * increment each PtgRef in expression stack via row or column based
         * on rowInc or colInc values
         * Used in OOXML parsing
         */
        fun incrementSharedFormula(origStack: java.util.Stack<*>, rowInc: Int, colInc: Int, range: IntArray) {
            // traverse thru ptg's, incrementing row reference
            //java.util.Stack origStack= form.getExpression();	Don't do "in place" as alters original expression
            //Logger.logInfo("Before Inc: " + this.getFormulaString());
            for (i in origStack.indices) {
                val p = origStack.elementAt(i) as Ptg
                try {
                    val s = p.location
                    if (p.isReference) {
                        if (!((p as PtgRef).isWholeRow && colInc != 0 || p.isWholeCol && rowInc != 0)) {
                            if (p !is PtgArea) {
                                val bRelRefs = booleanArrayOf(p.isRowRel, p.isColRel)
                                val rc = ExcelTools.getRowColFromString(s)
                                if (bRelRefs[0])
                                    rc[0] += rowInc
                                if (bRelRefs[1])
                                    rc[1] += colInc
                                val pr = PtgRef()
                                pr.parentRec = p.parentRec
                                pr.useReferenceTracker = false    // 20090827 KSC: don't petform expensive calcs on init + removereferenceTracker blows out cached value
                                pr.location = ExcelTools.formatLocation(rc, bRelRefs[0], bRelRefs[1])
                                pr.useReferenceTracker = true
                                origStack.set(i, pr)
                            } else {
                                val sh = ExcelTools.stripSheetNameFromRange(s)[0]
                                val rc = ExcelTools.getRangeRowCol(s)
                                val bRelRefs = booleanArrayOf(p.firstPtg!!.isRowRel, p.lastPtg!!.isRowRel, p.firstPtg!!.isColRel, p.lastPtg!!.isColRel)
                                if (bRelRefs[0])
                                    rc[0] += rowInc
                                if (bRelRefs[1])
                                    rc[2] += rowInc
                                if (bRelRefs[2])
                                    rc[1] += colInc
                                if (bRelRefs[3])
                                    rc[3] += colInc
                                val pa = PtgArea(false)
                                pa.parentRec = p.parentRec
                                if (sh != null)
                                    pa.location = sh + "!" + ExcelTools.formatRangeRowCol(rc, bRelRefs)
                                else
                                    pa.location = ExcelTools.formatRangeRowCol(rc, bRelRefs)
                                pa.useReferenceTracker = true
                                origStack.set(i, pa)
                            }
                        }
                    }
                } catch (ex: Exception) {
                    Logger.logErr("Formula.incrementSharedFormula: $ex")
                }

            }
        }

        /**
         * returns true if the value s is one of the Excel defined Error Strings
         *
         * @param s
         * @return
         */
        fun isErrorValue(s: String?): Boolean {
            return if (s == null) false else Collections.binarySearch(Arrays.asList("#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"), s.trim { it <= ' ' }) > -1
        }
    }


}