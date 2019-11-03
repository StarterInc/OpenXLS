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


import io.starter.OpenXLS.DateConverter
import io.starter.formats.XLS.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.Calendar

abstract class GenericPtg : Ptg, Cloneable {

    internal var doublePrecision = 0.00000001        // doubles/floats cannot be compared for exactness so use precision comparator
    override var opcode: Byte = 0
        internal set
    /**
     * So, here you see we can get the static type from the record itself
     * then format the output record.  Some shorthand techniques are shown.
     */
    override var record: ByteArray
        internal set

    internal var vars: Array<Ptg>? = null
    /**
     * a locking mechanism so that Ptgs are not endlessly
     * re-calculated
     *
     * @return
     */
    /**
     * a locking mechanism so that Ptgs are not endlessly
     * re-calculated
     *
     * @return
     */
    override var lock = -1
    /**
     * returns the Location Policy of the Ptg is locked
     * used during automated BiffRec movement updates
     *
     * @return int
     */
    /**
     * lock the Location of the Ptg so that it will not
     * be updated during automated BiffRec movement updates
     *
     * @param b setting of the lock the location policy for this Ptg
     */
    override var locationPolicy = Ptg.PTG_LOCATION_POLICY_UNLOCKED
    /**
     * @return Returns the trackercell.
     */
    /**
     * @param trackercell The trackercell to set.
     */
    override var trackercell: BiffRec? = null

    // determine behavior
    override val isOperator: Boolean
        get() = false

    override val isBinaryOperator: Boolean
        get() = false

    override val isUnaryOperator: Boolean
        get() = false

    override val isStandAloneOperator: Boolean
        get() = false

    override val isPrimitiveOperator: Boolean
        get() = false

    override val isOperand: Boolean
        get() = false

    override val isFunction: Boolean
        get() = false

    override val isControl: Boolean
        get() = false

    override val isArray: Boolean
        get() = false

    override val isReference: Boolean
        get() = false

    /**
     * Returns the number of Params to pass to the Ptg
     */
    override val numParams: Int
        get() = if (isPrimitiveOperator) 2 else 0

    /*
        Return all of the cells in this range as an array
        of Ptg's.  This is used for range calculations.
    */
    override val components: Array<Ptg>?
        get() = null

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    override// KSC: added to skip External name reference for Add-in Formulas
    // 20060408 KSC: added quoting in PtgStr.getTextString
    //	                    if (vars[x] instanceof PtgStr) // 20060214 KSC: Quote string params
    //	                    	part= "\"" + part + "\"";
            /*if(!part.equals(""))*/// don't strip 1st paren if no params!  20060501 KSC
    // strip trailing comma
    val textString: String
        get() {

            var strx: String? = ""

            try {
                strx = string
            } catch (e: Exception) {
                Logger.logErr("Function not supported: " + this.parentRec!!.toString())
            }

            if (strx == null)
                return ""

            val out = StringBuffer(strx)
            if (vars != null) {
                val numvars = vars!!.size
                if (this.isPrimitiveOperator && this.isUnaryOperator) {
                    if (numvars > 0)
                        out.append(vars!![0].textString)

                } else if (this.isPrimitiveOperator) {
                    out.setLength(0)
                    for (x in 0 until numvars) {
                        out.append(vars!![x].textString)
                        if (x + 1 < numvars) out.append(this.string)
                    }
                } else if (this.isControl) {
                    for (x in 0 until numvars) {
                        out.append(vars!![x].textString)
                    }
                } else {
                    for (x in vars!!.indices) {
                        if (!(x == 0 && vars!![x] is PtgNameX)) {
                            val part = vars!![x].textString
                            out.append(part)
                            out.append(",")
                        }
                    }
                    if (vars!!.size > 0)
                        out.setLength(out.length - 1)
                }
            }
            out.append(string2)
            return out.toString()
        }

    /*text1 and 2 for this Ptg
     */
    override val string: String
        get() = toString()

    /**
     * return the human-readable String representation of
     * the "closing" portion of this Ptg
     * such as a closing parenthesis.
     */

    override val string2: String
        get() {
            if (this.isPrimitiveOperator) return ""
            return if (this.isOperator) ")" else ""
        }

    /**
     * Gets the (return) value of this Ptg as an operand Ptg.
     */
    override val ptgVal: Ptg
        get() {
            val value = this.value
            return value as? Ptg ?: if (value is Boolean)
                PtgBool(value.booleanValue())
            else if (value is Int)
                PtgInt(value.toInt())
            else if (value is Number)
                PtgNumber(value.toDouble())
            else if (value is String)
                PtgStr(value as String?)
            else
                PtgErr(PtgErr.ERROR_VALUE)
        }

    /**
     * returns the value of an operand ptg.
     *
     * @return null for non-operand Ptg.
     */
    override val value: Any?
        get() = null

    /**
     * Gets the value of the ptg represented as an int.
     *
     *
     * This can result in loss of precision for floating point values.
     *
     *
     * overridden in PtgInt to natively return value.
     *
     * @return integer representing the ptg, or NAN
     */
    override// we should be throwing something better
    // don't report an error if it's already an error
    ///  RIIIIGHT!  throw new FormulaCalculationException();
    val intVal: Int
        get() {
            try {
                return Double(this.value!!.toString()).toInt()
            } catch (e: NumberFormatException) {
                if (this !is PtgErr)
                    Logger.logErr("GetIntVal failed for formula: " + this.parentRec!!.toString() + " " + e)
                return 0
            }

        }

    /**
     * Gets the value of the ptg represented as an double.
     *
     *
     * This can result in loss of precision for floating point values.
     *
     *
     * NAN will be returned for values that are not translateable to an double
     *
     *
     * overrideen in PtgNumber
     *
     * @return integer representing the ptg, or NAN
     */
    override// Logger.logWarn("Error in Ptg Calculator getting Double Value: " + e3);
    val doubleVal: Double
        get() {
            var pob: Any? = null
            var d: Double? = null
            try {
                pob = this.value
                if (pob == null) {
                    Logger.logErr("Unable to calculate Formula at " + this.location!!)
                    return java.lang.Double.NaN
                }
                d = pob as Double?
            } catch (e: ClassCastException) {
                try {
                    val f = pob as Float?
                    d = f!!.toDouble()
                } catch (e2: ClassCastException) {
                    try {
                        val `in` = pob as Int?
                        d = `in`!!.toDouble()
                    } catch (e3: Exception) {
                        if (pob == null || pob.toString() == "") {
                            d = 0
                        } else {
                            try {
                                return Double(pob.toString())
                            } catch (e4: Exception) {
                                return java.lang.Double.NaN
                            }

                        }
                    }

                }

            } catch (exp: Throwable) {
                Logger.logErr("Unexpected Exception in PtgCalculator.getDoubleValue()", exp)
            }

            return d!!.toDouble()
        }

    // these do nothing here...
    override var location: String?
        @Throws(FormulaNotFoundException::class)
        get() = null
        set(s) {}

    override val intLocation: IntArray?
        @Throws(FormulaNotFoundException::class)
        get() = null

    // Parent Rec is the BiffRec record referenced by Operand Ptgs
    override var parentRec: XLSRecord? = null

    //TODO: PtgRef.isBlank should override!
    override val isBlank: Boolean
        get() = false


    override fun clone(): Any? {
        try {
            return super.clone()
        } catch (e: CloneNotSupportedException) {
            // This is, in theory, impossible
            return null
        }

    }

    /**
     * update the Ptg
     */
    override fun updateRecord() {

    }

    /* ################################################### EXPLANATION ###################################################

    1. set string varetvar in all Ptgs
    2. varetvar goes between ptg return vals if any
    3. if this is a funcvar then we loop ptgs and out
    4. when we call getString or evaluate, we loop into the
        recursive tree and execute on up.

   ################################################### EXPLANATION ###################################################*/


    /**
     * Operator Ptgs take other Ptgs as arguments
     * so we need to pass them in to get a meaningful
     * value.
     */
    override fun setVars(parr: Array<Ptg>) {
        this.vars = parr
    }

    /**
     * pass  in arbitrary number of values (probably other Ptgs)
     * and return the resultant value.
     *
     *
     * This effectively calculates the Expression.
     */
    override fun evaluate(obj: Array<Any>): Any {
        // do something useful
        return this.string
    }

    open fun init(b: ByteArray) {
        opcode = b[0]
        record = b
    }

    /**
     * return a Ptg  consisting of the calculated values
     * of the ptg's passed in.  Returns null for any non-operand
     * ptg.
     *
     * @throws CalculationException
     */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    override fun calculatePtg(parsething: Array<Ptg>): Ptg? {
        return null

    }

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     * tells the Ptg to get it on its own...
     */
    override fun updateAddressFromTrackerCell() {
        this.initTrackerCell()
        val trk = trackercell
        if (trk != null) {
            val nad = trk.cellAddress
            location = nad
        }
    }

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     * tells the Ptg to get it on its own...
     */
    override fun initTrackerCell() {
        if (trackercell == null) {
            try {
                val trk = this.parentRec!!.sheet!!.getCell(this.location)
                trackercell = trk
            } catch (e: Exception) {
                Logger.logWarn("Formula reference could not initialize:$e")
            }

        }
    }

    /**
     * generic reading of a row byte pair with handling for Excel 2007 if necessary
     *
     * @param b0
     * @param b1
     * @return int row
     */
    fun readRow(b0: Byte, b1: Byte): Int {
        if (parentRec != null && !parentRec!!.workBook!!.isExcel2007) {
            var rw = io.starter.toolkit.ByteTools.readInt(b0, b1, 0.toByte(), 0.toByte())
            if (rw >= XLSConstants.MAXROWS_BIFF8 - 1 || rw < 0 || this is PtgRefN)
            // PtgRefN's are ALWAYS relative and therefore never over 32xxx
                rw = ByteTools.readShort(b0.toInt(), b1.toInt()).toInt()
            return rw
        }
        // issue when reading Excel2007 rw from bytes as limits exceed ... try to interpret as best one can
        var rw = io.starter.toolkit.ByteTools.readInt(b0, b1, 0.toByte(), 0.toByte())
        if (rw == 65535) {    // have to assume that this means a wholeCol reference
            rw = -1
            (this as PtgRef).isWholeCol = true
        }
        return rw
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        parentRec = null
        trackercell = null
        // vars??

    }

    companion object {
        val serialVersionUID = 666555444333222L

        /**
         * Returns an array of doubles from number-type ptg's sent in.
         * This should only be referenced by sub-classes.
         *
         *
         * Null values accessed are treated as 0.  Within excel (empty cell values == 0) Tested!
         * Sometimes as well you can get empty string values, "".  These are NOT EQUAL ("" != 0)
         *
         * @param pthings
         * @return
         */
        protected fun getValuesFromPtgs(pthings: Array<Ptg>): Array<Any>? {
            val obar = arrayOfNulls<Any>(pthings.size)
            for (t in obar.indices) {
                if (pthings[t] is PtgErr)
                    return null
                if (pthings[t] is PtgArray) {
                    obar[t] = pthings[t].components    // get all items in array as Ptgs
                    var v: Any? = null
                    try {
                        v = getValuesFromObjects(obar[t] as Array<Any>)    // get value array from the ptgs
                    } catch (e: NumberFormatException) {    // string or non-numeric values
                        v = getStringValuesFromPtgs(obar[t] as Array<Ptg>)
                    }

                    obar[t] = v
                } else {
                    val pval = pthings[t].value
                    if (pval is PtgArray) {
                        obar[t] = pval.components    // get all items in array as Ptgs
                        var v: Any? = null
                        try {
                            v = getValuesFromObjects(obar[t] as Array<Any>)    // get value array from the ptgs
                        } catch (e: NumberFormatException) {    // string or non-numeric values
                            v = getStringValuesFromPtgs(obar[t] as Array<Ptg>)
                        }

                        obar[t] = v
                    } else if (pval is Name) {    // then get it's components ...
                        obar[t] = pthings[t].components
                        var v: Any? = null
                        try {
                            v = getValuesFromPtgs(obar[t] as Array<Ptg>)    // get value array from the ptgs
                        } catch (e: NumberFormatException) {    // string or non-numeric values
                            v = getStringValuesFromPtgs(obar[t] as Array<Ptg>)
                        }

                        obar[t] = v
                    } else {    // it's a single value
                        try {
                            obar[t] = getDoubleValueFromObject(pval)
                        } catch (e: NumberFormatException) {
                            if (pval is CalculationException)
                                obar[t] = pval.toString()
                            else
                                obar[t] = pval
                        }

                    }
                }
            }
            return obar
        }


        /**
         * Returns an array of doubles from number-type ptg's sent in.
         * This should only be referenced by sub-classes.
         *
         *
         * Null values accessed are treated as 0.  Within excel (empty cell values == 0) Tested!
         * Sometimes as well you can get empty string values, "".  These are NOT EQUAL ("" != 0)
         *
         * @param pthings
         * @return
         */
        @Throws(NumberFormatException::class)
        protected fun getValuesFromObjects(pthings: Array<Any>): DoubleArray {
            val returnDbl = DoubleArray(pthings.size)
            for (i in pthings.indices) {

                // Object o = pthings[i].getValue();
                val o = pthings[i]

                if (o == null) {    // NO!! "" is NOT "0", blank is, but not a zero length string.  Causes calc errors, need to handle diff somehow20081103 KSC: don't error out if "" */
                    returnDbl[i] = 0.0
                } else if (o is Double) {
                    returnDbl[i] = o.toDouble()
                } else if (o is Int) {
                    returnDbl[i] = o.toInt().toDouble()
                } else if (o is Boolean) {    // Excel converts booleans to numbers in calculations 20090129 KSC
                    returnDbl[i] = if (o.booleanValue()) 1.0 else 0.0
                } else if (o is PtgBool) {
                    returnDbl[i] = if ((o.value as Boolean).booleanValue()) 1.0 else 0.0
                } else if (o is PtgErr) {
                    // ?
                } else {
                    val s = o.toString()
                    val d = Double(s)
                    returnDbl[i] = d
                }
            }
            return returnDbl
        }

        /**
         * convert a value to a double, throws exception if cannot
         *
         * @param o
         * @throws NumberFormatException
         * @return double value if possible
         */
        @Throws(NumberFormatException::class)
        fun getDoubleValue(o: Any?, parent: XLSRecord?): Double {
            if (o is Double)
                return o.toDouble()
            if (o == null || o.toString() == "") {
                // empty string is interpreted as 0 if show zero values
                if (parent != null && parent.sheet!!.window2!!.showZeroValues)
                    return 0.0
                // otherwise, throw error
                throw NumberFormatException()
            }
            return Double(o.toString())    // will throw NumberFormatException if cannot convert
        }

        /**
         * converts a single Ptg number-type value to a double
         */
        fun getDoubleValueFromObject(o: Any?): Double {
            var ret = 0.0
            if (o == null) {    // 20081103 KSC: don't error out if "" */
                ret = 0.0
            } else if (o is Double) {
                ret = o.toDouble()
            } else if (o is Int) {
                ret = o.toInt().toDouble()
            } else if (o is Boolean) {    // Excel converts booleans to numbers in calculations 20090129 KSC
                ret = if (o.booleanValue()) 1.0 else 0.0
            } else if (o is PtgErr) {
                // ?
            } else {
                val s = o.toString()
                // handle formatted dates from fields like TEXT() calcs
                if (s.indexOf("/") > -1) {
                    try {
                        val c = DateConverter.convertStringToCalendar(s)
                        if (c != null) ret = DateConverter.getXLSDateVal(c)
                    } catch (e: Exception) {//guess not
                    }

                }
                if (ret == 0.0) {
                    val d = Double(s)
                    ret = d
                }
            }
            return ret
        }

        /**
         * returns an array of strings from ptg's sent in.
         * This should only be referenced by sub-classes.
         */
        protected fun getStringValuesFromPtgs(pthings: Array<Ptg>): Array<String> {
            val returnStr = arrayOfNulls<String>(pthings.size)
            for (i in pthings.indices) {
                if (pthings[i] is PtgErr)
                    return arrayOf("#VALUE!")    // 20081202 KSC: return error value ala Excel

                val o = pthings[i].value
                if (o != null) { // 20070215 KSC: avoid nullpointererror
                    try {    // 20090205 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
                        returnStr[i] = (o as Double).toInt().toString()
                    } catch (e: Exception) {
                        val s = o.toString()
                        returnStr[i] = s
                    }

                } else
                    returnStr[i] = "null" // 20070216 KSC: Shouldn't match empty string!
            }
            return returnStr
        }

        /**
         * return properly quoted sheetname
         *
         * @param s
         * @return
         */
        fun qualifySheetname(s: String?): String? {
            if (s == null || s == "") return s
            try {
                if (s[0] != '\'' && (s.indexOf(' ') > -1 || s.indexOf('&') > -1 || s.indexOf(',') > -1 || s.indexOf('(') > -1)) {
                    return if (s.indexOf("'") == -1) "'$s'" else "\"" + s + "\""
                }
            } catch (e: StringIndexOutOfBoundsException) {
            }

            return s
        }

        /**
         * return cell address with $'s e.g.
         * cell AB12 ==> $AB$12
         * cell Sheet1!C2=>Sheet1!$C$2
         * Does NOT handle ranges
         *
         * @param s
         * @return
         */
        fun qualifyCellAddress(s: String): String {
            var s = s
            var prefix = ""
            if (s.indexOf("$") == -1) {    // it's not qualified yet
                var i = s.indexOf("!")
                if (i > -1) {
                    prefix = s.substring(0, i + 1)
                    s = s.substring(i + 1)
                }
                s = "$$s"
                i = 1
                while (i < s.length && !Character.isDigit(s[i++]));
                i--
                if (i > 0 && i < s.length)
                    s = s.substring(0, i) + "$" + s.substring(i)
            }
            return prefix + s
        }

        fun getArrayLen(o: Any): Int {
            var len = 0
            if (o is DoubleArray)
                len = o.size
            return len
        }
    }


} 