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

import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

import java.util.Locale

/**
 * PtgFunc is a fuction operator that refers to the header file in order to
 * use the correct function.
 *
 *
 * PtgFuncVar is only used with a variable number of arguments.
 *
 *
 * Opcode = 22h
 *
 * <pre>
 * Offset      Bits    Name        Mask        Contents
 * --------------------------------------------------------
 * 0           6-0     cargs       7Fh         The number of arguments to the function
 * 7       fPrompt     80h         =1, function prompts the user
 * 1           14-0    iftab       7FFFh       The index to the function table
 * see GenericPtgFunc for details
 * 15      fCE         8000h       This function is a command equivalent
</pre> *
 *
 * @see Ptg
 *
 * @see GenericPtgFunc
 */
class PtgFuncVar : GenericPtg, Ptg {
    // should be handled by super?
    override var opcode: Byte = 0
        internal set
    internal var loc: String? = null
    internal var cargs: Byte = 0
    internal var fprompt: Boolean = false
    /**
     * Get the function ID for this PtgFuncVar
     *
     * @return function Id
     */
    var functionId: Short = 0
        internal set
    internal var fCE: Boolean = false

    override val isFunction: Boolean
        get() = true

    /**
     * Returns the number of Params to pass to the Ptg
     */
    override val numParams: Int
        get() = cargs.toInt()


    /**
     * GetString - is this toString, what is it returning?
     */
    override val string: String
        get() {
            if (functionId.toInt() != FunctionConstants.xlfADDIN) {
                var f: String? = null
                if (Locale.JAPAN == Locale.getDefault()) {
                    f = FunctionConstants.getJFunctionString(functionId)
                }
                if (f == null) f = FunctionConstants.getFunctionString(functionId)
                return f
            }
            return addInFunctionString
        }

    // KSC: added to handle string version of add-in formulas
    private val addInFunctionString: String
        get() = if (vars != null && vars!![0] is PtgNameX) {
            vars!![0].toString() + "("
        } else "("

    override val string2: String
        get() = ")"


    val `val`: Int
        get() = functionId.toInt()


    // THis will have to be modified when we start modifying the record.
    override val record: ByteArray
        get() = record

    override val length: Int
        get() = Ptg.PTG_FUNCVAR_LENGTH


    constructor(funcType: Int, numArgs: Int, parentRec: XLSRecord) : this(funcType, numArgs) {
        this.parentRec = parentRec
    }

    constructor(funcType: Int, numArgs: Int) {
        val recbyte = ByteArray(4)
        // 20100609 KSC: there are three types of funcvars:
        // 22H (tFuncVarR), 42H (tFuncVarV), 62H (tFuncVarA)
        // tFUncVarR = reference return value (most common?)
        // tFuncVarV = value type of return value (ROW, SUM)
        // tFuncVarA = Array return type (TREND)
        // TODO: figure out which other functions are type V or type A
        when (funcType) {
            FunctionConstants.XLF_ROW        // ROW
                , FunctionConstants.xlfColumn        // COLUMN
                , FunctionConstants.xlfIndex    // INDEX
                , FunctionConstants.xlfVlookup    // VLOOKUP
                , FunctionConstants.xlfSumproduct    // SUMPRODUCT
            -> recbyte[0] = 0x42
            else        // default= tFuncVarR
            -> recbyte[0] = 0x22
        }
        val b = ByteTools.shortToLEBytes(funcType.toShort())
        recbyte[1] = numArgs.toByte()
        recbyte[2] = b[0]
        recbyte[3] = b[1]
        this.init(recbyte)
    }

    constructor() {
        // placeholder
    }

    /**
     * set the number of parmeters in the FuncVar record
     *
     * @param byte nParams
     */
    // 20060131 KSC: Added to set # params separately from init
    fun setNumParams(nParams: Byte) {
        record[1] = nParams
        this.populateVals()
    }

    override fun init(b: ByteArray) {
        opcode = 0x22
        record = b
        fprompt = false
        fCE = false
        this.populateVals()
    }

    /**
     * parse all the values out of the byte array and
     * populate the classes values
     */
    private fun populateVals() {

        cargs = record[1]
        if (cargs and 0x80 == 0x80) { // is fprompt set?
            fprompt = true
        }
        cargs = (cargs and 0x7f).toByte()
        functionId = ByteTools.readShort(record[2].toInt(), record[3].toInt())
        if (functionId and 0x8000 == 0x8000) { // is fCE set?
            fCE = true
        }
        functionId = (functionId and 0x7fff).toShort() // cut out the fCE
    }

    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    override fun calculatePtg(pthings: Array<Ptg>): Ptg? {
        val ptgarr = arrayOfNulls<Ptg>(pthings.size + 1)
        ptgarr[0] = this
        // add this into the array so the functionHandler has a handle to the function
        System.arraycopy(pthings, 0, ptgarr, 1, pthings.size)
        return FunctionHandler.calculateFunction(ptgarr)
    }

    /**
     * return String representation of function id for this funcvar
     */
    override fun toString(): String {
        return "FUNCVAR $functionId"
    }

    /**
     * given this specific Func Var, ensure that it's parameters are of the correct Ptg type
     * <br></br>Value, Reference or Array
     * <br></br>This is necessary when functions are added via String
     * <br></br>NOTE: eventually all FuncVars which require a specific type of parameter will be handled here
     *
     * @see FormulaParser.adjustParameterIds
     */
    fun adjustParameterIds() {
        if (vars == null) return  // no parameters to worry about
        when (functionId) {
            FunctionConstants.xlfVlookup -> {
                setParameterType(0, PtgRef.VALUE)
                setParameterType(1, PtgRef.REFERENCE)
                setParameterType(2, PtgRef.VALUE)
                setParameterType(3, PtgRef.VALUE)
            }
            FunctionConstants.xlfColumn, FunctionConstants.XLF_ROW -> setParameterType(0, PtgRef.REFERENCE)
            FunctionConstants.xlfIndex -> {
                setParameterType(0, PtgRef.REFERENCE)
                setParameterType(1, PtgRef.VALUE)
            }
            FunctionConstants.XLF_SUM_IF -> setParameterType(0, PtgRef.REFERENCE)
            FunctionConstants.xlfSumproduct -> {
                setParameterType(0, PtgRef.ARRAY)
                setParameterType(1, PtgRef.ARRAY)
            }
            else -> {
            }
        }
    }


    /**
     * utility for adjustParameterIds to set the PtgRef-type or PtgName-type pareameter to the correct type
     * either PtgRef.REFERENCE, PtgRef.VALUE or PtgRef.ARRAY
     * dependent upon the function they are used in
     *
     * @param n
     * @param type
     */
    private fun setParameterType(n: Int, type: Short) {
        if (vars!!.size > n) {
            if (vars!![n] is PtgArea) {
                (vars!![n] as PtgArea).setPtgType(type)
            } else if (vars!![n] is PtgRef) {
                (vars!![n] as PtgRef).setPtgType(type)
            } else if (vars!![n] is PtgName) {
                (vars!![n] as PtgName).setPtgType(type)
            }
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 1478629759437556620L
        var LENGTH = 3
    }
}