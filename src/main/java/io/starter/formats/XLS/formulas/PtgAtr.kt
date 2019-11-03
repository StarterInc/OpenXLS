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

import io.starter.formats.XLS.Formula
import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.toolkit.ByteTools

/**
 * Displays "special" attributes like spaces and "optimized SUMs"
 *
 *
 * Offset Size Contents
 * 0 1 19H
 * 1 1 Attribute type flags:
 * 01H = This is a tAttrVolatile token (volatile function)
 * 02H = This is a tAttrIf token (IF function control)
 * 04H = This is a tAttrChoose token (CHOOSE function control)
 * 08H = This is a tAttrSkip token (skip part of token array)
 * 10H = This is a tAttrSum token (SUM function with one parameter)
 * 20H = This is a tAttrAssign token (assignment-style formula in a macro sheet)
 * 40H = This is a tAttrSpace token (spaces and carriage returns, BIFF3-BIFF8)
 * 41H = This is a tAttrSpaceVolatile token (BIFF3-BIFF8, see below)
 * 2 var. Additional information dependent on the attribute type
 *
 *
 * tAttrSpace:
 * 0 1 19H
 * 1 1 40H (identifier for the tAttrSpace token), or
 * 41H (identifier for the tAttrSpaceVolatile token)
 * 2 1 Type and position of the inserted character(s):
 * 00H = Spaces before the next token (not allowed before tParen token)
 * 01H = Carriage returns before the next token (not allowed before tParen token)
 * 02H = Spaces before opening parenthesis (only allowed before tParen token)
 * 03H = Carriage returns before opening parenthesis (only allowed before tParen token)
 * 04H = Spaces before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens)
 * 05H = Carriage returns before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens)
 * 06H = Spaces following the equality sign (only in macro sheets)
 * 3 1 Number of inserted spaces or carriage returns
 *
 * @see Ptg
 *
 * @see Formula
 */
class PtgAtr : GenericPtg, Ptg {
    internal var grbit: Byte = 0x0
    //    int bitAttrSemi     = 0;		// changed to BitAttrVolatile
    internal var bitAttrVolatile = 0
    internal var bitAttrIf = 0
    internal var bitAttrChoose = 0
    internal var bitAttrGoto = 0        // == bitAttrSkip
    internal var bitAttrSum = 0
    internal var bitAttrAssign = 0        // changed from bitAttrBaxcel
    internal var bitAttrSpace = 0
    internal var bitAttrSpaceVolatile = 0    // added

    override/* We may not want any text from this record.. *///"IF("; this is already taken care of by another ptg.
    // 	if(bitAttrSemi      > 0)   return "SEMI(";
    // this may be wrong, but as far as I can tell it is just internal for calc purposes.
    val string: String
        get() {
            this.init()
            if (bitAttrIf > 0) return ""
            if (bitAttrSum > 0) return "SUM("
            if (bitAttrVolatile > 0) return ""
            if (bitAttrAssign > 0) return " EQUALS "
            if (bitAttrChoose > 0) return "CHOOSE("
            if (bitAttrSpace > 0) return " "
            return if (bitAttrGoto > 0) "" else "UNKNOWN("


        }

    /**
     * return the human-readable String representation of
     * the "closing" portion of this Ptg
     * such as a closing parenthesis.
     */
    override//")"; this is already taken care of by another ptg.
    //        if(bitAttrSemi      > 0)   return ")";
    // this may be wrong, but as far as I can tell it is just internal for calc purposes.
    val string2: String
        get() {
            if (bitAttrIf > 0) return ""
            if (bitAttrSum > 0) return ")"
            if (bitAttrVolatile > 0) return ""
            if (bitAttrAssign > 0) return ")"
            if (bitAttrChoose > 0) return ")"
            return if (bitAttrGoto > 0) "" else ""
        }

    override/* TODO: Rework bitAttrIf.  It optimizes the calculation of if statements
         * should not normally be a big deal, but saves the calculation of one of
         * the result fields if needed.*///        if(bitAttrSemi   > 0)return false;
    val isControl: Boolean
        get() {
            this.init()
            if (isPrimitiveOperator) return false
            if (bitAttrIf > 0) return false
            if (bitAttrSum > 0) return true
            if (bitAttrVolatile > 0) return false
            if (bitAttrAssign > 0) return true
            if (bitAttrGoto > 0) return false
            return if (bitAttrChoose > 0) false else false
        }

    /**
     * is the space special -- does it go between vars?
     * for now we say sure why not.
     */
    override val isPrimitiveOperator: Boolean
        get() {
            this.init()
            return bitAttrSpace > 0
        }

    override val isUnaryOperator: Boolean
        get() = if (bitAttrIf > 0) false else bitAttrChoose <= 0


    override val isOperator: Boolean
        get() = false

    val isSpace: Boolean
        get() = bitAttrSpace > 0

    override//
    // if(bitAttrSpace > 0)return true;
    // Old version?		if(bitAttrSemi   > 0)return true; // it just shows that this is a volatile function
    // it just shows that this is a volatile function
    val isOperand: Boolean
        get() = bitAttrVolatile > 0

    override val length: Int
        get() = Ptg.PTG_ATR_LENGTH


    internal var alloperands: Array<Ptg>? = null // cached!

    constructor() {}

    constructor(type: Byte) {
        record = ByteArray(4)
        record[0] = 0x19
        record[1] = type
    }

    override fun toString(): String {
        return this.string + this.string2
    }

    /*
        Sets the grbit for the record
    *///[25, 2, 10, 0]		grbit= 2
    fun init() {
        grbit = this.record[1]
        /* john, the following syntax was not reliable, switched with syntax below....
           bitAttrIf       = ((grbit &     0x2)  >>  4);
        */
        // 20060501 KSC: Changed bitAttrVolatile operation + some names
        //        if ((grbit & 0x1)== 0x1){bitAttrSemi = 1;}
        if (grbit and 0x1 == 0x1) {
            bitAttrVolatile = 1
        }  // volatile= a function that needs to be recalculated always, such as NOW()
        if (grbit and 0x2 == 0x2) {
            bitAttrIf = 1
        }
        if (grbit and 0x4 == 0x4) {
            bitAttrChoose = 1
        }
        if (grbit and 0x8 == 0x8) {
            bitAttrGoto = 1
        }
        if (grbit and 0x10 == 0x10) {
            bitAttrSum = 1
        }
        if (grbit and 0x20 == 0x20) {
            bitAttrAssign = 1
        }    // changed name from bitAttrBaxcel
        if (grbit and 0x40 == 0x40) {
            bitAttrSpace = 1
        }
        if (grbit and 0x41 == 0x41) {
            bitAttrSpaceVolatile = 1
        }


    }

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     *
     *
     * public String getString(){
     * byte[] br = this.getRecord();
     * byte[] db = new byte[br.length-1]; // strip opcode
     * System.arraycopy(br, 1, db, 0, db.length);
     * return ""; // new String(db);
     * }
     */
    fun getLength(b: ByteArray): Int {
        if (b[0] and 0x4 == 0x4) {
            var i = ByteTools.readShort(b[1].toInt(), b[2].toInt()).toInt()
            i += 4
            return i
        }
        return length
    }

    /*
        Calculate the value of this ptg.
    */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    override fun calculatePtg(pthing: Array<Ptg>): Ptg? {
        var returnPtg: Ptg? = null
        if (this.bitAttrSum > 0) {
            returnPtg = MathFunctionCalculator.calcSum(pthing)
        }
        return returnPtg
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2825828785221803436L
    }
}