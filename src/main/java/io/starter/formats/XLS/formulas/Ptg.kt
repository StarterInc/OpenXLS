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

import io.starter.formats.XLS.*

import java.io.Serializable


/**
 * Ptg is the interface all ptgs implement in order to be handled equally under the
 * eyes of the all seeing one, "he that shall not be named"  A ptg is a unique segment
 * of a formula stack that indicates a value, a reference to a value, or an operation.
 * See the docs under Formula for more information.
 *
 * @see Ptg
 *
 * @see Formula
 */

interface Ptg : XLSConstants, Serializable {

    /**
     * return the length of the Ptg
     */
    val length: Int

    /**
     * return the number of parameters to this Ptg
     */
    val numParams: Int

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     *
     * @return trackercell The trackercell to set.
     */
    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     *
     * @param trackercell The trackercell to set.
     */
    var trackercell: BiffRec

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
    var lock: Int


    /**
     * determine the general Ptg type
     */
    val isOperator: Boolean

    val isBinaryOperator: Boolean

    val isUnaryOperator: Boolean

    val isStandAloneOperator: Boolean

    val isOperand: Boolean

    val isControl: Boolean

    val isFunction: Boolean

    val isReference: Boolean

    /**
     * determines whether this operator is a 'primitive' such as +,-,=,<,>,!=,==,etc.
     * the upshot is that primitives go BETWEEN operands, and non-primitives
     * encapsulate
     *
     *
     * ie:
     *
     *
     * SUM(A1:A4)  non-primitive
     * A1+A4       primitive
     */
    val isPrimitiveOperator: Boolean

    /*
       Determines whether the ptg represents multiple ptg's in reality.
       ie ptgArea ia actually a collection of ptgRef's, so ptgArea.getIsArray returns 'true'
    */
    val isArray: Boolean

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    val textString: String

    /**
     * If a record consists of multiple sub records (ie PtgArea) return those
     * records, else return null;
     */
    val components: Array<Ptg>

    /**
     * @return byte[] containing the whole ptg, including identifying opcode
     */
    val record: ByteArray

    /*
        @return XLSRecord containing the whole ptg
    */
    /**
     * constructor must pass in 'parent' XLSRecord so that there
     * is a handle for updating...
     *
     * @return
     */
    var parentRec: XLSRecord

    /**
     * returns whether the Location of the Ptg is locked
     * used during automated BiffRec movement updates
     *
     * @return location policy
     */
    /**
     * lock the Location of the Ptg so that it will not
     * be updated during automated BiffRec movement updates
     *
     * @param b whether to lock the location of this Ptg
     */
    var locationPolicy: Int

    /**
     * When the ptg is a reference to a location this returns that location
     *
     * @return Location
     */
    /**
     * setLocation moves a ptg that is a reference to a location, such as
     * a ptg range being modified
     *
     * @param String location, such as A1:D4
     */
    var location: String

    val intLocation: IntArray

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    val string: String

    /**
     * return the byte header for the Ptg
     */
    val opcode: Byte

    /**
     * return the human-readable String representation of
     * the "closing" portion of this Ptg
     * such as a closing parenthesis.
     */
    val string2: String

    /**
     * Gets the (return) value of this Ptg as an operand Ptg.
     */
    val ptgVal: Ptg

    /**
     * returns the value of an operand ptg.
     *
     * @return null for non-operand Ptg.
     */
    val value: Any

    /**
     * Gets the value of the ptg represented as an int.
     *
     *
     * This can result in loss of precision for floating point values.
     *
     *
     * -1 will be returned for values that are not translateable to an integer
     *
     * @return integer representing the ptg, or NAN
     */
    val intVal: Int

    val doubleVal: Double

    val isBlank: Boolean    // 20081112 KSC

    /**
     * Creates a deep clone of this Ptg.
     */
    override fun clone(): Any

    /**
     * update the values of the Ptg
     */
    fun updateRecord()

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     * tells the Ptg to get it on its own...
     */
    fun updateAddressFromTrackerCell()

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     * tells the Ptg to get it on its own...
     */
    fun initTrackerCell()

    /**
     * Operator Ptgs take other Ptgs as arguments
     * so we need to pass them in to get a meaningful
     * value.
     */
    fun setVars(parr: Array<Ptg>)

    /**
     * pass  in arbitrary number of values (probably other Ptgs)
     * and return the resultant value.
     *
     *
     * This effectively calculates the Expression.
     */
    fun evaluate(obj: Array<Any>): Any

    /**
     * return a Ptg  consisting of the calculated values
     * of the ptg's passed in.  Returns null for any non-operater
     * ptg.
     *
     * @throws CalculationException
     */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    fun calculatePtg(parsething: Array<Ptg>): Ptg

    fun close()

    companion object {

        /**
         * VALUE type Reference (Id=0x44)
         */
        val VALUE: Short = 0
        /**
         * REFERENCE type Reference (Id=0x24)
         */
        val REFERENCE: Short = 1
        /**
         * ARRAY type Reference (Id=0x64)
         */
        val ARRAY: Short = 2

        val PTG_LOCATION_POLICY_UNLOCKED = 0
        val PTG_LOCATION_POLICY_LOCKED = 1
        val PTG_LOCATION_POLICY_TRACK = 2


        val PTG_TYPE_SINGLE = 1 // single-byte record
        val PTG_TYPE_ARRAY = 2 // array of bytes record

        //ptg lengths
        val PTG_NUM_LENGTH = 9
        val PTG_ADD_LENGTH = 1
        val PTG_AREA_LENGTH = 9
        val PTG_AREA3D_LENGTH = 11
        val PTG_AREAERR3D_LENGTH = 11
        val PTG_ATR_LENGTH = 4
        val PTG_CONCAT_LENGTH = 1
        val PTG_DIV_LENGTH = 1
        val PTG_EQ_LENGTH = 1
        val PTG_EXP_LENGTH = 5
        val PTG_FUNC_LENGTH = 3
        val PTG_FUNCVAR_LENGTH = 4
        val PTG_GE_LENGTH = 1
        val PTG_GT_LENGTH = 1
        val PTG_INT_LENGTH = 3
        val PTG_ISECT_LENGTH = 1
        val PTG_LE_LENGTH = 1
        val PTG_LT_LENGTH = 1
        val PTG_MEMERR_LENGTH = 7
        val PTG_MEM_AREA_N_LENGTH = 7
        val PTG_MEM_AREA_NV_LENGTH = 7
        val PTG_MLT_LENGTH = 1
        val PTG_MYSTERY_LENGTH = 1
        val PTG_NE_LENGTH = 1
        val PTG_NAME_LENGTH = 5
        val PTG_NAMEX_LENGTH = 7
        val PTG_PAREN_LENGTH = 1
        val PTG_POWER_LENGTH = 1
        val PTG_RANGE_LENGTH = 1
        val PTG_REF_LENGTH = 5
        val PTG_REF3D_LENGTH = 7
        val PTG_REFERR_LENGTH = 5
        val PTG_REFERR3D_LENGTH = 7
        val PTG_ENDSHEET_LENGTH = 1
        val PTG_SUB_LENGTH = 1
        val PTG_UNION_LENGTH = 1
        val PTG_BOOL_LENGTH = 2
        val PTG_UPLUS_LENGTH = 1
        val PTG_UMINUS_LENGTH = 1
        val PTG_PERCENT_LENGTH = 1


        //TODO:  add all the opcodes here
        val PTG_INT: Byte = 0x1e

        val CALCULATED = 0
        val UNCALCULATED = -1
    }

}