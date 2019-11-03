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
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger

import java.util.Enumeration

/**
 * PtgCalculator handles some of the standard calls that all of the
 * calculator classes need.
 */


object PtgCalculator {

    /*
     * creates a generic error ptg
     */
    val error: Ptg
        get() = PtgErr(PtgErr.ERROR_NULL)

    /**
     * returns an #VALUE! error ptg
     *
     * @return
     */
    val valueError: Ptg
        get() = PtgErr(PtgErr.ERROR_VALUE)

    /**
     * returns an #NA! error ptg
     *
     * @return
     */
    val naError: Ptg
        get() = PtgErr(PtgErr.ERROR_NA)


    /**
     * getLongValue is for single-operand functions.
     * It returns NaN for calculations that have too many operands.
     *
     * @param operands
     * @return
     */
    internal fun getLongValue(operands: Array<Ptg>): Long {
        val components = operands[0].components
        if (components != null) { // check if too many operands TODO: check that ONE is ok??
            if (components.size > 1) {
                return java.lang.Double.NaN.toLong()
            }
        }
        if (operands.size > 1) { // not supported by function
            Logger.logWarn("PtgCalculator getting Long Value for operand failed: - UNSUPPORTED BY FUNCTION")
            return java.lang.Double.NaN.toLong()
        }
        var d: Double? = null
        try {
            d = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            Logger.logWarn("PtgCalculator getting Long Value for operand failed: $e")
            return java.lang.Double.NaN.toLong()
        }

        return d.toLong()
    }

    /*
     * See getLongValue(operand[])
     * Does the same thing with a single operand
     */
    internal fun getLongValue(operand: Ptg): Long {
        val ptgArr = arrayOfNulls<Ptg>(1)
        ptgArr[0] = operand
        return getLongValue(ptgArr)
    }


    // returns an array of longs from an array of ptg's
    internal fun getLongValueArray(operands: Array<Ptg>): LongArray {
        val alloperands = getAllComponents(operands)
        val l = LongArray(alloperands.size)
        for (i in alloperands.indices) {
            try {
                val dd = operands[i].doubleVal
                l[i] = dd.toLong()
            } catch (e: NumberFormatException) {
                Logger.logWarn("PtgCalculator getting Long value array failed: $e")
                l[i] = java.lang.Double.NaN.toLong()
            }

        }
        return l
    }

    /**
     * getDoubleValue is for multi-operand functions.  It returns NaN
     * for calculations that have to many operands.
     *
     * @param operands
     * @return
     * @throws CircularReferenceException TODO
     */
    @Throws(CalculationException::class)
    internal fun getDoubleValueArray(operands: Array<Ptg>): DoubleArray? {

        var d: Double? = null
        // we don't know the size ahead of time, so use a vector for now.
        val cv = CompatibleVector()

        for (i in operands.indices) {
            // is it multidimensional?
            val pthings = operands[i].components // optimized -- do it once!  -jm
            if (pthings != null) {
                for (x in pthings.indices) {
                    cv.add(pthings[x])
                }
            } else {
                cv.add(operands[i])
            }
        }

        val darr = DoubleArray(cv.size)
        var i = 0
        val en = cv.elements()
        while (en.hasMoreElements()) {
            d = null//new Double(0.0);		// 20081229 KSC: reset
            val pthing = en.nextElement() as Ptg
            val ob = pthing.value
            if (ob == null || ob.toString().trim { it <= ' ' } == "") {    // 20060802 KSC: added trim
                darr[i] = 0.0
            } else if (ob.toString() == "#CIR_ERR!") {
                throw CircularReferenceException(CalculationException.VALUE)
            } else {

                try {
                    if (ob is Double) {
                        d = ob
                    } else {
                        val s = ob.toString()
                        d = Double(s)
                    }
                } catch (e: NumberFormatException) {
                    try {
                        val s = ob.toString()
                        if (s == "#N/A") {    // 20090130 KSC: if error value, propagate error (ala Excel) -- null caught in calling method propagates "#N/A"
                            return null
                        }
                    } catch (ee: Exception) {
                        Logger.logWarn("PtgCalculator getting Double value array failed: $ee")
                        d = java.lang.Double.NaN
                    }

                }

                if (d != null)
                    darr[i] = d.toDouble()
            }
            i++
        }
        return darr
    }

    @Throws(CalculationException::class)
    internal fun getDoubleValueArray(operands: Ptg): DoubleArray? {
        val ptgarr = arrayOfNulls<Ptg>(1)
        ptgarr[0] = operands
        return getDoubleValueArray(ptgarr)
    }

    /**
     * return a 2-dimenional array of double values
     * i.e. keep array structure of reference and array type parameters
     * <br></br> ASSUMPTIONS:
     * <br></br> 1- accepts only 2-d references i.e. ranges are on 1 sheet
     * <br></br> 2- assumes that range reference are in proper notation i.e A1:B6, NOT B6:A1
     *
     * @param operand Ptg
     * @return double[][]
     */
    @Throws(Exception::class)
    internal fun getArray(operand: Ptg): Array<DoubleArray> {
        val nrows: Int
        val ncols: Int
        var arr: Array<DoubleArray>? = null

        if (operand is PtgRef) {
            val rc = operand.intLocation
            val sheet = operand.sheetName
            val bk = operand.parentRec.workBook
            nrows = rc!![2] - rc[0] + 1
            ncols = rc[3] - rc[1] + 1
            arr = Array(nrows) { DoubleArray(ncols) }
            for (j in rc[1]..rc[3]) {
                for (i in rc[0]..rc[2]) {
                    val cell = ExcelTools.formatLocation(intArrayOf(i, j))
                    arr[i - rc[0]][j - rc[1]] = bk!!.getCell(sheet, cell).dblVal
                }
            }

        } else { // should be an array
            var arrStr = operand.toString().substring(1)
            arrStr = arrStr.substring(0, arrStr.length - 1)
            val rows = arrStr.split(";".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            arr = arrayOfNulls(rows.size)
            for (i in rows.indices) {
                val s = rows[i].split(",".toRegex()).toTypedArray()    // include empty strings
                arr[i] = DoubleArray(s.size)
                for (j in s.indices) {
                    arr[i][j] = Double(s[j])
                }
            }
        }
        return arr
    }


    /*
     *  Get all components recurses through the ptg's and returns an array
     * of single ptg's for all the ptgs in the operands array.  This means it
     * converts arrays to ref's, etc.
     */
    internal fun getAllComponents(operands: Array<Ptg>): Array<Ptg> {

        if (operands.size == 1) {
            val ret = operands[0].components ?: return operands
        }

        val v = FastAddVector()
        for (i in operands.indices) {
            val pthings = operands[i].components // optimized -- do it once!  -jm
            if (pthings != null) {
                for (x in pthings.indices) {
                    v.add(pthings[x])
                }
            } else {
                v.add(operands[i])
            }
        }
        var res = arrayOfNulls<Ptg>(v.size)
        res = v.toTypedArray() as Array<Ptg>
        return res
    }

    /*
     *  Get all components recurses through the ptg's and returns an array
     * of single ptg's for all the ptgs in the operands array.  This means it
     * converts arrays to ref's, etc.
     */
    internal fun getAllComponents(operand: Ptg): Array<Ptg> {
        val ptgArr = arrayOfNulls<Ptg>(1)
        ptgArr[0] = operand
        return getAllComponents(ptgArr)
    }

    /*
     * Returns the boolean value of a PTG.  If no boolean available then it returns false
     * Add more types here if they are available/needed
     */
    internal fun getBooleanValue(operand: Ptg): Boolean {
        if (operand is PtgBool) {
            return operand.booleanValue
        } else if (operand is PtgInt) {
            return operand.booleanVal
        }
        return false
    }

}