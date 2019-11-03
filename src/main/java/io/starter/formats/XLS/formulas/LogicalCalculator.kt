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


/*
    LogicalCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/
object LogicalCalculator {
    /**
     * AND
     *
     *
     * Returns TRUE if all its arguments are TRUE;
     * returns FALSE if one or more arguments is FALSE.
     *
     *
     * Syntax
     * AND(logical1,logical2, ...)
     *
     *
     * Logical1, logical2, ...   are 1 to 30 conditions you want to test
     * that can be either TRUE or FALSE.
     *
     *
     * The arguments must evaluate to logical values such as TRUE or FALSE,
     * or the arguments must be arrays or references that contain logical values.
     * If an array or reference argument contains text or empty cells,
     * those values are ignored.
     * If the specified range contains no logical values,
     * AND returns the #VALUE! error value.
     */
    internal fun calcAnd(operands: Array<Ptg>): Ptg {
        val b = true
        val alloperands = PtgCalculator.getAllComponents(operands)
        for (i in alloperands.indices) {
            if (alloperands[i] is PtgBool) {
                val bo = alloperands[i] as PtgBool
                val bool = bo.value as Boolean
                if (bool.booleanValue() == false) {
                    return PtgBool(false)
                }
            } else {
                // probably a ref, hopefully to a bool
                val s = alloperands[i].value.toString()
                if (s.equals("false", ignoreCase = true)) return PtgBool(false)
            }
        }
        return PtgBool(true)
    }

    /**
     * IF
     * Returns one value if a condition you specify evaluates to
     * TRUE and another value if it evaluates to FALSE.
     *
     *
     * Use IF to conduct conditional tests on values and formulas.
     *
     *
     * Syntax 1
     *
     *
     * IF(logical_test,value_if_true,value_if_false)
     *
     *
     * Logical_test   is any value or expression that can be evaluated to TRUE or FALSE.
     *
     *
     * Value_if_true   is the value that is returned if logical_test is TRUE. If logical_test is TRUE and value_if_true is omitted, TRUE is returned. Value_if_true can be another formula.
     *
     *
     * Value_if_false   is the value that is returned if logical_test is FALSE. If logical_test is FALSE and value_if_false is omitted, FALSE is returned. Value_if_false can be another formula.
     *
     *
     * Remarks
     *
     *
     * Up to seven IF functions can be nested as value_if_true and value_if_false
     * arguments to construct more elaborate tests. See the following last example.
     * When the value_if_true and value_if_false arguments are evaluated,
     * IF returns the value returned by those statements.
     * If any of the arguments to IF are arrays, every element of the array
     * is evaluated when the IF statement is carried out. If some of the
     * value_if_true and value_if_false arguments are action-taking functions,
     * all of the actions are taken.
     */
    internal fun calcIf(operands: Array<Ptg>): Ptg {
        // lets assume for now there are always 3 operands.. NOPE!  sometimes the missarg gets
        // lost for some reason, so we need to treat it like that.
        if (operands.size < 2)
            return PtgErr(PtgErr.ERROR_VALUE)
        val determine = operands[0]
        var iftrue: Ptg
        if (operands.size > 1)
            iftrue = operands[1]    // 20070212 KSC: if this is blank, return 0 (according to help!)
        else {
            // return strings
            return if (operands[0] is PtgStr) operands[0] else PtgInt(0)
        }
        var iffalse: Ptg
        if (operands.size > 2) {
            iffalse = operands[2]
        } else {
            iffalse = PtgMissArg()
        }
        if (iftrue is PtgMissArg) {
            iftrue = PtgNumber(0.0)
        }
        if (iffalse is PtgMissArg) {
            iffalse = PtgNumber(0.0)
        }
        if (determine !is PtgArray) {
            var strval: String? = null
            if (determine !is PtgRef)
                strval = determine.toString()
            else {
                try {
                    strval = determine.value.toString()
                } catch (e: Exception) {    // could be a formula not found error, etc. don't ignore
                    return PtgErr(PtgErr.ERROR_VALUE)
                }

            }
            return if (strval.equals("true", ignoreCase = true)) {
                iftrue
            } else {
                iffalse
            }
        } else {
            try {    // see what type of operands iftrue and iffalse arrays are
                var retArry = ""
                val p = determine.components
                //			boolean trueIsArray= iftrue instanceof
                var res = true
                val trueValueIsArray = iftrue is PtgArray
                val falseValueIsArray = iffalse is PtgArray
                for (i in p.indices) {
                    res = p[i].toString().equals("true", ignoreCase = true)
                    if (res) {
                        if (trueValueIsArray)
                            retArry = retArry + (iftrue as PtgArray).arrVals[i].toString() + ","
                        else
                            retArry = "$retArry$iftrue,"
                    } else { // false
                        if (falseValueIsArray)
                            retArry = retArry + (iffalse as PtgArray).arrVals[i].toString() + ","
                        else
                            retArry = "$retArry$iffalse,"
                    }
                }
                retArry = "{" + retArry.substring(0, retArry.length - 1) + "}"
                val pa = PtgArray()
                pa.setVal(retArry)
                return pa
            } catch (e: Exception) {    // this should hit if iftrue and iffalse are array types .... TODO: handle!
                return iffalse
            }

        }
    }

    /**
     * Returns the logical function False
     */
    internal fun calcFalse(operands: Array<Ptg>): Ptg {
        return PtgBool(false)
    }

    /**
     * Returns the logical function true
     */
    internal fun calcTrue(operands: Array<Ptg>): Ptg {
        return PtgBool(true)
    }


    /**
     * Returns the opposite boolean
     */
    internal fun calcNot(operands: Array<Ptg>): Ptg {
        return if (operands[0].value.toString().equals("false", ignoreCase = true)) {
            PtgBool(true)
        } else {
            PtgBool(false)
        }
    }

    /**
     * Returns the opposite boolean
     */
    internal fun calcOr(operands: Array<Ptg>): Ptg {
        val alloperands = PtgCalculator.getAllComponents(operands)
        // KSC: TESTING
        //System.out.print("\tOR " + operands[0].toString() + " " + operands[1].toString() + "? ");
        for (i in alloperands.indices) {
            if (alloperands[i].value.toString().equals("true", ignoreCase = true)) {
                return PtgBool(true)
            }
        }
        return PtgBool(false)
    }


    /**
     * IFERROR function
     * Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula.
     * Use the IFERROR function to trap and handle errors in a formula (formula: A sequence of values, cell references, names, functions, or operators in a cell that together produce a new value. A formula always begins with an equal sign (=).).
     *
     *
     * Value   is the argument that is checked for an error.
     *
     *
     * Value_if_error   is the value to return if the formula evaluates to an error. The following error types are evaluated: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!.
     *
     *
     * Remarks
     *
     *
     * If value or value_if_error is an empty cell, IFERROR treats it as an empty string value ("").
     * If value is an array formula, IFERROR returns an array of results for each cell in the range specified in value. See the second example below.
     * Example: Trapping division errors by using a regular formula
     */
    internal fun calcIferror(operands: Array<Ptg>?): Ptg {
        if (operands == null || operands.size != 2)
            return PtgErr(PtgErr.ERROR_VALUE)
        if (operands[0] !is PtgArray) {
            val ret = InformationCalculator.calcIserror(operands) as PtgBool
            return if (ret.booleanValue)
            // it's an error
                operands[1]
            else
            // it's not; return calculated results of 1st operand
                operands[0]
        } else {
            val components = operands[0].components
            var retArray = ""
            for (i in components.indices) {
                val test = arrayOfNulls<Ptg>(1)
                test[0] = components[i]
                val ret = InformationCalculator.calcIserror(test) as PtgBool
                if (ret.booleanValue)
                // it's an error
                    retArray = retArray + operands[1] + ","
                else
                // it's not; return calculated results of 1st operand
                    retArray = retArray + test[0] + ","
            }
            retArray = "{" + retArray.substring(0, retArray.length - 1) + "}"
            val pa = PtgArray()
            pa.setVal(retArray)
            return pa
        }
    }

}