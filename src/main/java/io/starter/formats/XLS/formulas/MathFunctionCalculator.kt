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

import io.starter.toolkit.Logger

import java.math.BigDecimal
import java.util.ArrayList
import java.util.Random


/*
    MathFunctionCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/

object MathFunctionCalculator {


    /**
     * SUM
     * Adds all the numbers in a range of cells.
     * Ignores non-number fields
     *
     *
     * Usage@ SUM(number1,number2, ...)
     * Return@ PtgNumber
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcSum(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        var result = 0.0
        val dub = PtgCalculator.getDoubleValueArray(operands) ?: return PtgCalculator.naError
        for (i in dub.indices) {
            result += dub[i]
        }
        return PtgNumber(result)
    }

    /**
     * ABS
     * Returns the absolute value of a number
     */
    internal fun calcAbs(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        dd = Math.abs(dd)
        return if (dd.isNaN()) PtgCalculator.error else PtgNumber(dd)

    }

    /**
     * ACOS
     * Returns the arccosine of a number
     */
    internal fun calcAcos(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        dd = Math.acos(dd)
        return if (dd.isNaN()) PtgCalculator.error else PtgNumber(dd)
    }

    /**
     * ACOSH
     * Returns the inverse hyperbolic cosine of a number
     */
    internal fun calcAcosh(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var x = 0.0
        try {
            x = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        val dd = Math.log(x + (1.0 + x) * Math.sqrt((x - 1.0) / (x + 1.0)))
        return if (dd.isNaN()) PtgCalculator.error else PtgNumber(dd)
    }

    /**
     * ASIN
     * Returns the arcsine of a number
     */
    internal fun calcAsin(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        dd = Math.asin(dd)
        return if (dd.isNaN()) PtgCalculator.error else PtgNumber(dd)
    }

    /**
     * ASINH
     * Returns the inverse hyperbolic sine of a number
     */
    internal fun calcAsinh(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var x = 0.0
        try {
            x = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (x.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        // KSC: TESTING: I BELIEVE THE CALCULATION IS NOT CORRECT!
        val bd = BigDecimal((if (x > 0.0) 1.0 else -1.0) * getAcosh(Math.sqrt(1.0 + x * x)))
        bd.setScale(15, BigDecimal.ROUND_HALF_UP)
        //	PtgNumber ptnum = new PtgNumber(dd);
        return PtgNumber(bd.toDouble())
    }


    /**
     * ATAN
     * Returns the arctangent of a number
     */
    internal fun calcAtan(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        dd = Math.atan(dd)
        return if (dd.isNaN()) PtgCalculator.error else PtgNumber(dd)
    }

    /**
     * ATAN2
     * Returns the arctangent from x- and y- coordinates
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcAtan2(operands: Array<Ptg>): Ptg {
        if (operands.size != 2) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
        val res = Math.atan2(dd[0], dd[1])
        return PtgNumber(res)
    }

    /**
     * ATANH
     * Returns the inverse hyperbolic tangent of a number
     */
    internal fun calcAtanh(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd > 1 || dd < -1) return PtgCalculator.error
        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        val res = 0.5 * Math.log((1.0 + dd) / (1.0 - dd))
        return PtgNumber(res)
    }

    /**
     * CEILING
     * Rounds a number to the nearest multiple of significance;
     * This takes 2 values, first the number to round, next the value of signifigance.
     * Ick.  This is pretty intensive, so maybe a better way?
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcCeiling(operands: Array<Ptg>): Ptg {
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        if (dd.size != 2) return PtgCalculator.naError
        val num = dd[0]
        val multiple = dd[1]
        var res = 0.0
        while (res < num) {
            res += multiple
        }
        return PtgNumber(res)
    }

    /**
     * COMBIN
     * Returns the number of combinations for a given number of objects
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcCombin(operands: Array<Ptg>): Ptg {
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        if (dd.size != 2) return PtgCalculator.naError
        val num1 = Math.round(dd[0])
        val num2 = Math.round(dd[1])
        if (num1 < num2) return PtgCalculator.error
        val res1 = stepFactorial(num1, num2.toInt())
        val res2 = factorial(num2)
        val res = (res1 / res2).toDouble()
        return PtgNumber(res)
    }

    /**
     * COS
     * Returns the cosine of a number
     */
    internal fun calcCos(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        val res = Math.cos(dd)
        return PtgNumber(res)
    }

    /**
     * COSH
     * Returns the hyperbolic cosine of a number
     */
    internal fun calcCosh(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        val res = 0.5 * (Math.exp(dd) + Math.exp(-dd))
        return PtgNumber(res)
    }

    /**
     * DEGREES
     * Converts radians to degrees
     */
    internal fun calcDegrees(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        val res = Math.toDegrees(dd)
        return PtgNumber(res)
    }

    /**
     * EVEN
     * Rounds a number up to the nearest even integer
     */
    internal fun calcEven(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        var resnum = Math.round(dd)
        val remainder = dd % 2
        if (remainder != 0.0) {
            resnum++
        }
        return PtgInt(resnum.toInt())
    }

    /**
     * EXP
     * Returns e raised to the power of number.
     * The constant e equals 2.71828182845904,
     * the base of the natural logarithm.
     *
     *
     * Example
     * EXP(2) equals e2, or 7.389056
     */
    internal fun calcExp(operands: Array<Ptg>): Ptg {
        if (operands.size > 1 || operands[0].components != null) { // not supported by function
            return PtgErr(PtgErr.ERROR_NULL)
        }
        val d = 2.718281828459045
        val p = operands[0]
        try {
            val dub = Double(p.value.toString())
            //	    double result = Math.pow(d, dub.doubleValue());
            val result = BigDecimal(Math.pow(d, dub))
            result.setScale(15, BigDecimal.ROUND_HALF_UP)
            //	    PtgNumber pnum = new PtgNumber(result);
            return PtgNumber(result.toDouble())
        } catch (e: NumberFormatException) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    /**
     * FACT
     * Returns the factorial of a number
     */
    internal fun calcFact(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd: Long = 0
        try {
            dd = PtgCalculator.getLongValue(operands)
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        val res = MathFunctionCalculator.factorial(dd)
        return PtgNumber(res.toDouble())
    }

    /**
     * FACTDOUBLE
     * Returns the double factorial of a number
     * Example
     * !!6 = 6*4*2
     * !!7 = 7*5*3*1;
     */
    internal fun calcFactDouble(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        val n = PtgCalculator.getLongValue(operands)
        if (n.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        if (n < 0)
            return PtgErr(PtgErr.ERROR_NUM)
        var res = 1.0
        val endPoint = (if (n % 2 == 0L) 2 else 1).toLong()
        var i = n
        while (i >= endPoint) {
            res *= i.toDouble()
            i -= 2
        }
        if (n == 0L)
            res = -1.0    // by convention ...
        // KSC: not quite right!!!	replaced with above
        //long res =  MathFunctionCalculator.doubleFactorial(n);
        return PtgNumber(res)
    }

    /**
     * FLOOR
     * Rounds a number down, toward zero.  Works just like Celing with two operands
     * See comment above for celing for more info.
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcFloor(operands: Array<Ptg>): Ptg {
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        if (dd.size != 2) return PtgCalculator.error
        val num = dd[0]
        val multiple = dd[1]
        var res = 0.0
        while (res < num) {
            res += multiple
        }
        // drop one from the celing code to get the floor...
        res -= multiple
        return PtgNumber(res)
    }

    /**
     * GCD
     * Returns the greatest common divisor
     * TODO: Finish!
     */
    internal fun calcGCD(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val numbers = PtgCalculator.getAllComponents(operands)
        var gcd: Long = 0
        try {
            val n1: Long
            if ((n1 = PtgCalculator.getLongValue(numbers[0])) < 0)
                return PtgErr(PtgErr.ERROR_NUM)

            gcd = n1
            for (i in 1 until numbers.size) {
                val n2: Long
                if ((n2 = PtgCalculator.getLongValue(numbers[i])) < 0)
                    return PtgErr(PtgErr.ERROR_NUM)

                var bigger: Long
                var smaller: Long
                var r: Long
                bigger = Math.max(n2, n1)
                smaller = Math.min(n2, n1)
                r = bigger % smaller
                while (r != 0L) {
                    bigger = smaller
                    smaller = r
                    r = bigger % smaller
                }
                gcd = Math.min(gcd, smaller)
            }
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

        return PtgNumber(gcd.toDouble())
    }

    /**
     * INT
     * Rounds a number down to the nearest integer
     */
    internal fun calcInt(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        var res = Math.round(dd)
        if (res > dd) res--
        return PtgInt(res.toInt())
    }

    /**
     * LCM
     * Returns the least common multiple
     */
    internal fun calcLCM(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_NA)
        val numbers = PtgCalculator.getAllComponents(operands)
        // algorithm:
        // LCM(a, b)= (a*b)/GCD(a, b)
        var lcm: Long = 0
        try {
            val ops = arrayOfNulls<Ptg>(2)
            if ((lcm = PtgCalculator.getLongValue(numbers[0])) < 0)
                return PtgErr(PtgErr.ERROR_NUM)

            for (i in 1 until numbers.size) {
                val n2: Long
                if ((n2 = PtgCalculator.getLongValue(numbers[i])) < 0)
                    return PtgErr(PtgErr.ERROR_NUM)

                ops[0] = PtgNumber(lcm.toDouble())
                ops[1] = PtgNumber(n2.toDouble())
                val gcd = PtgCalculator.getLongValue(calcGCD(ops))
                lcm = lcm * n2 / gcd
            }
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

        return PtgNumber(lcm.toDouble())
    }

    /**
     * LN
     * Returns the natural logarithm of a number
     */
    internal fun calcLn(operands: Array<Ptg>): Ptg {
        if (operands.size > 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = Math.log(dd)
        return PtgNumber(res)
    }

    /**
     * LOG
     * Returns the logarithm of a number to a specified base
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcLog(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        if (dd.size > 2) return PtgCalculator.error
        val num1: Double
        val num2: Double
        if (dd.size == 1) {
            num1 = dd[0]
            num2 = 10.0
        } else {
            num1 = dd[0]
            num2 = dd[1]
        }
        val res = Math.log(num1) / Math.log(num2)
        return PtgNumber(res)
    }

    /**
     * LOG10
     * Returns the base-10 logarithm of a number
     */
    internal fun calcLog10(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        //	double res = Math.log(dd)/Math.log(10);

        val res = BigDecimal(Math.log10(dd))
        res.setScale(15, BigDecimal.ROUND_HALF_UP)
        return PtgNumber(res.toDouble())
    }
    /*
MDETERM
Returns the matrix determinant of an array

MINVERSE
Returns the matrix inverse of an array
*/

    /**
     * MMULT
     * Returns the matrix product of two arrays:
     * The number of columns in array1 must be the same as the number of rows in array2,
     * and both arrays must contain only numbers.
     * MMULT returns the #VALUE! error when:
     * Any cells are empty or contain text.
     * The number of columns in array1 is different from the number of rows in array2.
     */
    internal fun calcMMult(operands: Array<Ptg>): Ptg {
        if (operands.size != 2) return PtgCalculator.naError
        try {
            // error trap params:  must be numeric arrays with no empty spaces
            val a1 = PtgCalculator.getArray(operands[0])
            val a2 = PtgCalculator.getArray(operands[1])
            if (a1[0].size != a2.size)
                PtgCalculator.valueError
            var sum = 0.0
            for (i in 0 until a1[0].size) {
                sum += a1[0][i] * a2[i][0]
            }
            return PtgNumber(sum)
        } catch (e: Exception) {
            //		Logger.logErr("MMULT: error in operands " + e.toString());
        }

        return PtgCalculator.valueError
    }

    /**
     * MOD
     * Returns the remainder from division
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcMod(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        if (dd.size != 2) return PtgCalculator.error
        val num1 = dd[0]
        val num2 = dd[1]
        val res = num1 % num2
        return PtgNumber(res)
    }

    /**
     * MROUND
     * Returns a number rounded to the desired multiple
     */
    internal fun calcMRound(operands: Array<Ptg>): Ptg {
        if (operands.size != 2) return PtgErr(PtgErr.ERROR_NA)
        var m = 0.0
        var n = 0.0
        try {
            n = operands[0].doubleVal
            m = operands[1].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (n < 0 && m > 0 || n > 0 && m < 0) return PtgErr(PtgErr.ERROR_NUM)
        val result = Math.round(n / m) * m
        return PtgNumber(result)
    }


    /**
     * MULTINOMIAL
     * Returns the multinomial of a set of numbers
     */
    internal fun calcMultinomial(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_NA)
        val numbers = PtgCalculator.getAllComponents(operands)
        var sum: Long = 0
        var facts = 1.0
        for (i in numbers.indices) {
            val n = PtgCalculator.getLongValue(operands[i])
            if (n < 1) return PtgErr(PtgErr.ERROR_NUM)
            sum += n
            facts *= factorial(n).toDouble()
        }
        val result = factorial(sum) / facts
        return PtgNumber(result)
    }

    /**
     * ODD
     * Rounds a number up to the nearest odd integer
     */
    internal fun calcOdd(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        var resnum = Math.round(dd)
        // always round up!!
        if (resnum < dd) resnum++
        val remainder = (resnum % 2).toDouble()
        if (remainder == 0.0) {
            resnum++
        }
        return PtgInt(resnum.toInt())
    }

    /**
     * PI
     * Returns the value of Pi
     */
    internal fun calcPi(operands: Array<Ptg>): Ptg {
        val pi = Math.PI
        return PtgNumber(pi)
    }

    /**
     * POWER
     * Returns the result of a number raised to a power
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcPower(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        if (dd.size != 2) return PtgCalculator.error
        //	double num1 = dd[0];
        //	double num2 = dd[1];
        val num1 = BigDecimal(dd[0])
        num1.setScale(15, BigDecimal.ROUND_HALF_UP)
        val num2 = BigDecimal(dd[1])
        num2.setScale(15, BigDecimal.ROUND_HALF_UP)
        val res = BigDecimal(Math.pow(num1.toDouble(), num2.toDouble()))
        res.setScale(15, BigDecimal.ROUND_HALF_UP)
        return PtgNumber(res.toDouble())
    }

    /**
     * PRODUCT
     * Multiplies its arguments
     * NOTE:  we gotta deal with ranges/refs/numbers here
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcProduct(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//propagate error
        var result = dd[0]
        for (i in 1 until dd.size) {
            result = result * dd[i]
        }
        return PtgNumber(result)
    }


    /**
     * QUOTIENT
     * Returns the integer portion of a division
     */
    internal fun calcQuotient(operands: Array<Ptg>): Ptg {
        if (operands.size != 2) return PtgErr(PtgErr.ERROR_NA)
        var numerator = 0.0
        var denominator = 0.0
        try {
            numerator = operands[0].doubleVal
            denominator = operands[1].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        val result = (numerator / denominator).toLong()
        return PtgNumber(result.toDouble())
    }

    /**
     * RADIANS
     * Converts degrees to radians
     */
    internal fun calcRadians(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN())
        // Not a Num -- possibly PtgErr
            return PtgCalculator.error

        val res = Math.toRadians(dd)
        return PtgNumber(res)
    }


    /**
     * RAND
     * Returns a random number between 0 and 1
     */
    internal fun calcRand(operands: Array<Ptg>): Ptg {
        val dd = Math.random()
        return PtgNumber(dd)
    }

    /**
     * RANDBETWEEN
     * Returns a random number between the numbers you specify
     */
    internal fun calcRandBetween(operands: Array<Ptg>): Ptg {
        if (operands.size != 2) return PtgErr(PtgErr.ERROR_NA)
        var lower = 0
        var upper = 0
        try {
            lower = operands[0].intVal
            upper = operands[1].intVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        val r = Random()
        val result = (r.nextInt(upper - lower + 1) + lower).toDouble()
        return PtgNumber(result)
    }


    /**
     * ROMAN
     * Converts an Arabic numeral to Roman, as text
     * This one is a trip!
     */
    internal fun calcRoman(operands: Array<Ptg>): Ptg {
        val numbers = intArrayOf(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
        val letters = arrayOf("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
        var roman = ""
        if (operands.size == 1) {
            val dd = operands[0].doubleVal
            if (dd.isNaN())
            // Not a Num -- possibly PtgErr
                return PtgCalculator.error

            var i = dd.toInt()
            if (i < 0 || i > 3999) return PtgCalculator.error // can't write nums that high!
            for (z in numbers.indices) {
                while (i >= numbers[z]) {
                    roman += letters[z]
                    i -= numbers[z]
                }
            }
        }
        return PtgStr(roman)
    }

    /**
     * ROUND
     * Rounds a number to a specified number of digits
     * This one is kind of nasty.  3 cases of rounding.  If the rounding is a positive
     * integer this is the number of digits to round to.  If negative, round up past the
     * decimal, if 0, give an integer
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcRound(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        //if (dd.length != 2) return PtgCalculator.getError();
        // we need to handle arrays sent in, just use the first and last elements.
        val num = dd[0]
        var round = dd[dd.size - 1]
        var res = 0.0
        if (round == 0.0) {//return an int
            res = Math.round(num).toDouble()
        } else if (round > 0) { //round the decimal that number of spaces
            var tempnum = num * Math.pow(10.0, round)
            tempnum = Math.round(tempnum).toDouble()
            res = tempnum / Math.pow(10.0, round)
        } else { //round up the decimal that numbe of places
            round = round * -1
            var tempnum = num / Math.pow(10.0, round)
            tempnum = Math.round(tempnum).toDouble()
            res = tempnum * Math.pow(10.0, round)
        }
        return PtgNumber(res)
    }

    /**
     * ROUNDDOWN
     * Rounds a number down, toward zero.  Acts much like round above
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcRoundDown(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        //if (dd.length != 2) return PtgCalculator.getError();
        val num = dd[0]
        var round = dd[dd.size - 1]
        var res = 0.0
        if (round == 0.0) {//return an int
            res = Math.round(num).toDouble()
            if (res > num) res--
        } else if (round > 0) { //round the decimal that number of spaces
            val tempnum = num * Math.pow(10.0, round)
            var tempnum2 = Math.round(tempnum).toDouble()
            if (tempnum2 > tempnum) tempnum2--
            res = tempnum2 / Math.pow(10.0, round)
        } else { //round up the decimal that numbe of places
            round = round * -1
            val tempnum = num / Math.pow(10.0, round)
            var tempnum2 = Math.round(tempnum).toDouble()
            if (tempnum2 > tempnum) tempnum2--
            res = tempnum2 * Math.pow(10.0, round)
        }
        return PtgNumber(res)
    }

    /**
     * ROUNDUP
     * Rounds a number up, away from zero
     *
     * @throws CalculationException
     */
    @Throws(CalculationException::class)
    internal fun calcRoundUp(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        val dd = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
        //if (dd.length != 2) return PtgCalculator.getError();
        val num = dd[0]
        var round = dd[dd.size - 1]
        var res = 0.0
        if (round == 0.0) {//return an int
            res = Math.round(num).toDouble()
            if (res < num) res++
        } else if (round > 0) { //round the decimal that number of spaces
            val tempnum = num * Math.pow(10.0, round)
            var tempnum2 = Math.round(tempnum).toDouble()
            if (tempnum2 < tempnum) tempnum2++
            res = tempnum2 / Math.pow(10.0, round)
        } else { //round up the decimal that numbe of places
            round = round * -1
            val tempnum = num / Math.pow(10.0, round)
            var tempnum2 = Math.round(tempnum).toDouble()
            if (tempnum2 < tempnum) tempnum2++
            res = tempnum2 * Math.pow(10.0, round)
        }
        return PtgNumber(res)
    }
    /*
SERIESSUM
Returns the sum of a power series based on the formula
TODO: requires pack to run
*/

    /**
     * SIGN
     * Returns the sign of a number
     * return 1 if positive, -1 if negative, or 0 if 0;
     */
    internal fun calcSign(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        var res = 0
        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        if (dd == 0.0) res = 0
        if (dd > 0) res = 1
        if (dd < 0) res = -1
        return PtgInt(res)
    }

    /**
     * SIN
     * Returns the sine of the given angle
     */
    internal fun calcSin(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = Math.sin(dd)
        return PtgNumber(res)
    }

    /**
     * SINH
     * Returns the hyperbolic sine of a number
     */
    internal fun calcSinh(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = 0.5 * (Math.exp(dd) - Math.exp(-dd))
        return PtgNumber(res)
    }

    /**
     * SQRT
     * Returns a positive square root
     */
    internal fun calcSqrt(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = Math.sqrt(dd)
        return PtgNumber(res)
    }

    /**
     * SQRTPI
     * Returns the square root of (number * PI)
     */
    internal fun calcSqrtPi(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = Math.sqrt(dd * Math.PI)
        return PtgNumber(res)

    }
    /*
SUBTOTAL
Returns a subtotal in a list or database
*/

    /**
     * SUMIF
     * Adds the cells specified by a given criteria
     *
     *
     * You use the SUMIF function to sum the values in a range that meet criteria that you specify.
     * For example, suppose that in a column that contains numbers, you want to sum only the values that are larger than 5. You can use the following formula:
     *
     *
     * =SUMIF(B2:B25,">5")
     *
     *
     * you can apply the criteria to one range and sum the corresponding values in a different range.
     * For example, the formula =SUMIF(B2:B5, "John", C2:C5) sums only the values in the range C2:C5, where the corresponding cells in the range B2:B5 equal "John."
     *
     *
     * range  Required. The range of cells that you want evaluated by criteria. Cells in each range must be numbers or names, arrays, or references that contain numbers. Blank and text values are ignored.
     * criteria  Required. The criteria in the form of a number, expression, a cell reference, text, or a function that defines which cells will be added. For example, criteria can be expressed as 32, ">32", B5, 32, "32", "apples", or TODAY().
     * Important  Any text criteria or any criteria that includes logical or mathematical symbols must be enclosed in double quotation marks ("). If the criteria is numeric, double quotation marks are not required.
     * sum_range  Optional. The actual cells to add, if you want to add cells other than those specified in the range argument. If the sum_range argument is omitted, Excel adds the cells that are specified in the range argument
     * (the same cells to which the criteria is applied).
     *
     *
     * The sum_range argument does not have to be the same size and shape as the range argument.
     * The actual cells that are added are determined by using theupper leftmost cell in the sum_range argument as the beginning cell,
     * and then including cells that correspond in size and shape to the range argument.
     * For example:If range is And sum_range is Then the actual cells are
     * A1:A5 B1:B5 B1:B5
     * A1:A5 B1:B3 B1:B5
     * A1:B4 C1:D4 C1:D4
     * A1:B4 C1:C2 C1:D4
     *
     *
     * You can use the wildcard characters — the question mark (?) and asterisk (*) — as the criteria argument.
     * A question mark matches any single character; an asterisk matches any sequence of characters.
     * If you want to find an actual question mark or asterisk, type a tilde (~) preceding the character.
     */
    internal fun calcSumif(operands: Array<Ptg>): Ptg {
        try {
            val range = Calculator.getRange(operands[0])
            var sum_range: PtgArea? = null

            try {
                val criteria = operands[1]
                if (operands.size > 2) {    // see if has a sum_range; if not, source range is used for values as well as test
                    sum_range = Calculator.getRange(operands[2])
                }
                // OK at this point should have criteria, range and, if necessary, sum_range
                // algorithm:  for each entry that meets the criterium in range, get the cell;
                // if there is a sum_range, sum the values of those cells in the sum_range that correspond to the range

                // Parse the criteria string: can be a double, a comparison, a string with wildcards ...
                // more info: You can use the wildcard characters — the question mark (?) and asterisk (*) — as the criteria argument.
                // A question mark matches any single character; an asterisk matches any sequence of characters.
                // If you want to find an actual question mark or asterisk, type a tilde (~) preceding the character.
                var op = "="    // operator, default is =
                var sCriteria = criteria.string    // criteria in string form
                // strip operator, if any, and parse criteria
                val j = Calculator.splitOperator(sCriteria)
                if (j > 0) op = sCriteria.substring(0, j)    // extract operator, if any
                sCriteria = sCriteria.substring(j)
                sCriteria = Calculator.translateWildcardsInCriteria(sCriteria)

                // stores the cells that pass the criteria expression and therefore will be summed up
                val passesList = ArrayList()

                // test criteria for all cells in range, storing those cells (or sum_range cells)
                // that pass in passesList
                val cells = range!!.components
                var sumrangecells: Array<Ptg>? = null
                if (sum_range != null)
                    sumrangecells = sum_range.components
                for (i in cells!!.indices) {
                    var passes = false
                    try {
                        val v = cells[i].value
                        passes = Calculator.compareCellValue(v, sCriteria, op)
                    } catch (e: Exception) {    // don't report error
                        // Logger.logErr("MathFunctionCalculator.calcSumif:  error parsing " + e.toString());	// debugging only; take out when fully tested
                    }

                    if (passes) {
                        if (sumrangecells != null)
                            passesList.add(sumrangecells[i])
                        else
                            passesList.add(cells[i])
                    }
                }

                // At this point we have a collection of all the cells that pass (or their corresponding cell in sum_range);
                // Now we sum up the values of these cells and return
                var ret = 0.0
                for (i in passesList.indices) {
                    val cell = passesList.get(i) as Ptg
                    try {
                        ret += cell.doubleVal
                    } catch (e: Exception) {
                        Logger.logErr("MathFunctionCalculator.calcSumif:  error obtaining cell value: $e")    // debugging only; take out when fully tested
                        // keep going
                    }

                }
                return PtgNumber(ret)
            } catch (e: Exception) {
                Logger.logWarn("could not calculate SUMIF function: $e")
            }

        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_NULL)
    }

    /**
     * SUMIFS
     * Adds the cells in a range (range: Two or more cells on a sheet.
     * The cells in a range can be adjacent or nonadjacent.) that meet multiple criteria
     *
     *
     * sum_range  Required. One or more cells to sum, including numbers or names, ranges, or cell references (cell reference: The set of coordinates that a cell occupies on a worksheet. For example, the reference of the cell that appears at the intersection of column B and row 3 is B3.) that contain numbers. Blank and text values are ignored.
     * criteria_range1  Required. The first range in which to evaluate the associated criteria.
     * criteria1  Required. The criteria in the form of a number, expression, cell reference, or text that define which cells in the criteria_range1 argument will be added. For example, criteria can be expressed as 32, ">32", B4, "apples", or "32."
     * criteria_range2, criteria2, …  Optional. Additional ranges and their associated criteria. Up to 127 range/criteria pairs are allowed.
     *
     * @param operands
     * @return
     */
    internal fun calcSumIfS(operands: Array<Ptg>): Ptg {
        try {
            val sum_range = Calculator.getRange(operands[0])
            val sumrangecells = sum_range!!.components
            val ops = arrayOfNulls<String>((operands.size - 1) / 2)
            val criteria = arrayOfNulls<String>((operands.size - 1) / 2)
            val criteria_cells = arrayOfNulls<Array<Ptg>>((operands.size - 1) / 2)
            var j = 0
            run {
                var i = 1
                while (i + 1 < operands.size) {
                    //criteria range - parse and get comprising cells
                    val cr = Calculator.getRange(operands[i])
                    criteria_cells[j] = cr!!.components
                    // each criteria_range must contain the same number of rows and columns as the sum_range
                    if (criteria_cells[j].size != sumrangecells!!.size)
                        return PtgErr(PtgErr.ERROR_VALUE)
                    // criteria for comparison, including operator
                    criteria[j] = operands[i + 1].toString()
                    // strip operator, if any, and parse criteria
                    ops[j] = "="    // operator, default is =
                    val k = Calculator.splitOperator(criteria[j])
                    if (k > 0) ops[j] = criteria[j].substring(0, k)    // extract operator, if any
                    criteria[j] = criteria[j].substring(k)
                    criteria[j] = Calculator.translateWildcardsInCriteria(criteria[j])
                    j++
                    i += 2
                }
            }

            // test criteria for all cells in range, storing those corresponding sum_range cells
            // that pass in passesList
            // stores the cells that pass the criteria expression and therefore will be summed up
            val passesList = ArrayList()
            // for each set of criteria, test all cells in range and evaluate
            // NOTE:  this is an implicit AND evaluation
            for (i in sumrangecells!!.indices) {
                var passes = true
                for (k in criteria.indices) {
                    try {
                        val v = criteria_cells[k][i].value
                        passes = Calculator.compareCellValue(v, criteria[k], ops[k]) && passes
                        if (!passes)
                            break    // no need to continue
                    } catch (e: Exception) {    // don't report error
                    }

                }
                if (passes) {
                    passesList.add(sumrangecells[i])
                }
            }
            // At this point we have a collection of all the cells that pass (or their corresponding cell in sum_range);
            // Now we sum up the values of these cells and return
            var ret = 0.0
            for (i in passesList.indices) {
                val cell = passesList.get(i) as Ptg
                try {
                    ret += cell.doubleVal
                } catch (e: Exception) {
                    Logger.logErr("MathFunctionCalculator.calcSumif:  error obtaining cell value: $e")    // debugging only; take out when fully tested
                    // keep going
                }

            }

            return PtgNumber(ret)

        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_NULL)
    }

    /**
     * SUMPRODUCT
     * Returns the sum of the products of corresponding array components
     */
    internal fun calcSumproduct(operands: Array<Ptg>): Ptg {
        var res = 0.0
        var dim = 0    // all arrays must have same dimension see below
        val arrays = ArrayList()
        for (i in operands.indices) {
            if (operands[i] is PtgErr) return PtgErr(PtgErr.ERROR_NA)    // it's what excel does
            val a = operands[i].components
            if (a == null) {
                arrays.add(operands[i])
                if (dim == 0) {
                    dim = 1
                    continue
                } else {
                    return PtgErr(PtgErr.ERROR_VALUE)
                }
            }
            if (dim == 0)
                dim = a.size
            else if (dim != a.size)
                return PtgErr(PtgErr.ERROR_VALUE)
            arrays.add(a)
        }
        for (j in 0 until dim) {
            var d = 1.0
            for (i in arrays.indices) {
                val o = arrays.get(i)[j].value
                if (o is Double)
                    d = d * o.toDouble()
                else if (o is Int)
                    d = d * o.toInt()
                else if (o is Float)
                    d = d * o.toFloat()
                else
                    d = 0.0    // non-numeric values are treated as 0's
            }
            res += d
        }
        return PtgNumber(res)
    }

    /*
SUMSQ
Returns the sum of the squares of the arguments

SUMX2MY2
Returns the sum of the difference of squares of corresponding values in two arrays

SUMX2PY2
Returns the sum of the sum of squares of corresponding values in two arrays

SUMXMY2
Returns the sum of squares of differences of corresponding values in two arrays
*/

    /**
     * TAN
     * Returns the tangent of a number
     */
    internal fun calcTan(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = Math.tan(dd)
        return PtgNumber(res)
    }

    /**
     * TANH
     * Returns the hyperbolic tangent of a number
     */
    internal fun calcTanh(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var x = 0.0
        try {
            x = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        val A = Math.pow(Math.E, x)
        val B = Math.pow(Math.E, -x)
        val result = (A - B) / (A + B)
        return PtgNumber(result)
    }

    /**
     * TRUNC
     * Truncates a number to an integer
     */
    internal fun calcTrunc(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.naError
        var dd = 0.0
        try {
            dd = operands[0].doubleVal
        } catch (e: NumberFormatException) {
            return PtgCalculator.valueError
        }

        if (dd.isNaN()) return PtgCalculator.error // Not a Num -- possibly PtgErr
        val res = dd.toInt()
        return PtgInt(res)
    }

    /*
     *
     * These are some helper methods for the more brutal of the math functions
     */
    //helper for asinh
    private fun getAcosh(x: Double): Double {
        return Math.log(x + (1.0 + x) * Math.sqrt((x - 1.0) / (x + 1.0)))
    }

    // factorial helper
    fun factorial(n: Long): Long {
        var result: Long
        if (n <= 1) {
            result = 1 // 1! is 1
        } else {
            result = n
            val partial = factorial(n - 1)
            result = result * partial
        }// The recursive part
        return result
    }

    /*
     * Step factoral calculates a factorial of a number
     * the number of steps specified
     *
     */
    //Combin helper, steps factorials
    private fun stepFactorial(n: Long, numsteps: Int): Long {
        var n = n
        var numsteps = numsteps
        var result = n
        if (n < numsteps) return -1
        while (numsteps > 1) {
            val partial = n - 1
            result = partial * result
            n--
            numsteps--

        }
        return result
    }

    // double factorial helper
    private fun doubleFactorial(n: Long): Long {
        var result: Long
        if (n <= 1) {
            result = 1 // 1! is 1
        } else if (n == 2L) {
            result = 2
        } else {
            result = n
            val partial = factorial(n - 2)
            result = result * partial
        }// The recursive part
        return result
    }

}
