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
import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.math.BigDecimal
import java.util.ArrayList
import java.util.Vector


/*
    StatisticalCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/
class StatisticalCalculator {

    private fun R_Q_P01_check(p: Int, log_p: Boolean): Int {
        return if (log_p && p > 0 || !log_p && (p < 0 || p > 1)) 0 else 1
    }

    companion object {


        /**
         * AVERAGE
         * Returns the average (arithmetic mean) of the arguments.
         * Ignores non-numbers
         * This cannot recurse, due to averaging needs.
         *
         *
         * Usage@ AVERAGE(number1,number2, ...)
         * Returns@ PtgNumber
         */
        fun calcAverage(operands: Array<Ptg>): Ptg {
            val vect = Vector()

            for (i in operands.indices) {
                val pthings = operands[i].components // optimized -- do it once!! -jm
                if (pthings != null) {
                    for (z in pthings.indices) {
                        vect.add(pthings[z])
                    }
                } else {
                    val p = operands[i]
                    vect.add(p)
                }
            }
            var count = 0
            //        double total = 0;
            var bd = BigDecimal(0)
            for (i in vect.indices) {
                val p = vect.elementAt(i) as Ptg
                try {
                    if (p.isBlank) continue
                    val ov = p.value
                    if (ov != null) {
                        //                    total += Double.parseDouble(String.valueOf(ov));
                        bd = bd.add(BigDecimal(java.lang.Double.parseDouble(ov.toString())))
                        count++
                    }
                } catch (e: NumberFormatException) {
                }

            }
            bd = bd.setScale(15, java.math.RoundingMode.HALF_UP)
            val total = bd.toDouble()
            if (count == 0)
                return PtgErr(PtgErr.ERROR_DIV_ZERO)
            val result = total / count
            return PtgNumber(result)
        }

        /**
         * AVERAGEIF function
         * Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.
         *
         *
         * AVERAGEIF(range,criteria,average_range)
         *
         *
         * Range  is one or more cells to average, including numbers or names, arrays, or references that contain numbers.
         * Criteria  is the criteria in the form of a number, expression, cell reference, or text that defines which cells are averaged. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.
         * Average_range  is the actual set of cells to average. If omitted, range is used.
         * Average_range does not have to be the same size and shape as range. The actual cells that are averaged are determined by using the top, left cell in average_range as the beginning cell, and then including cells that correspond in size and shape to range.
         *
         * @param operands
         * @return
         */
        fun calcAverageIf(operands: Array<Ptg>): Ptg {
            if (operands.size < 2)
                return PtgErr(PtgErr.ERROR_DIV_ZERO)
            // range used to test criteria
            val range = operands[0].components
            // TODO: if range is blank  or a text value returns ERROR_DIV_ZERO
            var criteria = operands[1].string.trim { it <= ' ' }
            // Parse criteria into op  + criteria
            val i = Calculator.splitOperator(criteria)
            val op = criteria.substring(0, i)    // extract operand
            criteria = criteria.substring(i)
            criteria = Calculator.translateWildcardsInCriteria(criteria)
            // Average_range, if present, is used for return values if range passes criteria
            var average_range: Array<Ptg>? = null
            var varyRow = false
            if (operands.size > 2) {
                //The actual cells that are averaged are determined by using the top, left cell in average_range as the beginning cell,
                // and then including cells that correspond in size and shape to range.
                var rc: IntArray? = null
                average_range = arrayOfNulls(range.size)
                average_range[0] = operands[2].components[0]    // start with top left of average range
                var sheet = ""
                try {
                    rc = average_range[0].intLocation
                    if (range[0].intLocation[0] != range[range.size - 1].intLocation[0])
                    // determine if range is varied across row or column
                        varyRow = true
                    sheet = (average_range[0] as PtgRef).sheetName!! + "!"
                } catch (e: Exception) {
                }

                for (j in 1 until average_range.size) {
                    if (varyRow)
                        rc!![0]++
                    else
                        rc!![1]++
                    average_range[j] = PtgRef3d()
                    average_range[j].parentRec = range[0].parentRec
                    average_range[j].location = sheet + ExcelTools.formatLocation(rc)
                }
            }
            var nresults = 0
            var result = 0.0
            for (j in range.indices) {
                var `val`: Any? = range[j].value
                // TODO: TRUE and FALSE values are ignored
                // TODO: blank cells are treated as 0's
                if (Calculator.compareCellValue(`val`, criteria, op)) {
                    try {
                        if (average_range != null) {
                            `val` = average_range[j].value
                            if (`val` == null)
                            // if a cell is empty it's ignored --
                                continue
                        }
                        result += (`val` as Number).toDouble()
                    } catch (e: ClassCastException) {
                    }

                    nresults++
                }
            }

            // If no cells in the range meet the criteria, AVERAGEIF returns the #DIV/0! error value
            return if (nresults == 0) PtgErr(PtgErr.ERROR_DIV_ZERO) else PtgNumber(result / nresults)
            // otherwise, average
        }

        /**
         * AVERAGEIFS
         * Returns the average (arithmetic mean) of all cells that meet multiple criteria.
         * AVERAGEIFS(average_range,criteria_range1,criteria1,criteria_range2,criteria2…)
         * Average_range   is one or more cells to average, including numbers or names, arrays, or references that contain numbers.
         * Criteria_range1, criteria_range2, …   are 1 to 127 ranges in which to evaluate the associated criteria.
         * Criteria1, criteria2, …   are 1 to 127 criteria in the form of a number, expression, cell reference, or text that define which cells will be averaged. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.
         *
         * @param operands
         * @return
         */
        fun calcAverageIfS(operands: Array<Ptg>): Ptg {
            try {
                val average_range = Calculator.getRange(operands[0])
                val averagerangecells = average_range!!.components
                if (averagerangecells!!.size == 0)
                    return PtgErr(PtgErr.ERROR_DIV_ZERO)
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
                        if (criteria_cells[j].size != averagerangecells.size)
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

                // test criteria for all cells in range, storing those corresponding average_range cells
                // that pass in passesList
                // stores the cells that pass the criteria expression and therefore will be averaged
                val passesList = ArrayList()
                // for each set of criteria, test all cells in range and evaluate
                // NOTE:  this is an implicit AND evaluation
                for (i in averagerangecells.indices) {
                    var passes = true
                    for (k in criteria.indices) {
                        try {
                            val v = criteria_cells[k][i].value
                            // If cells in average_range cannot be translated into numbers, AVERAGEIFS returns the #DIV0! error value.
                            passes = Calculator.compareCellValue(v, criteria[k], ops[k]) && passes
                            if (!passes)
                                break    // no need to continue
                        } catch (e: Exception) {    // don't report error
                        }

                    }
                    if (passes) {
                        passesList.add(averagerangecells[i])
                    }
                }
                // If no cells in the range meet the criteria, AVERAGEIF returns the #DIV/0! error value
                if (passesList.size == 0)
                    return PtgErr(PtgErr.ERROR_DIV_ZERO)

                // At this point we have a collection of all the cells that pass (or their corresponding cell in sum_range);
                // Now we sum up the values of these cells and return
                var result = 0.0
                for (i in passesList.indices) {
                    val cell = passesList.get(i) as Ptg
                    try {
                        result += cell.doubleVal
                    } catch (e: Exception) {
                        Logger.logErr("MathFunctionCalculator.calcAverageIfS:  error obtaining cell value: $e")    // debugging only; take out when fully tested
                        // keep going
                    }

                }

                // otherwise, average
                return PtgNumber(result / passesList.size)

            } catch (e: Exception) {
            }

            return PtgErr(PtgErr.ERROR_NULL)
        }

        /**
         * AVEDEV(number1,number2, ...)
         * Number1, number2, ...   are 1 to 30 arguments for which you want
         * the average of the absolute deviations. You can also use a
         * single array or a reference to an array instead of arguments
         * separated by commas.
         *
         *
         * The arguments must be either numbers or names,
         * arrays, or references that contain numbers.
         *
         *
         * If an array or reference argument contains text,
         * logical values, or empty cells, those values are ignored;
         * however, cells with the value zero are included.
         */
        fun calcAveDev(operands: Array<Ptg>): Ptg {
            if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
            // Get the average for the mean
            val av = StatisticalCalculator.calcAverage(operands) as PtgNumber
            var average = -0.001
            try {
                val dd = Double(av.value.toString())
                average = dd
            } catch (e: NumberFormatException) {
            }

            if (average == -0.001) return PtgCalculator.error

            // work out the total deviation
            var total = 0.0
            var count = 0
            var d: Double?
            val alloperands = PtgCalculator.getAllComponents(operands)
            for (i in alloperands.indices) {
                val resPtg = alloperands[i]
                try {  // some fields may be text, so handle gracefully
                    if (resPtg.value != null) {
                        d = Double(resPtg.value.toString())
                        var dub = d.toDouble()
                        dub = average - dub
                        dub = Math.abs(dub)
                        total += dub
                        count++
                    }
                } catch (e: NumberFormatException) {
                }

            }
            // work out the mean deviation
            val mean = total / count
            return PtgNumber(mean)
        }

        /**
         * AVERAGEA
         * Returns the average of its arguments, including numbers, text,
         * and logical values
         *
         *
         * The arguments must be numbers, names, arrays, or references.
         * Array or reference arguments that contain text evaluate as 0 (zero).
         * Empty text ("") evaluates as 0 (zero).
         * Arguments that contain TRUE evaluate as 1;
         * arguments that contain FALSE evaluate as 0 (zero).
         */
        fun calcAverageA(operands: Array<Ptg>): Ptg {
            val alloperands = PtgCalculator.getAllComponents(operands)
            var total = 0.0
            for (i in alloperands.indices) {
                val p = alloperands[i]
                try {
                    val ov = p.value
                    if (ov != null) {
                        if (ov.toString() === "true") {
                            total++
                        } else {
                            total += java.lang.Double.parseDouble(ov.toString())
                        }
                    }
                } catch (e: NumberFormatException) {
                }

            }
            val result = total / alloperands.size
            return PtgNumber(result)
        }
        /*
BETADIST
 Returns the cumulative beta probability density function

BETAINV
 Returns the inverse of the cumulative beta probability density function

BINOMDIST
 Returns the individual term binomial distribution probability

CHIDIST
 Returns the one-tailed probability of the chi-squared distribution

CHIINV
 Returns the inverse of the one-tailed probability of the chi-squared distribution

CHITEST
 Returns the test for independence

CONFIDENCE
 Returns the confidence interval for a population mean
 */

        /**
         * CORREL
         * Returns the correlation coefficient between two data sets
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcCorrel(operands: Array<Ptg>): Ptg {
            // get the covariance
            var pnum = calcCovar(operands) as PtgNumber
            val covar = pnum.`val`
            val xPtg = arrayOfNulls<Ptg>(1)
            xPtg[0] = operands[0]
            val yPtg = arrayOfNulls<Ptg>(1)
            yPtg[0] = operands[1]
            pnum = calcAverage(xPtg) as PtgNumber
            val xMean = pnum.`val`
            pnum = calcAverage(yPtg) as PtgNumber
            val yMean = pnum.`val`
            val xVals = PtgCalculator.getDoubleValueArray(xPtg)
            val yVals = PtgCalculator.getDoubleValueArray(yPtg)
            if (xVals == null || yVals == null)
                return PtgErr(PtgErr.ERROR_NA)//20090130 KSC: propagate error
            var xstat = 0.0
            for (i in xVals.indices) {
                xstat += Math.pow(xVals[i] - xMean, 2.0)
            }
            xstat = xstat / xVals.size
            xstat = Math.sqrt(xstat)
            var ystat = 0.0
            for (i in yVals.indices) {
                ystat += Math.pow(yVals[i] - yMean, 2.0)
            }
            ystat = ystat / yVals.size
            ystat = Math.sqrt(ystat)
            val retval = covar / (ystat * xstat)
            return PtgNumber(retval)
        }

        /**
         * COUNT
         * Counts how many numbers are in the list of arguments
         */
        /**
         * Counts the number of cells that contain numbers
         * and or dates.
         * Use COUNT to get the number of entries in a number
         * field in a range or array of numbers.
         *
         *
         * Usage@ COUNT(A1:A5,A9)
         * Return@ PtgInt
         * TODO: implement counting of dates!
         */
        fun calcCount(operands: Array<Ptg>): Ptg {
            var count = 0
            for (i in operands.indices) {
                val pref = operands[i].components // optimized -- do it once!! -jm
                if (pref != null) { // it is some sort of range
                    for (z in pref.indices) {
                        val o = pref[z].value
                        if (o != null) {
                            try {
                                val n = Double(o.toString())
                                count++
                            } catch (e: NumberFormatException) {
                            }

                        }
                    }
                } else {  // it's a single ptgref
                    val o = operands[i].value
                    if (o != null) {
                        try {
                            val n = Double(o.toString())
                            count++
                        } catch (e: NumberFormatException) {
                        }

                    }
                }
            }
            return PtgInt(count)

        }

        /**
         * COUNTA
         * Counts the number of non-blank cells within a range
         */
        fun calcCountA(operands: Array<Ptg>): Ptg {
            val allops = PtgCalculator.getAllComponents(operands)
            var count = 0
            for (i in allops.indices) {
                /* 20081120 KSC: blnaks are handled differently as Excel counts blank cells as 0's
		   Object o = allops[i].getValue();
		   if (o != null) count++;
		   */
                if (!allops[i].isBlank)
                    count++
            }
            return PtgInt(count)
        }

        /**
         * COUNTBLANK
         * Counts the number of blank cells within a range
         */
        fun calcCountBlank(operands: Array<Ptg>): Ptg {
            val allops = PtgCalculator.getAllComponents(operands)
            var count = 0
            for (i in allops.indices) {
                if (allops[i].isBlank)
                // 20081112 KSC: was Object o = getValue(); if (o==null) count++;
                    count++
            }
            return PtgInt(count)
        }


        /**
         * COUNTIF
         * Counts the number of non-blank cells within a range which meet the given criteria
         * BRUTAL!!
         */
        @Throws(FunctionNotSupportedException::class)
        fun calcCountif(operands: Array<Ptg>): Ptg {
            if (operands.size != 2) return PtgCalculator.error
            val matchStr = operands[1].value.toString()
            var donumber = true
            var matchDub = 0.0
            try {    // this method matches strings or numbers, here is where we differentiate
                val d = Double(matchStr)
                matchDub = d
            } catch (e: Exception) {
                donumber = false
            }

            var count = 0.0
            val pref = operands[0].components // optimize by doing it one time!!! this thing gets slow....-jm
            if (pref != null) { // it is some sort of range
                for (z in pref.indices) {
                    val o = pref[z].value
                    if (o != null) {
                        val match2 = o.toString()
                        if (donumber) {
                            try {
                                val d = Double(match2)
                                if (matchDub == d) count++
                            } catch (e: NumberFormatException) {
                            }

                        } else {
                            if (matchStr.equals(match2, ignoreCase = true)) count++
                        }
                    }
                }
            } else {  // it's a single ptgref
                val o = operands[0].value
                if (o != null) {
                    if (o != null) {
                        val match2 = o.toString()
                        if (donumber) {
                            try {
                                val d = Double(match2)
                                if (matchDub == d) count++
                            } catch (e: NumberFormatException) {
                            }

                        } else {
                            if (matchStr.equals(match2, ignoreCase = true)) count++
                        }
                    }

                }
            }
            return PtgNumber(count)
        }

        /**
         * COUNTIFS
         * criteria_range1  Required. The first range in which to evaluate the associated criteria.
         * criteria1  Required. The criteria in the form of a number, expression, cell reference, or text that define which cells will be counted. For example, criteria can be expressed as 32, ">32", B4, "apples", or "32".
         * criteria_range2, criteria2, ...  Optional. Additional ranges and their associated criteria. Up to 127 range/criteria pairs are allowed.
         *
         * @param operands
         * @return
         */
        fun calcCountIfS(operands: Array<Ptg>): Ptg {
            try {
                val ops = arrayOfNulls<String>(operands.size / 2)
                val criteria = arrayOfNulls<String>(operands.size / 2)
                val criteria_cells = arrayOfNulls<Array<Ptg>>(operands.size / 2)
                run {
                    var i = 0
                    while (i + 1 < operands.size) {
                        //criteria range - parse and get comprising cells
                        val cr = Calculator.getRange(operands[i])
                        criteria_cells[i / 2] = cr!!.components
                        // each criteria_range must contain the same number of rows and columns as the criteriarange
                        if (i > 0 && criteria_cells[i / 2].size != criteria_cells[0].size)
                            return PtgErr(PtgErr.ERROR_VALUE)
                        // criteria for comparison, including operator
                        criteria[i / 2] = operands[i + 1].toString()
                        // strip operator, if any, and parse criteria
                        ops[i / 2] = "="    // operator, default is =
                        val k = Calculator.splitOperator(criteria[i / 2])
                        if (k > 0) ops[i / 2] = criteria[i / 2].substring(0, k)    // extract operator, if any
                        criteria[i / 2] = criteria[i / 2].substring(k)
                        criteria[i / 2] = Calculator.translateWildcardsInCriteria(criteria[i / 2])
                        i += 2
                    }
                }

                // test criteria for all cells in range, counting each cell that passes
                // for each set of criteria, test all cells in range and evaluate
                var count = 0
                for (i in 0 until criteria_cells[0].size) {
                    var passes = true
                    for (k in criteria.indices) {
                        try {
                            val v = criteria_cells[k][i].value
                            //  the criteria argument is a reference to an empty cell, the COUNTIFS function treats the empty cell as a 0 value.
                            passes = Calculator.compareCellValue(v, criteria[k], ops[k]) && passes
                            if (!passes)
                                break    // no need to continue
                        } catch (e: Exception) {    // don't report error
                        }

                    }
                    if (passes)
                        count++
                }

                return PtgNumber(count.toDouble())

            } catch (e: Exception) {
            }

            return PtgErr(PtgErr.ERROR_NULL)
        }

        /**
         * COVAR
         * Returns covariance, the average of the products of paired deviations
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcCovar(operands: Array<Ptg>): Ptg {
            val xMeanPtg = arrayOfNulls<Ptg>(1)
            xMeanPtg[0] = operands[0]
            val yMeanPtg = arrayOfNulls<Ptg>(1)
            yMeanPtg[0] = operands[1]
            var pnum = calcAverage(xMeanPtg) as PtgNumber
            val xMean = pnum.`val`
            pnum = calcAverage(yMeanPtg) as PtgNumber
            val yMean = pnum.`val`
            val xVals = PtgCalculator.getDoubleValueArray(xMeanPtg)
            val yVals = PtgCalculator.getDoubleValueArray(yMeanPtg)
            if (xVals == null || yVals == null)
                return PtgErr(PtgErr.ERROR_NA)//propagate error
            var xyMean = 0.0
            if (xVals.size == yVals.size) {
                var addvals = 0
                for (i in xVals.indices) {
                    addvals += (xVals[i] * yVals[i]).toInt()
                }
                xyMean = (addvals / xVals.size).toDouble()
            } else {
                return PtgErr(PtgErr.ERROR_NA)
            }

            val retval = xyMean - xMean * yMean
            return PtgNumber(retval)
        }
        /*
CRITBINOM
 Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value

DEVSQ
 Returns the sum of squares of deviations

EXPONDIST
 Returns the exponential distribution

FDIST
 Returns the F probability distribution

FINV
 Returns the inverse of the F probability distribution

FISHER
 Returns the Fisher transformation

FISHERINV
 Returns the inverse of the Fisher transformation
 */

        /**
         * FORECAST
         * Returns a value along a linear trend
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcForecast(operands: Array<Ptg>): Ptg {
            if (operands.size != 2) return PtgErr(PtgErr.ERROR_VALUE)
            val p = arrayOfNulls<Ptg>(2)
            p[0] = operands[0]
            p[1] = operands[1]
            val icept = calcIntercept(p) as PtgNumber
            val intercept = icept.`val`
            val slp = calcSlope(p) as PtgNumber
            val slope = slp.`val`
            val px = operands[0]
            val knownX = Double(px.value.toString())
            val retval = slope * knownX + intercept
            return PtgNumber(retval)
        }

        /**
         * FREQUENCY
         * Returns a frequency distribution as a vertical array
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcFrequency(operands: Array<Ptg>): Ptg {
            val firstArr = PtgCalculator.getAllComponents(operands[0])
            val secondArr = PtgCalculator.getAllComponents(operands[1])
            val t = CompatibleVector()
            for (i in secondArr.indices) {
                val p = secondArr[i]
                try {
                    val d = Double(p.value.toString())
                    t.addOrderedDouble(d)
                } catch (e: NumberFormatException) {
                }

            }
            val binsArr = arrayOfNulls<Double>(t.size)
            val dataArr: DoubleArray?
            dataArr = PtgCalculator.getDoubleValueArray(firstArr)
            if (dataArr == null) return PtgErr(PtgErr.ERROR_NA)//20090130 KSC: propagate error
            t.toTypedArray()
            val retvals = IntArray(secondArr.size + 1)
            for (i in dataArr.indices) {
                var x = 0
                while (x < binsArr.size) {
                    if (dataArr[i] <= binsArr[x].toDouble()) {
                        retvals[x]++
                        x = binsArr.size
                    } else if (dataArr[i] > binsArr[binsArr.size - 1].toDouble()) {
                        retvals[binsArr.size]++
                        x = binsArr.size
                    }
                    x++
                }
            }
            // keep the original locations, so we can put the end result array in the correct order.
            // not used!	double[] originalLocs = PtgCalculator.getDoubleValueArray(secondArr);
            var ret = "{"
            for (i in retvals.indices) {
                ret += retvals[i].toString() + ","
            }
            ret = ret.substring(0, ret.length - 1) // get rid of final comma
            ret += "}"

            val returnArr = PtgArray()
            returnArr.setVal(ret)
            return returnArr
        }
        /*
FTEST
 Returns the result of an F-test

GAMMADIST
 Returns the gamma distribution

GAMMAINV
 Returns the inverse of the gamma cumulative distribution

GAMMALN
 Returns the natural logarithm of the gamma function, G(x)

GEOMEAN
 Returns the geometric mean

GROWTH
 Returns values along an exponential trend

HARMEAN
 Returns the harmonic mean

HYPGEOMDIST
 Returns the hypergeometric distribution
 */

        /**
         * INTERCEPT
         * Returns the intercept of the linear regression line
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcIntercept(operands: Array<Ptg>): Ptg {
            val yvals: DoubleArray?
            yvals = PtgCalculator.getDoubleValueArray(operands[0])

            val xvals = PtgCalculator.getDoubleValueArray(operands[1])
            if (xvals == null || yvals == null) return PtgErr(PtgErr.ERROR_NA)//20090130 KSC: propagate error
            var sumXVals = 0.0
            for (i in xvals.indices) {
                sumXVals += xvals[i]
            }
            var sumYVals = 0.0
            for (i in yvals.indices) {
                sumYVals += yvals[i]
            }
            var sumXYVals = 0.0
            for (i in yvals.indices) {
                sumXYVals += xvals[i] * yvals[i]
            }
            var sqrXVals = 0.0
            for (i in xvals.indices) {
                sqrXVals += xvals[i] * xvals[i]
            }
            val toparg = sumXVals * sumXYVals - sumYVals * sqrXVals
            val bottomarg = sumXVals * sumXVals - sqrXVals * xvals.size
            val res = toparg / bottomarg
            return PtgNumber(res)
        }
        /*
KURT
 Returns the kurtosis of a data set
 */

        /**
         * LARGE
         *
         *
         * Returns the k-th largest value in a data set. You can use this function to select a value based on its relative standing.
         * For example, you can use LARGE to return the highest, runner-up, or third-place score.
         *
         *
         * LARGE(array,k)
         *
         *
         * Array   is the array or range of data for which you want to determine the k-th largest value.
         * K   is the position (from the largest) in the array or cell range of data to return.
         *
         *
         * If array is empty, LARGE returns the #NUM! error value.
         * If k ≤ 0 or if k is greater than the number of data points, LARGE returns the #NUM! error value.
         * If n is the number of data points in a range, then LARGE(array,1) returns the largest value, and LARGE(array,n) returns the smallest value.
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcLarge(operands: Array<Ptg>): Ptg {
            if (operands.size != 2) return PtgErr(PtgErr.ERROR_VALUE)
            val rng = operands[0]
            val array = PtgCalculator.getAllComponents(rng)
            if (array.size == 0) return PtgErr(PtgErr.ERROR_NUM)
            val k = PtgCalculator.getDoubleValueArray(operands[1])!![0].toInt()
            if (k <= 0 || k > array.size)
                return PtgErr(PtgErr.ERROR_NUM)

            val sortedValues = CompatibleVector()
            for (i in array.indices) {
                val p = array[i]
                try {
                    val d = Double(p.value.toString())
                    sortedValues.addOrderedDouble(d)
                } catch (e: NumberFormatException) {
                }

            }
            // reverse array
            val dubRefs = arrayOfNulls<Double>(sortedValues.size)
            for (i in dubRefs.indices) {
                dubRefs[i] = sortedValues.last() as Double
                sortedValues.removeAt(sortedValues.size - 1)
            }

            return PtgNumber(dubRefs[k - 1].toDouble())
            /*

	 try {
		 Ptg[] parray= PtgCalculator.getAllComponents(operands[0]);
		 Object[] array= new Object[parray.length];
		 for (int i= 0; i < array.length; i++) {
			 array[i]= parray[i].getValue();
			 if (array[i] instanceof Integer) {	// convert all to double if possible for sort below (cannot have mixed array for Arrays.sort)
				 try {
					 array[i]= new Double(((Integer)array[i]).intValue());
				 } catch (Exception e) {
				 }
		 	}
		 }
		 // now sort
		 java.util.Arrays.sort(array);
		 int position= (int) PtgCalculator.getLongValue(operands[1]);
		 // now return the nth item in the sorted (asc) array
		 if (position >=0 && position <=array.length) {
			 Object ret= array[array.length-position];
			 if (ret instanceof Double)
				 return new PtgNumber(((Double)ret).doubleValue());
			 else if (ret instanceof Boolean)
				 return new PtgBool(((Boolean)ret).booleanValue());
			 else if (ret instanceof String)
				 return new PtgStr((String)ret);
		 }
	 } catch (Exception e) {

	 }
	 return new PtgErr(PtgErr.ERROR_NUM);
*/
        }


        /**
         * LINEST
         * Returns the parameters of a linear trend
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcLineSt(operands: Array<Ptg>): Ptg {
            val ys = PtgCalculator.getDoubleValueArray(operands[0]) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
            val xs: DoubleArray?
            if (operands.size == 1 || operands[1] is PtgMissArg) {
                // create a 1,2,3 array
                xs = DoubleArray(ys.size)
                for (i in ys.indices)
                    xs[i] = i.toDouble()
            } else {
                xs = PtgCalculator.getDoubleValueArray(operands[1])
                if (xs == null) return PtgErr(PtgErr.ERROR_NA)//20090130 KSC: propagate error
            }

            var getIntercept = false
            if (operands.size > 2) {
                if (operands[2] !is PtgMissArg)
                    getIntercept = PtgCalculator.getBooleanValue(operands[2])
            }
            var statistics = false
            if (operands.size > 3) {
                if (operands[3] !is PtgMissArg)
                    statistics = PtgCalculator.getBooleanValue(operands[3])
            }

            val ps = calcSlope(operands)
            if (ps is PtgErr) return ps
            val Pslope = calcSlope(operands) as PtgNumber
            val slope = Pslope.`val`  // a1 val
            val Pintercept = calcIntercept(operands) as PtgNumber
            val intercept = Pintercept.`val` // b1 val

            if (operands.size > 3 && (operands[3] is PtgBool || operands[3] is PtgInt)) {
                val b = PtgCalculator.getBooleanValue(operands[3])
                if (!b) {
                    var retstr = "{$slope,$intercept},"
                    retstr += "{$slope,$intercept},"
                    retstr += "{$slope,$intercept},"
                    retstr += "{$slope,$intercept},"
                    retstr += "{$slope,$intercept}"
                    val para = PtgArray()
                    para.setVal(retstr)
                    return para
                }
            }
            var p = arrayOfNulls<Ptg>(1)

            // figure out the stdev of the slope
            val Psteyx = calcSteyx(operands) as PtgNumber
            val steyx = Psteyx.`val` // b3 val

            // calc the y error percentage
            var yError = steyx * steyx
            p[0] = operands[1]
            val vp = calcVarp(p) as PtgNumber
            val Sxx = vp.`val` * ys.size
            yError = yError / Sxx
            yError = Math.sqrt(yError) // A2 val

            // calculate degrees of freedom
            val degFreedom = ys.size - 2 // b4 val

            // calculate standard error of intercept
            var sumXsquared = 0.0
            var sumSquaredX = 0.0
            var sumXYsquared = 0.0
            for (i in xs.indices) {
                sumSquaredX += xs[i] * xs[i]
                sumXsquared += xs[i]
                sumXYsquared += xs[i] * ys[i]
            }
            sumXsquared *= sumXsquared
            sumXYsquared *= sumXYsquared
            var interceptError = 1 / (xs.size - sumXsquared / sumSquaredX)
            interceptError = Math.sqrt(interceptError)
            interceptError *= steyx //b2val

            // calculate residual SS
            // first create array of predicted values for the linear array
            val predicted = DoubleArray(xs.size)
            var residualSS = 0.0 // b5value
            for (i in xs.indices) {
                predicted[i] = intercept + xs[i] * slope
                val d = predicted[i] - ys[i]
                residualSS += d * d
            }

            // calculate regression SS
            p[0] = operands[0]
            var pnum = calcAverage(p) as PtgNumber
            val average = pnum.`val`
            var regressionSS = 0.0
            for (i in xs.indices) {
                val d = predicted[i] - average
                regressionSS += d * d//A5 value
            }
            p = arrayOfNulls(2)
            p[0] = operands[0]
            p[1] = operands[1]
            pnum = calcRsq(p) as PtgNumber
            val r2 = pnum.`val`    // A3

            // calculate the F value
            val F = regressionSS / 1 / (residualSS / degFreedom) // A4

            // construct the string for creating ptgarray
            var retstr = "{$slope,$intercept},"
            retstr += "{$yError,$interceptError},"
            retstr += "{$r2,$steyx},"
            retstr += "{$F,$degFreedom},"
            retstr += "{$regressionSS,$residualSS}"

            val parr = PtgArray()
            parr.setVal(retstr)

            return parr
        }
        /*
LOGEST
 Returns the parameters of an exponential trend

LOGINV
 Returns the inverse of the lognormal distribution

LOGNORMDIST
 Returns the cumulative lognormal distribution
 */


        /**
         * MAX
         * Returns the largest value in a set of values.
         * Ignores non-number fields
         * Recursively calls for ranges.
         *
         *
         * Usage@ MAX(number1,number2, ...)
         * returns@ PtgNumber
         */
        //untested
        fun calcMax(operands: Array<Ptg>): Ptg {
            var result = java.lang.Double.MIN_VALUE        // 20090129 KSC -1;
            var d: Double? = null
            for (i in operands.indices) {
                val pthings = operands[i].components // optimized -- do it once!! -jm
                if (pthings != null) {
                    val resPtg = StatisticalCalculator.calcMax(pthings)
                    try {  // some fields may be text, so handle gracefully
                        if (resPtg.value != null) d = Double(resPtg.value.toString())
                        if (d!!.toDouble() > result) {
                            result = d.toDouble()
                        }
                    } catch (e: NumberFormatException) {
                    }

                } else {
                    val p = operands[i]
                    try {
                        val ov = p.value
                        if (ov != null)
                            d = Double(ov.toString())
                        if (d!!.toDouble() > result) {
                            result = d.toDouble()
                        }
                    } catch (e: NumberFormatException) {
                    } catch (e: NullPointerException) {
                    }

                }
            }
            if (result == java.lang.Double.MIN_VALUE)
            // 20090129 KSC:
                result = 0.0        //appears to be default in error situations
            return PtgNumber(result)
        }

        /**
         * MAXA
         * Returns the maximum value in a list of arguments, including numbers, text, and logical values
         *
         *
         * Arguments can be the following: numbers; names, arrays, or references that contain numbers; text representations of numbers; or logical values, such as TRUE and FALSE, in a reference.
         * Logical values and text representations of numbers that you type directly into the list of arguments are counted.
         */
        fun calcMaxA(operands: Array<Ptg>): Ptg {
            val alloperands = PtgCalculator.getAllComponents(operands)
            if (alloperands.size == 0)
                return PtgNumber(0.0)
            var max = java.lang.Double.MIN_VALUE
            for (i in alloperands.indices) {
                val o = alloperands[i].value
                try {
                    var d = java.lang.Double.MIN_VALUE
                    if (o is Number)
                        d = o.toDouble()
                    else if (o is Boolean)
                        d = (if (o.booleanValue()) 1 else 0).toDouble()
                    else
                        d = Double(o.toString())
                    max = Math.max(max, d)
                } catch (e: NumberFormatException) {
                    // Arguments that are error values or text that cannot be translated into numbers cause errors.
                    return PtgErr(PtgErr.ERROR_VALUE)
                }

            }
            return PtgNumber(max)
        }

        /**
         * MEDIAN
         * Returns the median of the given numbers
         */
        fun calcMedian(operands: Array<Ptg>): Ptg {
            if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
            val alloperands = PtgCalculator.getAllComponents(operands)
            val t = CompatibleVector()
            var retval = 0.0
            for (i in alloperands.indices) {
                val p = alloperands[i]
                try {
                    val d = Double(p.value.toString())
                    t.addOrderedDouble(d)
                } catch (e: NumberFormatException) {
                }

            }

            try {
                val dub = arrayOfNulls<Double>(t.size)
                t.toTypedArray()
                val dd = t.size.toDouble() % 2
                if (t.size.toDouble() % 2 == 0.0) {
                    val firstValLoc = t.size / 2 - 1
                    val lastValLoc = firstValLoc + 1
                    val firstVal = dub[firstValLoc].toDouble()
                    val lastVal = dub[lastValLoc].toDouble()
                    retval = (firstVal + lastVal) / 2
                } else {
                    // it's odd
                    val firstValLoc = (t.size - 1) / 2
                    val firstVal = dub[firstValLoc].toDouble()
                    retval = firstVal
                }
                return PtgNumber(retval)
            } catch (e: ArrayIndexOutOfBoundsException) {    // 20090701 KSC: catch exception
                return PtgErr(PtgErr.ERROR_VALUE)
            }

        }

        /**
         * MIN
         * Returns the smallest number in a set of values.
         * Ignores non-number fields.  Note that it also recursivly calls itself
         * for things like PtgRange.
         *
         *
         * Usage@ MIN(number1,number2, ...)
         * returns PtgNumber
         */
        fun calcMin(operands: Array<Ptg>): Ptg {
            var result = java.lang.Double.MAX_VALUE
            var d: Double? = null
            for (i in operands.indices) {
                val pthings = operands[i].components // optimized -- do it once!! -jm
                if (pthings != null) {
                    val resPtg = StatisticalCalculator.calcMin(pthings)
                    try {  // some fields may be text, so handle gracefully
                        if (resPtg is PtgErr) return resPtg    // 20090205 KSC: propagate error
                        if (resPtg.value != null) {
                            d = Double(resPtg.value.toString())
                            // 20090129 KSC; if (d.doubleValue() < result || result == -1){result = d.doubleValue();} // 20070215 KSC: only access d if not null!
                            if (d.toDouble() < result) {
                                result = d.toDouble()
                            } // 20070215 KSC: only access d if not null!
                        }
                    } catch (e: NumberFormatException) {
                    } catch (e: NullPointerException) {
                    }
                    // 20070209 KSC
                } else {
                    val p = operands[i]
                    try {
                        val ov = p.value
                        if (ov != null) {
                            if (ov.toString() == PtgErr(PtgErr.ERROR_NA).toString())
                            // 20090205 KSC: propagate error value
                                return PtgErr(PtgErr.ERROR_NA)
                            d = Double(ov.toString())
                            // 20090129 KSC; result is defaulted to max
                            if (d.toDouble() < result) {
                                result = d.toDouble()
                            } // 20070215 KSC: only access d if not null!
                        }
                    } catch (e: NumberFormatException) {
                    } catch (e: NullPointerException) {
                    }

                }
            }
            if (result == java.lang.Double.MAX_VALUE)
            // 20090129 KSC:
                result = 0.0        //appears to be default in error situations
            return PtgNumber(result)
            // return pnum;
        }

        /**
         * MINA
         * Returns the smallest value in a list of arguments, including numbers, text, and logical values
         *
         *
         * Arguments can be the following: numbers; names, arrays, or references that contain numbers; text representations of numbers; or logical values, such as TRUE and FALSE, in a reference.
         * Logical values and text representations of numbers that you type directly into the list of arguments are counted.
         */
        fun calcMinA(operands: Array<Ptg>): Ptg {
            val alloperands = PtgCalculator.getAllComponents(operands)
            if (alloperands.size == 0)
                return PtgNumber(0.0)
            var min = java.lang.Double.MAX_VALUE
            for (i in alloperands.indices) {
                val o = alloperands[i].value
                try {
                    var d = java.lang.Double.MAX_VALUE
                    if (o is Number)
                        d = o.toDouble()
                    else if (o is Boolean)
                        d = (if (o.booleanValue()) 1 else 0).toDouble()
                    else
                        d = Double(o.toString())
                    min = Math.min(min, d)
                } catch (e: NumberFormatException) {
                    // Arguments that are error values or text that cannot be translated into numbers cause errors.
                    return PtgErr(PtgErr.ERROR_VALUE)
                }

            }
            return PtgNumber(min)
        }

        /**
         * MODE
         * Returns the most common value in a data set
         */
        fun calcMode(operands: Array<Ptg>): Ptg {
            val alloperands = PtgCalculator.getAllComponents(operands)
            val vals = Vector()
            val occurences = Vector()
            var retval = 0.0
            for (i in alloperands.indices) {
                val p = alloperands[i]
                try {
                    val d = Double(p.value.toString())
                    if (vals.contains(d)) {
                        val loc = vals.indexOf(d)
                        val nums = occurences.get(loc) as Double
                        val newnum = nums + 1
                        occurences.setElementAt(newnum, loc)
                    } else {
                        vals.add(d)
                        occurences.add(1)
                    }
                } catch (e: NumberFormatException) {
                }

            }
            var biggest = 0.0
            val numvalues = 0.0
            for (i in vals.indices) {
                val size = occurences.elementAt(i) as Double
                if (size > biggest) {
                    biggest = size
                    val newhigh = vals.elementAt(i) as Double
                    retval = newhigh
                }
            }
            return PtgNumber(retval)
        }

        /*
NEGBINOMDIST
 Returns the negative binomial distribution
 */


        /**
         * NORMDIST
         * Returns the normal cumulative distribution
         * NORMDIST(x,mean,standard_dev,cumulative)
         * X     is the value for which you want the distribution.
         * Mean     is the arithmetic mean of the distribution.
         * Standard_dev     is the standard deviation of the distribution.
         * Cumulative     is a logical value that determines the form of the function.
         * If cumulative is TRUE, NORMDIST returns the cumulative distribution function;
         * if FALSE, it returns the probability mass function.
         *
         *
         * ********************************************************************************
         * IMPORTANT NOTE: when Cumulative=TRUE the results are not accurate to 9 siginfiicant digits in all cases
         * (When cumulative = TRUE, the formula is the integral from negative infinity to x of the given formula)
         * ********************************************************************************
         */
        fun calcNormdist(operands: Array<Ptg>): Ptg {
            if (operands.size < 4) return PtgErr(PtgErr.ERROR_VALUE)
            try {
                //If mean or standard_dev is nonnumeric, NORMDIST returns the #VALUE! error value.
                val x = operands[0].doubleVal
                val mean = operands[1].doubleVal
                // if standard_dev ≤ 0, NORMDIST returns the #NUM! error value.
                val stddev = operands[2].doubleVal
                if (stddev <= 0) return PtgErr(PtgErr.ERROR_NUM)
                val cumulative = PtgCalculator.getBooleanValue(operands[3])
                // If mean = 0, standard_dev = 1, and cumulative = TRUE, NORMDIST returns the standard normal distribution, NORMSDIST.
                if (mean == 0.0 && stddev == 1.0 && cumulative)
                    return calcNormsdist(operands)

                if (!cumulative) {    // return the probability mass function. *** definite excel algorithm
                    var a = Math.sqrt(2.0 * Math.PI * Math.pow(stddev, 2.0))
                    a = 1.0 / a
                    var exp = Math.pow(x - mean, 2.0)
                    exp = exp / (2 * Math.pow(stddev, 2.0))
                    val b = Math.exp(-exp)
                    return PtgNumber(a * b)
                } else {
                    // When cumulative = TRUE, the formula is the integral from negative infinity to x of the given formula.
                    // = the cumulative distribution function
                    val o = arrayOf<Ptg>(PtgNumber((x - mean) / (stddev * Math.sqrt(2.0))))
                    val erf = EngineeringCalculator.calcErf(o)
                    val cdf = 0.5 * (1 + erf.doubleVal)
                    return PtgNumber(cdf)
                    /*			 // try this:
			 Ptg[] o= { new PtgNumber((x-mean)/(stddev))};
			 return calcNormsdist(o);
	*/
                }
            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_VALUE)
            }

        }

        /**
         * The NORMSDIST function returns the result of the standard normal cumulative distribution function
         * for a particular value of the random variable X. The Excel function adheres to the
         * following mathematical approximation, P(x), of the following
         * standard normal cumulative distribution function (CDF):
         * P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x), where
         *
         * Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
         * t = 1/(1+px)
         * p = 0.2316419
         * b1 = 0.319381530
         * b2 = -0.356563782
         * b3 = 1.781477937
         * b4 = -1.821255978
         * b5 = 1.330274429
         *
         * with the following parameters:
         *
         * abs(error(x))<7.5 * 10^-8
         *
         * The NORMSDIST function returns the result of the standard normal
         * CDF for a standard normal random variable Z with a mean of 0 (zero)
         * and a standard deviation of 1. The CDF is found by taking the integral
         * of the following standard normal probability density function
         *
         * Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
         *
         * from negative infinity to the value (z) of the random variable in question.
         * The result of the integral gives the probability that Z will occur between the
         * values of negative infinity and z.
         *
         * NORMSDIST(z) must be evaluated by using an approximation procedure.
         * Earlier versions of Excel used the same procedure for all values of z.
         * For Excel 2003, two different approximations are used:
         * one for |z| less than or equal to five, and a second for |z| greater than five.
         * The two new procedures are each more accurate than the previous procedure
         * over the range that they are applied. In earlier versions of Excel,
         * accuracy deteriorates in the tails of the distribution yielding three significant
         * digits for z = 4 as reported in Knusel's paper. Also, in the neighborhood of z = 1.2,
         * NORMSDIST yields only six significant digits. However, in practice, this is likely
         * to be sufficient for most users.
         *
         * INFO ATP DEFINITION NORMDIST NOVEMBER 2006:
         * The NORMSDIST function returns the result of the standard normal cumulative distribution function for a particular value of the random variable X.
         * The Microsoft Excel function adheres to the following mathematical approximation, P(x), of the standard normal CDF
         *
         * P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x), where
         *
         * Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
         * t = 1/(1+px)
         * p = 0.2316419
         * b1 = 0.319381530
         * b2 = -0.356563782
         * b3 = 1.781477937
         * b4 = -1.821255978
         * b5 = 1.330274429
         *
         *
         * with these parameters, abs(error(x))<7.5 * 10^-8.
         *
         *
         * In summary, if you use Excel 2002 and earlier, you should be satisfied with NORMSDIST.
         * However, if you must have highly accurate NORMSDIST(z) values for z far from 0
         * (such as |z| greater than or equal to four), Excel 2003 might be required.
         * NORMSDIST(-4) = 0.0000316712; earlier versions would be accurate only as far as 0.0000317.
         *
         * from a forum:
         * Take into consideration that Z is related to x, xm(mean) and s(std.dev.)
         * through the expression Z = (x - xm) / s.
         * This means that as soon as you get Z, you can proceed and calculate the integral of the
         * CDF by using P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x).
         *
         * I wish you good code.
         *
         * Some other identities that express NORMSDIST in terms of other functions
         * that have no closed form are
         * NormSDist(x) = ErfC(-x/Sqrt(2))/2 = (1-Erf(-x/Sqrt(2)))/2 for x<=0
         * NormSDist(x) = 1-ErfC(x/Sqrt(2))/2 = (1+Erf(x/Sqrt(2)))/2 for x>=0
         * NormSDist(x) = (1â€“GammaDist(x^2/2,1/2,1,TRUE))/2 for x<=0
         * NormSDist(x) = (1+GammaDist(x^2/2,1/2,1,TRUE))/2 for x>=0
         * NormSDist(x) = ChiDist(x^2,1)/2 for x<=0
         * NormSDist(x) = 1-ChiDist(x^2,1)/2 for x>=0
         *
         * // for 2002:
         * The NORMSDIST function returns the result of the standard normal cumulative distribution
         * function for a particular value of the random variable X. The Excel function adheres to the
         * following mathematical approximation, P(x), of the following standard normal cumulative
         * distribution function (CDF)
         *
         * P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x), where
         *
         * Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
         * t = 1/(1+px)
         * p = 0.2316419
         * b1 = 0.319381530
         * b2 = -0.356563782
         * b3 = 1.781477937
         * b4 = -1.821255978
         * b5 = 1.330274429
         *
         *
         * with the following parameters:
         * abs(error(x))<7.5 * 10^-8
         *
         * The NORMSDIST function returns the result of the standard normal CDF for a standard
         * normal random variable Z with a mean of 0 (zero) and a standard deviation of 1. The CDF
         * is found by taking the integral of the following standard normal probability density
         * function
         *
         * Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
         *
         * from negative infinity to the value (z) of the random variable in question. The result
         * of the integral gives the probability that Z will occur between the values of negative
         * infinity and z. 	  *
         *
         * from openoffice:
         * The wrong results in NORMSDIST are due to cancellation for small negative
         * values, where gauss() is near -0.5
         * The problem can be solved in two ways:
         * (1) Use NORMSDIST(x)= 0.5*ERFC(-x/SQRT(2)). Unfortunaly ERFC is only an addin
         * function, see my issue 97091.
         * (2) Use NORMSDIST(x)
         * = 0.5+0.5*GetLowRegIGamma(0.5,0.5*x*x) for x>=0
         * = 0.5*GetUpRegIGamma(0.5,0.5*x*x)      for x<0
         *
         *
         * From a forum:
         * For z less than 2, ERF = 2/SQRT (pi) * e^(-z^2) * z (1+ (2z^2)/3 + ((2z^2)^2)/15 + â€¦
         * For z greater than 2, ERF = 1- (e^(-z^2))/(SQRT(pi)) * (1/z - 1/(2z^3) + 3/(4z^5) -â€¦.)
         */
        /**
         * NORMSDIST
         * Returns the standard normal cumulative distribution
         *
         *
         * NORMSDIST(z) returns the probability that the observed value of a
         * standard normal random variable will be less than or equal to z.
         * A standard normal random variable has mean 0 and standard deviation 1
         * (and also variance 1 because variance = standard deviation squared).
         *
         *
         * NOTE: THIS FUNCTION IS ACCURATE AS COMPARED TO EXCEL VALUES ONLY UP TO 7 SIGNIFICANT DIGITS
         */
        fun calcNormsdist(operands: Array<Ptg>): Ptg {
            if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
            try {
                val result: Double
                val x = operands[0].doubleVal
                val b1 = 0.319381530
                val b2 = -0.356563782
                val b3 = 1.781477937
                val b4 = -1.821255978
                val b5 = 1.330274429
                val p = 0.2316419
                val c = 0.39894228

                // below is consistently correct to at least 7 decimals using a range of test values
                if (x >= 0.0) {
                    val t = 1.0 / (1.0 + p * x)
                    result = 1.0 - c * Math.exp(-x * x / 2.0) * t * (t * (t * (t * (t * b5 + b4) + b3) + b2) + b1)
                } else {
                    val t = 1.0 / (1.0 - p * x)
                    result = c * Math.exp(-x * x / 2.0) * t * (t * (t * (t * (t * b5 + b4) + b3) + b2) + b1)
                }
                /*
			// try this one:
		 double z= (1/(Math.sqrt(2*Math.PI))*Math.exp(-Math.pow(x, 2)/2.0));
		 double t = 1/(1+p*x);
	     double e= EngineeringCalculator.calcErf(operands).getDoubleVal();
		 result = 1 -z*(b1*t+b2*Math.pow(t, 2)+b3*Math.pow(t, 3)+b4*Math.pow(t, 4)+b5*Math.pow(t, 5))+e;
*/
                val bd = BigDecimal(result)
                bd.setScale(15, java.math.RoundingMode.HALF_UP)
                return PtgNumber(bd.toDouble())
            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_VALUE)
            }

        }

        /**
         * NORMSINV
         * Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.
         *
         *
         * Syntax
         * NORMSINV(probability)
         * Probability   is a probability corresponding to the normal distribution.
         *
         *
         * If probability is nonnumeric, NORMSINV returns the #VALUE! error value.
         * If probability < 0 or if probability > 1, NORMSINV returns the #NUM! error value.
         *
         *
         * Because the calculation of the NORMSINV function uses a systematic search
         * over the returned values of the NORMSDIST function, the accuracy of the
         * NORMSDIST function is critical.
         *
         *
         * Also, the search must be sufficiently refined that it "homes in" on an appropriate
         * answer. To use the textbook Normal probability distribution table as an analogy,
         * entries in the table must be accurate. Also, the table must contain so many entries
         * that you can find the appropriate row of the table that yields a probability that is
         * correct to a specific number of decimal places.	Instead, individual entries are computed
         * on demand as the search through the "table"
         *
         *
         * However, the table must be accurate and the search must continue far enough
         * that it does not stop prematurely at an answer that has a corresponding probability
         * (or row of the table) that is too far from the value of p that you use in the call to
         * NORMSINV(p). Therefore, the NORMSINV function has been improved in the following ways:
         *
         *
         * - The accuracy of the NORMSDIST function has been improved.
         * - The search process has been improved to increase refinement.
         *
         *
         * The NORMSDIST function has been improved in Excel 2003 and in later versions of Excel.
         * Typically, inaccuracies in earlier versions of Excel occur for extremely small or extremely
         * large values of p in NORMSINV(p). The values in Excel 2003 and in later versions of Excel
         * are much more accurate.
         *
         *
         * Accuracy of NORMSDIST has been improved in Excel 2003 and in later versions of Excel.
         * In earlier versions of Excel, a single computational procedure was used for all values
         * of z. Results were essentially accurate to 7 decimal places, more than sufficient for
         * most practical examples.
         *
         *
         * Results in earlier versions of Excel
         * **	The accuracy of the NORMSINV function depends on two factors. Because the calculation of the NORMSINV
         * function uses a systematic search over the returned values of the NORMSDIST function, the accuracy of
         * the NORMSDIST function is critical.
         *
         *
         * Also, the search must be sufficiently refined that it "homes in" on an appropriate answer.
         * To use the textbook Normal probability distribution table as an analogy, entries in the table
         * must be accurate. Also, the table 	must contain so many entries that you can find the
         * appropriate row of the table that yields a probability that is correct to a specific number
         * of decimal places.
         *
         *
         * ' This function is a replacement for the Microsoft Excel Worksheet function NORMSINV.
         * ' It uses the algorithm of Peter J. Acklam to compute the inverse normal cumulative
         * ' distribution. Refer to http://home.online.no/~pjacklam/notes/invnorm/index.html for
         * ' a description of the algorithm.
         * ' Adapted to VB by Christian d'Heureuse, http://www.source-code.biz.
         * Public Function NormSInv(ByVal p As Double) As Double
         * Const a1 = -39.6968302866538, a2 = 220.946098424521, a3 = -275.928510446969
         * Const a4 = 138.357751867269, a5 = -30.6647980661472, a6 = 2.50662827745924
         * Const b1 = -54.4760987982241, b2 = 161.585836858041, b3 = -155.698979859887
         * Const b4 = 66.8013118877197, b5 = -13.2806815528857, c1 = -7.78489400243029E-03
         * Const c2 = -0.322396458041136, c3 = -2.40075827716184, c4 = -2.54973253934373
         * Const c5 = 4.37466414146497, c6 = 2.93816398269878, d1 = 7.78469570904146E-03
         * Const d2 = 0.32246712907004, d3 = 2.445134137143, d4 = 3.75440866190742
         * Const p_low = 0.02425, p_high = 1 - p_low
         * Dim q As Double, r As Double
         * If p < 0 Or p > 1 Then
         * Err.Raise vbObjectError, , "NormSInv: Argument out of range."
         * ElseIf p < p_low Then
         * q = Sqr(-2 * Log(p))
         * NormSInv = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
         * ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
         * ElseIf p <= p_high Then
         * q = p - 0.5: r = q * q
         * NormSInv = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
         * (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
         * Else
         * q = Sqr(-2 * Log(1 - p))
         * NormSInv = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
         * ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
         * End If
         * End Function
         *
         *
         * NORMSINV= NORMINV(p; 0; 1)
         */
        private fun expm1(x: Double): Double {
            val DBL_EPSILON = 0.0000001
            var y: Double
            val a = Math.abs(x)

            if (a < DBL_EPSILON) return x
            if (a > 0.697) return Math.exp(x) - 1  /* negligible cancellation */

            if (a > 1e-8)
                y = Math.exp(x) - 1
            else
            /* Taylor expansion, more accurate in this range */
                y = (x / 2 + 1) * x

            /* Newton step for solving   log(1 + y) = x   for y : */
            /* WARNING: does not work for y ~ -1: bug in 1.5.0 -- fixed??*/
            y -= (1 + y) * (Math.log(1 + y) - x)
            return y
        }

        private fun quartile(p: Double, mu: Double, sigma: Double): Double {
            var p = p
            val lower_tail = true
            val log_p = false
            val R_D__0 = 0.0
            val R_D__1 = 1.0
            val R_DT_0 = 0.0    //((lower_tail) ? R_D__0 : R_D__1);      /* 0 */
            val R_DT_1 = 1.0    // ((lower_tail) ? R_D__1 : R_D__0)      /* 1 */

            val p_: Double
            val q: Double
            var r: Double
            var `val`: Double
            if (p == R_DT_0) return java.lang.Double.NEGATIVE_INFINITY
            if (p == R_DT_1) return java.lang.Double.POSITIVE_INFINITY
            //R_Q_P01_check(p);

            if (sigma < 0) return 0.0
            if (sigma == 0.0) return mu

            p = if (log_p) if (lower_tail) Math.exp(p) else -expm1(p) else if (lower_tail) p else 1 - p
            p_ = if (log_p) if (lower_tail) Math.exp(p) else -expm1(p) else if (lower_tail) p else 1 - p/* real lower_tail prob. p */
            q = p_ - 0.5

            if (Math.abs(q) <= .425) {/* 0.075 <= p <= 0.925 */
                r = .180625 - q * q
                `val` = q * (((((((r * 2509.0809287301226727 + 33430.575583588128105) * r + 67265.770927008700853) * r + 45921.953931549871457) * r + 13731.693765509461125) * r + 1971.5909503065514427) * r + 133.14166789178437745) * r + 3.387132872796366608) / (((((((r * 5226.495278852854561 + 28729.085735721942674) * r + 39307.89580009271061) * r + 21213.794301586595867) * r + 5394.1960214247511077) * r + 687.1870074920579083) * r + 42.313330701600911252) * r + 1.0)
            } else { /* closer than 0.075 from {0,1} boundary */

                /* r = min(p, 1-p) < 0.075 */
                if (q > 0)
                    r = (if (log_p) if (lower_tail) -expm1(p) else Math.exp(p) else if (lower_tail) 1 - p else p)/* 1-p */
                else
                    r = p_/* = R_DT_Iv(p) ^=  p */

                r = Math.sqrt(-/* else */if (log_p && (lower_tail && q <= 0 || !lower_tail && q > 0)) p else Math.log(r))
                /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

                if (r <= 5.0) { /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
                    r += -1.6
                    `val` = (((((((r * 7.7454501427834140764e-4 + .0227238449892691845833) * r + .24178072517745061177) * r + 1.27045825245236838258) * r + 3.64784832476320460504) * r + 5.7694972214606914055) * r + 4.6303378461565452959) * r + 1.42343711074968357734) / (((((((r * 1.05075007164441684324e-9 + 5.475938084995344946e-4) * r + .0151986665636164571966) * r + .14810397642748007459) * r + .68976733498510000455) * r + 1.6763848301838038494) * r + 2.05319162663775882187) * r + 1.0)
                } else { /* very close to  0 or 1 */
                    r += -5.0
                    `val` = (((((((r * 2.01033439929228813265e-7 + 2.71155556874348757815e-5) * r + .0012426609473880784386) * r + .026532189526576123093) * r + .29656057182850489123) * r + 1.7848265399172913358) * r + 5.4637849111641143699) * r + 6.6579046435011037772) / (((((((r * 2.04426310338993978564e-15 + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r + 7.868691311456132591e-4) * r + .0148753612908506148525) * r + .13692988092273580531) * r + .59983220655588793769) * r + 1.0)
                }

                if (q < 0.0)
                    `val` = -`val`
                /* return (q >= 0.)? r : -r ;*/
            }
            return mu + sigma * `val`
        }

        fun calcNormInv(operands: Array<Ptg>): Ptg {
            try {

                val p = operands[0].doubleVal
                if (p < 0 || p > 1) return PtgErr(PtgErr.ERROR_NUM)
                val mean = operands[1].doubleVal
                val stddev = operands[2].doubleVal
                if (stddev <= 0) return PtgErr(PtgErr.ERROR_NUM)
                // If mean = 0 and standard_dev = 1, NORMINV uses the standard normal inverse (see NORMSINV).
                val result = quartile(p, mean, stddev)
                return PtgNumber(result)
            } catch (e: Exception) {
            }

            return PtgErr(PtgErr.ERROR_VALUE)
        }


        fun calcNormsInv(operands: Array<Ptg>): Ptg {
            if (operands.size != 1) return PtgCalculator.valueError
            try {
                val x = operands[0].doubleVal
                if (x < 0 || x > 1)
                    return PtgErr(PtgErr.ERROR_NUM)
                /*
		  * the algorithm is supposed to iterate over NORMSDIST values using Newton-Raphson's approximation
		  * Newton-Raphson uses an iterative process to approach one root of a function (i.e the zero of the function or
		  * where the function = 0)
		  * Newton-Raphson is in the form of:
		  *
		  * Xn+1= Xn- (f(xn)/f'(xn)
		  * where Xn is the current known X value, f(xn) is the value of the function at X, f'(Xn) is the derivative or slope
		  * at X, Xn+1 is the next X value. Essentially, f'(xn) represents (f(x)/delta x) so f(xn)/f'(xn)== delta x.
		  * the more iterations we run over, the closer delta x will be to 0.
		  *
		  * The Newton-Raphson method does not always work, however. It runs into problems in several places.
			What would happen if we chose an initial x-value of x=0? We would have a "division by zero" error, and would not be able to proceed.
			You may also consider operating the process on the function f(x) = x1/3, using an inital x-value of x=1.
			Do the x-values converge? Does the delta-x decrease toward zero (0)?

			 is the derivative of the standard normal distribution = the standard probability density function???
		  */
                /* below is not accurate enough - but N-R approximation is impossible ((:*/
                // Coefficients in rational approximations
                val a = doubleArrayOf(-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00)

                val b = doubleArrayOf(-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01)

                val c = doubleArrayOf(-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00)

                val d = doubleArrayOf(7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00)

                // Define break-points.
                val plow = 0.02425
                val phigh = 1 - plow
                val result: Double
                // Rational approximation for lower region:
                if (x < plow) {
                    val q = Math.sqrt(-2 * Math.log(x))
                    val r = BigDecimal((((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1))
                    r.setScale(15, java.math.RoundingMode.HALF_UP)
                    return PtgNumber(r.toDouble())
                }

                // Rational approximation for upper region:
                if (phigh < x) {
                    val q = Math.sqrt(-2 * Math.log(1 - x))
                    val r = BigDecimal(-(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1))
                    r.setScale(15, java.math.RoundingMode.HALF_UP)
                    return PtgNumber(r.toDouble())
                }

                // Rational approximation for central region:
                val q = x - 0.5
                val r = q * q
                val rr = BigDecimal((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1))

                rr.setScale(15, java.math.RoundingMode.HALF_UP)
                return PtgNumber(rr.toDouble())
            } catch (e: Exception) {
                return PtgCalculator.valueError
            }

        }

        /**
         * PEARSON
         * Returns the Pearson product moment correlation coefficient
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcPearson(operands: Array<Ptg>): Ptg {
            return calcCorrel(operands)
        }
        /*
PERCENTILE
 Returns the k-th percentile of values in a range

PERCENTRANK
 Returns the percentage rank of a value in a data set

PERMUT
 Returns the number of permutations for a given number of objects

POISSON
 Returns the Poisson distribution

PROB
 Returns the probability that values in a range are between two limits
 */

        /**
         * QUARTILE
         * Returns the quartile of a data set
         */
        fun calcQuartile(operands: Array<Ptg>): Ptg {
            val aveoperands = arrayOfNulls<Ptg>(1)
            aveoperands[0] = operands[0]
            val allVals = PtgCalculator.getAllComponents(aveoperands)
            val t = CompatibleVector()
            val retval = 0.0
            for (i in allVals.indices) {
                val p = allVals[i]
                try {
                    val d = Double(p.value.toString())
                    t.addOrderedDouble(d)
                } catch (e: NumberFormatException) {
                    Logger.logErr(e)
                }

            }

            val dub = arrayOfNulls<Double>(t.size)
            t.toTypedArray()
            val quart: Int?
            val o = operands[1].value
            if (o is Int)
                quart = operands[1].value as Int
            else
                quart = Integer.valueOf((operands[1].value as Double).toInt())

            val quartile = quart!!.toFloat()
            if (quart.toInt() == 0) {    // return minimum value
                return PtgNumber(dub[0].toDouble())
            } else if (quart.toInt() == 4) {    // return maximum value
                return PtgNumber(dub[t.size - 1].toDouble())
            } else if (quart.toInt() > 4 || quart.toInt() < 0) {
                return PtgErr(PtgErr.ERROR_NUM)
            }
            // find the kth smallest
            var kk = quartile / 4
            kk = (dub.size - 1) * kk
            kk++
            // truncate k, but keep the remainder.
            var k = -1
            var remainder = 0f
            if (kk % 1 != 0f) {
                remainder = kk % 1
                var s = kk.toString()
                var ss = s.substring(s.indexOf("."))
                ss = "0$ss"
                remainder = Float(ss)
                s = s.substring(0, s.indexOf("."))
                k = Integer.valueOf(s).toInt()
            } else {
                k = kk.toInt() / 1
            }
            if (k >= dub.size)
                return PtgErr(PtgErr.ERROR_VALUE)
            val firstVal = dub[k - 1].toDouble()
            val secondVal = dub[k].toDouble()
            val output = firstVal + remainder * (secondVal - firstVal)
            return PtgNumber(output)
        }

        /**
         * RANK
         * Returns the rank of a number in a list of numbers
         *
         *
         * RANK(number,ref,order)
         * Number   is the number whose rank you want to find.
         * Ref   is an array of, or a reference to, a list of numbers. Nonnumeric values in ref are ignored.
         * Order   is a number specifying how to rank number.
         *
         *
         * If order is 0 (zero) or omitted, Microsoft Excel ranks number as if ref were a list sorted in descending order.
         * If order is any nonzero value, Microsoft Excel ranks number as if ref were a list sorted in ascending order.
         */
        fun calcRank(operands: Array<Ptg>): Ptg {
            // the number
            if (operands.size < 2) return PtgErr(PtgErr.ERROR_VALUE)
            val num = operands[0]
            var theNum: Double? = null
            try {
                val o = num.value
                if (o == "")
                    theNum = 0.0
                else
                    theNum = Double(o.toString())
            } catch (nfm: NumberFormatException) {
                return PtgErr()
            }

            //ascending or decending?
            var ascending = true
            if (operands.size < 3) {
                ascending = false
            } else if (operands[2] is PtgMissArg) {
                ascending = false
            } else {
                val order = operands[2] as PtgInt
                val i = order.`val`
                if (i == 0) ascending = false
            }
            val aveoperands = arrayOfNulls<Ptg>(1)
            aveoperands[0] = operands[1]
            val refs = PtgCalculator.getAllComponents(aveoperands)
            val retList = CompatibleVector()
            val retval = 0.0
            for (i in refs.indices) {
                val p = refs[i]
                try {
                    val d = Double(p.value.toString())
                    retList.addOrderedDouble(d)
                } catch (e: NumberFormatException) {
                }

            }
            val dubRefs = arrayOfNulls<Double>(retList.size)
            if (ascending) {
                retList.toTypedArray()
            } else {
                for (i in dubRefs.indices) {
                    dubRefs[i] = retList.last() as Double
                    retList.removeAt(retList.size - 1)
                }
            }
            var res = -1
            var i = 0
            while (i < dubRefs.size) {
                if (dubRefs[i].toString().equals(theNum.toString(), ignoreCase = true)) {
                    res = i + 1
                    i = dubRefs.size
                }
                i++
            }
            return if (res == -1) {
                PtgErr(PtgErr.ERROR_NA)
            } else {
                PtgInt(res)
            }

        }


        /**
         * RSQ
         * Returns the square of the Pearson product moment correlatin coefficient
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcRsq(operands: Array<Ptg>): Ptg {
            val p = calcPearson(operands) as PtgNumber
            var d = p.`val`
            d = d * d
            return PtgNumber(d)
        }
        /*
SKEW
 Returns the skewness of a distribution
 */

        /**
         * SLOPE
         * Returns the slope of the linear regression line
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcSlope(operands: Array<Ptg>): Ptg {
            if (operands.size != 2)
                return PtgErr(PtgErr.ERROR_VALUE)
            val yvals = PtgCalculator.getDoubleValueArray(operands[0])
            val xvals = PtgCalculator.getDoubleValueArray(operands[1])
            if (xvals == null || yvals == null) return PtgErr(PtgErr.ERROR_NA)//20090130 KSC: propagate error
            var sumXVals = 0.0
            for (i in xvals.indices) {
                sumXVals += xvals[i]
            }
            var sumYVals = 0.0
            for (i in yvals.indices) {
                sumYVals += yvals[i]
            }
            var sumXYVals = 0.0
            for (i in yvals.indices) {
                sumXYVals += xvals[i] * yvals[i]
            }
            var sqrXVals = 0.0
            for (i in xvals.indices) {
                sqrXVals += xvals[i] * xvals[i]
            }
            val toparg = sumXVals * sumYVals - sumXYVals * yvals.size
            val bottomarg = sumXVals * sumXVals - sqrXVals * xvals.size
            val res = toparg / bottomarg
            return PtgNumber(res)
        }

        /**
         * SMALL
         * Returns the k-th smallest value in a data set
         *
         *
         * SMALL(array,k)
         * Array   is an array or range of numerical data for which you want to determine the k-th smallest value.
         * K   is the position (from the smallest) in the array or range of data to return.
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcSmall(operands: Array<Ptg>): Ptg {
            if (operands.size != 2) return PtgErr(PtgErr.ERROR_VALUE)
            val rng = operands[0]
            val array = PtgCalculator.getAllComponents(rng)
            if (array.size == 0) return PtgErr(PtgErr.ERROR_NUM)
            val k = PtgCalculator.getDoubleValueArray(operands[1])!![0].toInt()
            if (k <= 0 || k > array.size)
                return PtgErr(PtgErr.ERROR_NUM)

            val sortedValues = CompatibleVector()
            for (i in array.indices) {
                val p = array[i]
                try {
                    val d = Double(p.value.toString())
                    sortedValues.addOrderedDouble(d)
                } catch (e: NumberFormatException) {
                }

            }
            try {
                return PtgNumber((sortedValues[k - 1] as Double).toDouble())
            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_VALUE)
            }

        }

        /*
STANDARDIZE
 Returns a normalized value
 */

        /**
         * STDEV(number1,number2, ...)
         *
         *
         * Number1,number2, ...   are 1 to 255 number arguments corresponding to a sample of a population.
         * You can also use a single array or a reference to an array instead of arguments separated by commas.
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcStdev(operands: Array<Ptg>): Ptg {
            val allVals = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
            var sqrDev = 0.0
            for (i in allVals.indices) {
                val p = calcAverage(operands) as PtgNumber
                val ave = p.`val`
                sqrDev += Math.pow(allVals[i] - ave, 2.0)
            }
            val retval = Math.sqrt(sqrDev / (allVals.size - 1))
            return PtgNumber(retval)
        }
        /*
STDEVA
 Estimates standard deviation based on a sample, including numbers, text, and logical values

STDEVP
 Calculates standard deviation based on the entire population

STDEVPA
 Calculates standard deviation based on the entire population, including numbers, text, and logical values
 */

        /**
         * STEYX
         * Returns the standard error of the predicted y-value for each x in the regression
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcSteyx(operands: Array<Ptg>): Ptg {
            val arr = arrayOfNulls<Ptg>(1)
            arr[0] = operands[0]
            var pn = calcVarp(arr) as PtgNumber
            var yVarp = pn.`val`
            arr[0] = operands[1]
            pn = calcVarp(arr) as PtgNumber
            var xVarp = pn.`val`
            val y = PtgCalculator.getDoubleValueArray(operands[0]) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
            yVarp *= y.size.toDouble()
            xVarp *= y.size.toDouble()
            pn = calcSlope(operands) as PtgNumber
            val slope = pn.`val`
            var retval = yVarp - slope * slope * xVarp
            retval = retval / (y.size - 2)
            retval = Math.sqrt(retval)
            return PtgNumber(retval)
        }
        /*
TDIST
 Returns the Student's t-distribution

TINV
 Returns the inverse of the Student's t-distribution
 */

        /**
         * TREND
         * Returns values along a linear trend
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcTrend(operands: Array<Ptg>): Ptg {
            if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
            // KSC: THIS FUNCTION DOES NOT WORK AS EXPECTED: TODO: FIX!

            if (true) return PtgErr(PtgErr.ERROR_VALUE)
            // 	Ptg[] forecast = new Ptg[3];
            val forecast = arrayOfNulls<Ptg>(2)
            forecast[0] = operands[0]
            if (operands.size > 1)
                forecast[1] = operands[1]
            // TODO:
            // 	else // If known_x's is omitted, it is assumed to be the array {1,2,3,...} that is the same size as known_y's.
            val newXs: Array<Ptg>
            if (operands.size > 2)
                newXs = PtgCalculator.getAllComponents(operands[2])
            else
                newXs = PtgCalculator.getAllComponents(operands[1])

            var retval = ""
            for (i in newXs.indices) {
                //forecast[0] = newXs[i];
                val p = calcForecast(forecast) as PtgNumber
                val forcst = p.`val`
                retval += "{$forcst},"
            }
            // get rid of trailing comma
            retval = retval.substring(0, retval.length - 1)
            val pa = PtgArray()
            pa.setVal(retval)
            return pa
        }

        /*
TRIMMEAN
 Returns the mean of the interior of a data set

TTEST
 Returns the probability associated with a Student's t-Test
 */

        /**
         * VAR
         * Estimates variance based on a sample
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcVar(operands: Array<Ptg>): Ptg {
            val allVals = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
            var sqrDev = 0.0
            for (i in allVals.indices) {
                val p = calcAverage(operands) as PtgNumber
                val ave = p.`val`
                sqrDev += Math.pow(allVals[i] - ave, 2.0)
            }
            val retval = sqrDev / (allVals.size - 1)
            return PtgNumber(retval)
        }
        /*
VARA
 Estimates variance based on a sample, including numbers, text, and logical values
 */

        /**
         * VARp
         * Estimates variance based on a full population
         *
         * @throws CalculationException
         */
        @Throws(CalculationException::class)
        fun calcVarp(operands: Array<Ptg>): Ptg {
            val allVals = PtgCalculator.getDoubleValueArray(operands) ?: return PtgErr(PtgErr.ERROR_NA)
//20090130 KSC: propagate error
            var sqrDev = 0.0
            for (i in allVals.indices) {
                val p = calcAverage(operands) as PtgNumber
                val ave = p.`val`
                sqrDev += Math.pow(allVals[i] - ave, 2.0)
            }
            val retval = sqrDev / allVals.size
            return PtgNumber(retval)
        }
    }
    /*
VARPA
 Calculates variance based on the entire population, including numbers, text, and logical values
 
WEIBULL
 Returns the Weibull distribution
 
ZTEST
 Returns the two-tailed P-value of a z-test
 
*/

}