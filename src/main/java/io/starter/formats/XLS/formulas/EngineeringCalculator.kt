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
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool


/*
    EngineeringCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/
object EngineeringCalculator {
    var DEBUG = false

    internal val PI_SQRT = Math.sqrt(Math.PI)
    internal val TSQPI = 2.0 / PI_SQRT
    private var firstCall = true

    // ===========================================================================
    /*** Nominator coefficients for approximation to erf in first interval.    */
    private val ERF_A = doubleArrayOf(3.16112374387056560E00, 1.13864154151050156E02, 3.77485237685302021E02, 3.20937758913846947E03, 1.85777706184603153E-1)
    /*** Denominator coefficients for approximation to erf in first interval.  */
    private val ERF_B = doubleArrayOf(2.36012909523441209E01, 2.44024637934444173E02, 1.28261652607737228E03, 2.84423683343917062E03)
    // ===========================================================================
    /*** Nominator coefficients for approximation to erfc in second interval.  */
    private val ERF_C = doubleArrayOf(5.64188496988670089E-1, 8.88314979438837594E0, 6.61191906371416295E01, 2.98635138197400131E02, 8.81952221241769090E02, 1.71204761263407058E03, 2.05107837782607147E03, 1.23033935479799725E03, 2.15311535474403846E-8)
    /*** Denominator coefficients for approximation to erfc in second interval. */
    private val ERF_D = doubleArrayOf(1.57449261107098347E01, 1.17693950891312499E02, 5.37181101862009858E02, 1.62138957456669019E03, 3.29079923573345963E03, 4.36261909014324716E03, 3.43936767414372164E03, 1.23033935480374942E03)
    // ===========================================================================
    /*** Nominator coefficients for approximation to erfc in third interval.   */
    private val ERF_P = doubleArrayOf(3.05326634961232344E-1, 3.60344899949804439E-1, 1.25781726111229246E-1, 1.60837851487422766E-2, 6.58749161529837803E-4, 1.63153871373020978E-2)
    /*** Denominator coefficients for approximation to erfc in third interval.  */
    private val ERF_Q = doubleArrayOf(2.56852019228982242, 1.87295284992346047, 5.27905102951428412E-1, 6.05183413124413191E-2, 2.33520497626869185E-3)
    // ===========================================================================
    internal val THRESHOLD = 0.46875
    internal val SQRPI = 1.0 / PI_SQRT
    internal val X_INF = java.lang.Double.MAX_VALUE
    internal val X_MIN = java.lang.Double.MIN_VALUE
    // private static final double X_NEG = -9.38241396824444;
    internal val X_NEG = -Math.sqrt(Math.log(X_INF / 2))
    internal val X_SMALL = deps
    internal val X_HUGE = 1.0 / (2.0 * Math.sqrt(X_SMALL))
    internal val X_MAX = Math.min(X_INF, 1.0 / (PI_SQRT * X_MIN))
    // static final double X_BIG = 9.194E0;
    internal val X_BIG = 26.543

    private val FEPS_START = 2e-6f
    /** double valued machine precision.  */
    /**
     * float valued machine precision.
     */
    //public static final float FEPS;
    internal
            /**
             * Calculate the machine accuracy,
             * which is the smallest eps with
             * 1<1+eps
             *///FEPS = feps;
    //DEPS =  deps;
    //if (DEBUG)
    //  Logger.logInfo(format("feps:%8.2E  deps:%8.3G", FEPS, DEPS));
    val deps: Double
        get() {
            var feps = FEPS_START
            var fy = 1.0f + feps
            while (fy > 1.0f) {
                feps /= 2.0f
                fy = 1.0f + feps
            }
            var deps = (feps * FEPS_START).toDouble()
            var dy = 1.0 + deps
            while (dy > 1.0) {
                deps /= 2.0
                dy = 1.0 + deps
            }
            return deps
        }

    /*
	BESSELI
	 Returns the modified Bessel function In(x)

	BESSELJ
	 Returns the Bessel function Jn(x)

	BESSELK
	 Returns the modified Bessel function Kn(x)

	BESSELY
	 Returns the Bessel function Yn(x)
*/

    /**
     * BIN2DEC
     * Converts a binary number to decimal
     */
    internal fun calcBin2Dec(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "BIN2DEC")
        val bString = operands[0].string.trim { it <= ' ' }
        // 10 bits at most
        if (bString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        // must det. manually if binary string is negative because parseInt does not
        // handle two's complement input!!!
        val bIsNegative = bString.length == 10 && bString.substring(0, 1).equals("1", ignoreCase = true)

        var dec = 0
        try {
            dec = Integer.parseInt(bString, 2)
            if (bIsNegative) dec -= 1024    // 2^10 (= signed 11 bits)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val pnum = PtgNumber(dec.toDouble())
        if (DEBUG) Logger.logInfo("Result from BIN2DEC= " + pnum.`val`)
        return pnum
    }

    /**
     * BIN2HEX
     * Converts a binary number to hexadecimal
     */
    internal fun calcBin2Hex(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "BIN2HEX")
        val bString = operands[0].string.trim { it <= ' ' }
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        // 10 bits at most
        if (bString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        // must det. manually if binary string is negative because parseInt does not
        // handle two's complement input!!!
        val bIsNegative = bString.length == 10 && bString.substring(0, 1).equals("1", ignoreCase = true)

        var dec: Long
        var hString: String
        try {
            dec = java.lang.Long.parseLong(bString, 2)
            if (bIsNegative) dec -= 1024        // 2^10 (= signed 11 bits)
            hString = java.lang.Long.toHexString(dec).toUpperCase()
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            hString = hString.substring(Math.max(hString.length - 10, 0))
        } else if (places > 0) {
            if (hString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            hString = "0000000000$hString"    // maximum= 10 bits
            hString = hString.substring(hString.length - places)
        }

        val pstr = PtgStr(hString)
        if (DEBUG) Logger.logInfo("Result from BIN2HEX= " + pstr.string)
        return pstr
    }

    /**
     * BIN2OCT
     * Converts a binary number to octal
     */
    internal fun calcBin2Oct(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "Bin2Oct")
        val bString = operands[0].string.trim { it <= ' ' }
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        // 10 bits at most
        if (bString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        // must det. manually if binary string is negative because parseInt does not
        // handle two's complement input!!!
        val bIsNegative = bString.length == 10 && bString.substring(0, 1).equals("1", ignoreCase = true)

        var dec: Int
        var oString: String
        try {
            dec = Integer.parseInt(bString, 2)
            if (bIsNegative) dec -= 1024        // 2^10 (= signed 11 bits)
            oString = Integer.toOctalString(dec)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            oString = oString.substring(Math.max(oString.length - 10, 0))
        } else if (places > 0) {
            if (oString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            oString = "0000000000$oString"    // maximum= 10 bits
            oString = oString.substring(oString.length - places)
        }
        val pstr = PtgStr(oString)
        if (DEBUG) Logger.logInfo("Result from BIN2OCT= " + pstr.string)
        return pstr
    }

    /**
     * COMPLEX
     * Converts real and imaginary coefficients into a complex number
     */
    internal fun calcComplex(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "Complex")
        val real = operands[0].intVal
        val imaginary = operands[1].intVal
        var suffix = "i"
        if (operands.size > 2) {
            suffix = operands[2].string.trim { it <= ' ' }
            if (!(suffix == "i" || suffix == "j")) {
                return PtgErr(PtgErr.ERROR_VALUE)
            }
        }
        var complexString = ""

        // result:  real + imaginary suffix
        // 			real - imaginary suffix
        //			real				(if imaginary==0)
        //			real + suffix 		(if imaginary==1)
        //			imaginary suffix	(if (real==0)
        if (real != 0) {
            complexString = real.toString()
            if (imaginary > 0)
                complexString += " + "
        }

        if (imaginary != 0)
            complexString += (if (Math.abs(imaginary) != 1) imaginary.toString() else "") + suffix

        if (complexString == "") complexString = "0"

        val pstr = PtgStr(complexString)
        if (DEBUG) Logger.logInfo("Result from COMPLEX= " + pstr.string)
        return pstr
    }

    /**
     * CONVERT
     * Converts a number from one measurement system to another
     * Weight and mass 		From_unit or to_unit
     * Gram 					"g"
     * Slug 					"sg"
     * Pound mass (avoirdupois) "lbm"
     * U (atomic mass unit) 	"u"
     * Ounce mass (avoirdupois) "ozm"
     *
     *
     * Distance 				From_unit or to_unit
     * Meter 					"m"
     * Statute mile 			"mi"
     * Nautical mile 			"Nmi"
     * Inch 					"in"
     * Foot 					"ft"
     * Yard 					"yd"
     * Angstrom 				"ang"
     * Pica (1/72 in.) 		"Pica"
     *
     *
     * Time 					From_unit or to_unit
     * Year 					"yr"
     * Day 					"day"
     * Hour 					"hr"
     * Minute 					"mn"
     * Second 					"sec"
     *
     *
     * Pressure 				From_unit or to_unit
     * Pascal 					"Pa"
     * Atmosphere 				"atm"
     * mm of Mercury 			"mmHg"
     *
     *
     * Force 					From_unit or to_unit
     * Newton 					"N"
     * Dyne 					"dyn"
     * Pound force 			"lbf"
     *
     *
     * Energy 					From_unit or to_unit
     * Joule 					"J"
     * Erg 					"e"
     * Thermodynamic calorie 	"c"
     * IT calorie 				"cal"
     * Electron volt 			"eV"
     * Horsepower-hour 		"HPh"
     * Watt-hour 				"Wh"
     * Foot-pound 				"flb"
     * BTU 					"BTU"
     *
     *
     * Power 					From_unit or to_unit
     * Horsepower 				"HP"
     * Watt 					"W"
     *
     *
     * Magnetism 				From_unit or to_unit
     * Tesla 					"T"
     * Gauss 					"ga"
     *
     *
     * Temperature 			From_unit or to_unit
     * Degree Celsius 			"C"
     * Degree Fahrenheit 		"F"
     * Degree Kelvin 			"K"
     *
     *
     * Liquid measure 			From_unit or to_unit
     * Teaspoon 				"tsp"
     * Tablespoon 				"tbs"
     * Fluid ounce 			"oz"
     * Cup 					"cup"
     * U.S. pint 				"pt"
     * U.K. pint 				"uk_pt"
     * Quart 					"qt"
     * Gallon 					"gal"
     * Liter 					"l"
     *
     *
     *
     *
     * The following abbreviated unit prefixes can be prepended to any metric
     * from_unit or to_unit.
     *
     *
     * Prefix Multiplier Abbreviation
     * exa 1E+18 "E"
     * peta 1E+15 "P"
     * tera 1E+12 "T"
     * giga 1E+09 "G"
     * mega 1E+06 "M"
     * kilo 1E+03 "k"
     * hecto 1E+02 "h"
     * dekao 1E+01 "e"
     * deci 1E-01 "d"
     * centi 1E-02 "c"
     * milli 1E-03 "m"
     * micro 1E-06 "u"
     * nano 1E-09 "n"
     * pico 1E-12 "p"
     * femto 1E-15 "f"
     * atto 1E-18 "a"
     *
     *
     *
     *
     * Remarks
     *
     *
     * If the input data types are incorrect, CONVERT returns the #VALUE! error value.
     * If the unit does not exist, CONVERT returns the #N/A error value.
     * If the unit does not support an abbreviated unit prefix, CONVERT returns the #N/A error value.
     * If the units are in different groups, CONVERT returns the #N/A error value.
     * Unit names and prefixes are case-sensitive.
     */
    private fun findUnits(u: String, units: Array<String>): Int {
        val bFound = false
        for (i in units.indices) {
            if (u == units[i]) return i
        }
        return -1
    }

    private fun prefixMultiplier(p: String, prefixes: Array<String>): Double {
        var multiplier = 1.0
        if (p == "") return multiplier
        for (i in prefixes.indices) {
            if (p == prefixes[i]) {
                when (i) {
                    0    // "E" 	exa
                    -> multiplier = 1E+18
                    1 // "P" 	peta
                    -> multiplier = 1E+15
                    2 // "T" 	tera
                    -> multiplier = 1E+12
                    3 // "G" 	giga
                    -> multiplier = 1E+09
                    4 // "M"	mega
                    -> multiplier = 1E+06
                    5 // "k"	kilo
                    -> multiplier = 1E+03
                    6 // "h"	hecto
                    -> multiplier = 1E+02
                    7 // "e"	dekao
                    -> multiplier = 1E+01
                    8 // "d" 	deci
                    -> multiplier = 1E-01
                    9 // "c"	centi
                    -> multiplier = 1E-02
                    10 // "m"	milli
                    -> multiplier = 1E-03
                    11 // "u"	micro
                    -> multiplier = 1E-06
                    12 // "n"	nano
                    -> multiplier = 1E-09
                    13 // "p"	pico
                    -> multiplier = 1E-12
                    14 // "f"	femto
                    -> multiplier = 1E-15
                    15 // "a"	atto
                    -> multiplier = 1E-18
                }
                return multiplier
            }
        }

        return multiplier
    }

    internal fun calcConvert(operands: Array<Ptg>): Ptg {
        if (operands.size < 3) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "CONVERT")
        val number = operands[0].doubleVal
        var fromUnits = operands[1].string.trim { it <= ' ' }
        var toUnits = operands[2].string.trim { it <= ' ' }

        val allUnits = arrayOf("g", "sg", "lbm", "u", "ozm", "m", "mi", "Nmi", "in", "ft", "yd", "ang", "Pica", "yr", "day", "hr", "mn", "sec", "Pa", "atm", "mmHg", "N", "dyn", "lbf", "J", "e", "c", "cal", "eV", "HPh", "Wh", "flb", "BTU", "HP", "W", "T", "ga", "C", "F", "K", "tsp", "tbs", "oz", "cup", "pt", "uk_pt", "qt", "gal", "l")
        val weightUnits = arrayOf("g", "sg", "lbm", "u", "ozm")
        val distanceUnits = arrayOf("m", "mi", "Nmi", "in", "ft", "yd", "ang", "Pica")
        val timeUnits = arrayOf("yr", "day", "hr", "mn", "sec")
        val pressureUnits = arrayOf("Pa", "atm", "mmHg")
        val forceUnits = arrayOf("N", "dyn", "lbf")
        val energyUnits = arrayOf("J", "e", "c", "cal", "eV", "HPh", "Wh", "flb", "BTU")
        val powerUnits = arrayOf("HP", "W")
        val magnetismUnits = arrayOf("T", "ga")
        val temperatureUnits = arrayOf("C", "F", "K")
        val liquidMeasureUnits = arrayOf("tsp", "tbs", "oz", "cup", "pt", "uk_pt", "qt", "gal", "l")

        // for any metric unit, may be prefixed with
        var fromPrefix = ""
        var toPrefix = ""
        val metricPrefixes = arrayOf("E", "P", "T", "G", "M", "k", "h", "e", "d", "c", "m", "u", "n", "p", "f", "a")

        // first, see if fromUnits and toUnits are in list of acceptable units
        if (findUnits(fromUnits, allUnits) < 0) { // doesn't match; strip prefix and try again
            if (fromUnits.length > 1) {
                fromPrefix = fromUnits.substring(0, 1)
                fromUnits = fromUnits.substring(1)
                // now recheck
                if (findUnits(fromUnits, allUnits) < 0)
                    return PtgErr(PtgErr.ERROR_NA)
                // make sure that prefix is acceptable
                if (findUnits(fromPrefix, metricPrefixes) < 0)
                    return PtgErr(PtgErr.ERROR_NA)
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }
        if (findUnits(toUnits, allUnits) < 0) {// doesn't match; strip prefix and try again
            if (toUnits.length > 1) {
                toPrefix = toUnits.substring(0, 1)
                toUnits = toUnits.substring(1)
                // now recheck
                if (findUnits(toUnits, allUnits) < 0)
                    return PtgErr(PtgErr.ERROR_NA)
                // make sure that prefix is acceptable
                if (findUnits(toPrefix, metricPrefixes) < 0)
                    return PtgErr(PtgErr.ERROR_NA)
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // at here, we know that the prefixes and units are found, but we don't know if they match or make sense ...
        var result = 0.0
        var from = 0.0
        var i: Int
        var j = -1

        // WEIGHT conversion
        if ((i = findUnits(fromUnits, weightUnits)) >= 0) {
            j = findUnits(toUnits, weightUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                // { "g", "sg", "lbm", "u", "ozm" }
                // get fromUnit in grams
                when (i) {
                    // from:
                    0 // "g"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes)
                    1 // "sg"
                    -> from = 14593.84241892870000000000 * number
                    2 // "lbm"
                    -> from = 453.5923097488115 * number
                    3 // "u"
                    -> from = 1.660531004604650E-24 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    4 // "ozm"
                    -> from = 28.349515207973 * number
                }
                // now convert
                when (j) {
                    0    // "g"
                    -> result = from / prefixMultiplier(toPrefix, metricPrefixes)
                    1 // "sg"
                    -> result = from * 0.00006852205000534780
                    2 // "lbm"
                    -> result = from * 0.00220462291469134000
                    3 // "u"
                    -> result = from * 6.02217000000000000000E+23 / prefixMultiplier(toPrefix, metricPrefixes)
                    4 // "ozm"
                    -> result = from * 0.03527397180036270000
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // DISTANCE conversion
        if (j == -1 && (i = findUnits(fromUnits, distanceUnits)) >= 0) {
            j = findUnits(toUnits, distanceUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //  "m", "mi", "Nmi", "in", "ft", "yd", "ang", "Pica" ;
                // get fromUnits in m
                when (i) {
                    0    //"m"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes)
                    1 //"mi"
                    -> from = 1609.344000000000 * number
                    2 //"Nmi"
                    -> from = 1852.000000000000 * number
                    3 //"in"
                    -> from = 0.025400000000 * number
                    4 // "ft"
                    -> from = 0.304800000000 * number
                    5 // "yd"
                    -> from = 0.914400000300 * number
                    6 // "ang"
                    -> from = 0.000000000100 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    7 // "Pica"
                    -> from = 0.00035277777777780000 * number
                }
                when (j) {
                    0    //"m"
                    -> result = from / prefixMultiplier(toPrefix, metricPrefixes)
                    1 //"mi"
                    -> result = 0.00062137119223733400 * from
                    2 //"Nmi"
                    -> result = 0.00053995680345572400 * from
                    3 //"in"
                    -> result = 39.37007874015750000000 * from
                    4 // "ft"
                    -> result = 3.28083989501312000000 * from
                    5 // "yd"
                    -> result = 1.09361329797891000000 * from
                    6 // "ang"
                    -> result = 10000000000.000000000000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    7 // "Pica"
                    -> result = 2834.64566929116000000000 * from
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // TIME conversion
        if (j == -1 && (i = findUnits(fromUnits, timeUnits)) >= 0) {
            j = findUnits(toUnits, timeUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"yr", "day", "hr", "mn", "sec"
                when (i) {
                    0    // "yr"
                    -> from = number
                    1 // "day"
                    -> from = 0.00273785078713210000 * number
                    2 // "hr"
                    -> from = 0.00011407711613050400 * number
                    3 // "mn"
                    -> from = 0.00000190128526884174 * number
                    4 // "sec"
                    -> from = 0.00000003168808781403 * number
                }
                when (j) {
                    0    // "yr"
                    -> result = from
                    1 // "day"
                    -> result = 365.250000000000 * from
                    2 // "hr"
                    -> result = 8766.000000000000 * from
                    3 // "mn"
                    -> result = 525960.000000000000 * from
                    4 // "sec"
                    -> result = 31557600.000000000000 * from
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // PRESSURE conversion
        if (j == -1 && (i = findUnits(fromUnits, pressureUnits)) >= 0) {
            j = findUnits(toUnits, pressureUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"Pa", "atm", "mmHg"
                when (i) {
                    0    // "Pa"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes)
                    1 // "atm"
                    -> from = 101324.99658300000000000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    2 // "mmHg"
                    -> from = 133.32236392500000000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                }
                when (j) {
                    0    // "Pa"
                    -> result = from / prefixMultiplier(toPrefix, metricPrefixes)
                    1 // "atm"
                    -> result = 0.00000986923299998193 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    2 // "mmHg"
                    -> result = 0.00750061707998627000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // FORCE conversion
        if (j == -1 && (i = findUnits(fromUnits, forceUnits)) >= 0) {
            j = findUnits(toUnits, forceUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"N", "dyn", "lbf"
                when (i) {
                    0    // "N"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes)
                    1 // "dyn"
                    -> from = 0.000010000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    2 // "lbf"
                    -> from = 4.448222000000 * number
                }
                when (j) {
                    0    // "N"
                    -> result = from / prefixMultiplier(toPrefix, metricPrefixes)
                    1 // "dyn"
                    -> result = 100000.000000000000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    2 // "lbf"
                    -> result = 0.22480892365533900000 * from
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // ENERGY conversion
        if (j == -1 && (i = findUnits(fromUnits, energyUnits)) >= 0) {
            j = findUnits(toUnits, energyUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"J", "e", "c", "cal", "eV", "HPh", "Wh", "flb", "BTU"
                when (i) {
                    0 // "J"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes)
                    1 // "e"
                    -> from = 0.00000010000004806570 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    2 // "c"
                    -> from = 4.18399101363672000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    3 // "cal"
                    -> from = 4.18679484613929000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    4 // "eV"
                    -> from = 1.60217646E-19 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    5    // "HPh"
                    -> from = 2684517.41316170000000000000 * number
                    6    // "Wh"
                    -> from = 3599.99820554720000000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                    7 // "flb"
                    -> from = 0.04214000032364240000 * number
                    8 // "BTU"
                    -> from = 1055.05813786749000000000 * number
                }
                when (j) {
                    0 // "J"
                    -> result = from / prefixMultiplier(toPrefix, metricPrefixes)
                    1 // "e"
                    -> result = 9999995.19343231000000000000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    2 // "c"
                    -> result = 0.23900624947346700000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    3 // "cal"
                    -> result = 0.23884619064201700000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    4 // "eV"
                    -> result = 6241457000000000000.00000000000000000000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    5    // "HPh"
                    -> result = 0.00000037250643080100 * from
                    6    // "Wh"
                    -> result = 0.00027777791623871100 * from / prefixMultiplier(toPrefix, metricPrefixes)
                    7 // "flb"
                    -> result = 23.73042221926510000000 * from
                    8 // "BTU"
                    -> result = 0.00094781506734901500 * from
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // POWER conversion
        if (j == -1 && (i = findUnits(fromUnits, powerUnits)) >= 0) {
            j = findUnits(toUnits, powerUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"HP", "W"
                when (i) {
                    0 // "HP"
                    -> from = number
                    1 // "W"
                    -> from = 0.00134102006031908000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                }
                when (j) {
                    0 // "HP"
                    -> result = from
                    1 // "W"
                    -> result = 745.70100000000000000000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // MAGNETISM conversion
        if (j == -1 && (i = findUnits(fromUnits, magnetismUnits)) >= 0) {
            j = findUnits(toUnits, magnetismUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"T", "ga"
                when (i) {
                    0 // "T"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes)
                    1 // "ga"
                    -> from = 0.000100000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                }
                when (j) {
                    0 // "T"
                    -> result = from / prefixMultiplier(toPrefix, metricPrefixes)
                    1 // "ga"
                    -> result = 10000.000000000000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // TEMPERATURE conversion
        if (j == -1 && (i = findUnits(fromUnits, temperatureUnits)) >= 0) {
            j = findUnits(toUnits, temperatureUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"C", "F", "K"
                when (i) {
                    0 // "C"
                    -> from = number
                    1 // "F"
                    -> from = (number - 32) / 1.8
                    2 // "K"
                    -> from = number * prefixMultiplier(fromPrefix, metricPrefixes) - 273.15
                }
                when (j) {
                    0 // "C"
                    -> result = from
                    1 // "F"
                    -> result = from * 1.8 + 32
                    2 // "K"
                    -> result = (273.15000000000000000000 + from) / prefixMultiplier(toPrefix, metricPrefixes)
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        // LIQUID MEASURE conversion
        if (j == -1 && (i = findUnits(fromUnits, liquidMeasureUnits)) >= 0) {
            j = findUnits(toUnits, liquidMeasureUnits)
            if (j > -1) { // both fromUnits and toUnits are in same family
                //"tsp", "tbs", "oz", "cup", "pt", "uk_pt", "qt", "gal", "l"
                when (i) {
                    0    // "tsp"
                    -> from = number
                    1    // "tbs"
                    -> from = 3.00000000000000000000 * number
                    2  // "oz"
                    -> from = 6.00000000000000000000 * number
                    3 // "cup"
                    -> from = 48.00000000000000000000 * number
                    4 // "pt"
                    -> from = 96.00000000000000000000 * number
                    5    // "uk_pt"
                    -> from = 115.26600000000000000000 * number
                    6    // "qt"
                    -> from = 192.00000000000000000000 * number
                    7    // "gal"
                    -> from = 768.00000000000000000000 * number
                    8 // "l"
                    -> from = 202.84000000000000000000 * number * prefixMultiplier(fromPrefix, metricPrefixes)
                }
                when (j) {
                    0    // "tsp"
                    -> result = from
                    1    // "tbs"
                    -> result = 0.33333333333333300000 * from
                    2  // "oz"
                    -> result = 0.16666666666666700000 * from
                    3 // "cup"
                    -> result = 0.02083333333333330000 * from
                    4 // "pt"
                    -> result = 0.01041666666666670000 * from
                    5    // "uk_pt"
                    -> result = 0.00867558516821960000 * from
                    6    // "qt"
                    -> result = 0.00520833333333333000 * from
                    7    // "gal"
                    -> result = 0.00130208333333333000 * from
                    8 // "l"
                    -> result = 0.00492999408400710000 * from / prefixMultiplier(toPrefix, metricPrefixes)
                }
            } else
                return PtgErr(PtgErr.ERROR_NA)
        }

        val pnum = PtgNumber(result)
        if (DEBUG) Logger.logInfo("Result from CONVERT= " + pnum.string)
        return pnum
    }

    /**
     * DEC2BIN
     * Converts a decimal number to binary
     */
    internal fun calcDec2Bin(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "DEC2BIN")
        val dec = operands[0].intVal
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        if (dec < -512 || dec > 511 || places < 0) return PtgErr(PtgErr.ERROR_NUM)

        var bString = Integer.toBinaryString(dec)
        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            bString = bString.substring(Math.max(bString.length - 10, 0))
        } else if (places > 0) {
            if (bString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            bString = "0000000000$bString"    // maximum= 10 bits
            bString = bString.substring(bString.length - places)
        }
        val pstr = PtgStr(bString)
        if (DEBUG) Logger.logInfo("Result from DEC2BIN= " + pstr.string)
        return pstr
    }

    /**
     * DEC2HEX
     * Converts a decimal number to hexadecimal
     */
    internal fun calcDec2Hex(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "DEC2HEX")
        val dec = PtgCalculator.getLongValue(operands[0])
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal
        if (dec < -549755813888L || dec > 549755813887L || places < 0) return PtgErr(PtgErr.ERROR_NUM)
        var hString = java.lang.Long.toHexString(dec).toUpperCase()
        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            hString = hString.substring(Math.max(hString.length - 10, 0))
        } else if (places > 0) {
            if (hString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            hString = "0000000000$hString"    // maximum= 10 places
            hString = hString.substring(hString.length - places)

        }
        val pstr = PtgStr(hString)
        if (DEBUG) Logger.logInfo("Result from DEC2HEX= " + pstr.string)
        return pstr
    }

    /**
     * DEC2OCT
     * Converts a decimal number to octal
     */
    internal fun calcDec2Oct(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "DEC2OCT")
        val dec = PtgCalculator.getLongValue(operands[0])
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal
        if (dec < -536870912L || dec > 536870911L || places < 0) return PtgErr(PtgErr.ERROR_NUM)
        var oString = java.lang.Long.toOctalString(dec)
        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            oString = oString.substring(Math.max(oString.length - 10, 0))
        } else if (places > 0) {
            if (oString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            oString = "0000000000$oString"    // maximum= 10 places
            oString = oString.substring(oString.length - places)

        }
        val pstr = PtgStr(oString)
        if (DEBUG) Logger.logInfo("Result from DEC2OCT= " + pstr.string)
        return pstr
    }

    /**
     * DELTA
     * Tests whether two values are equal
     */
    internal fun calcDelta(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "DELTA")

        val number1 = operands[0].doubleVal
        var number2 = 0.0
        if (operands.size > 1)
            number2 = operands[1].doubleVal
        var result = 0
        if (number1 == number2)
            result = 1

        val pnum = PtgNumber(result.toDouble())
        if (DEBUG) Logger.logInfo("Result from DELTA= " + pnum.string)
        return pnum
    }


    /**
     * helper erf calc - seems to work *ALRIGHT* for values over 1
     * NOTE: not accurate to 9 digits for every case
     *
     * @param x
     * @return
     */
    private fun erf_try1(x: Double): Double {
        var t = x
        val x2 = Math.pow(x, 2.0)
        var n = 1000.0
        while (n >= 0.5) {
            t = x + n / t
            n -= 0.5
        }
        t = 1.0 / t
        val tt = Math.exp(-x2) / Math.sqrt(Math.PI)
        return 1 - tt * t
    }

    /**
     * ERF
     * Returns the error function integrated between lower_limit and upper_limit.
     * ERF(lower_limit,upper_limit)
     *
     *
     * Lower_limit     is the lower bound for integrating ERF.
     * Upper_limit     is the upper bound for integrating ERF. If omitted, ERF integrates between zero and lower_limit.
     *
     *
     * With a single argument ERF returns the error function, defined as
     * erf(x) = 2/sqrt(pi)* integral from 0 to x of exp(-t*t) dt.
     * If two arguments are supplied, they are the lower and upper limits of the integral.
     * NOTE: Accuracy is not always to 9 digits
     * NOTE: Version with two parameters is NOT supported
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun calcErf(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        try {
            var lower_limit = operands[0].doubleVal
            var upper_limit = java.lang.Double.NaN    //lower_limit;
            if (operands.size == 2)
                upper_limit = operands[1].doubleVal
            // If lower_limit is negative, ERF returns the #NUM! error value.
            // If upper_limit is negative, ERF returns the #NUM! error value.
            //			if (lower_limit < 0 /*|| upper_limitupper_limit < 0*/) return new PtgErr(PtgErr.ERROR_NUM);
            val neg = lower_limit < 0
            lower_limit = Math.abs(lower_limit)


            var result: Double
            val limit = lower_limit
            /*
			// try this: from "Computation of the error function erf in arbitrary precision with correct rounding"
			double r= 0;
			double r1= 0;
			double estimate= (2/Math.sqrt(Math.PI))*(limit - Math.pow(limit, 3)/3.0);
			double convergence= Math.pow(2, estimate-15);
			for (int i= 0, n= 0; n < 100; i++, n++) {
				double factor= 2.0*n + 1.0;
				double z= Math.pow(limit, factor);
				double zz= (MathFunctionCalculator.factorial(n)*factor);
				double zzz= z/zz;
				if ((i % 2)!=0)
					r= r-zzz;
				else
					r= r+zzz;

				if (Math.abs(r)-r1) <

				r1= Math.abs(r1);
			}
			result= r*(2.0/Math.sqrt(Math.PI));*/

            if (limit < 0.005) {
                /* A&S 7.1.1 - good to at least 6 digts ... but not for larger values ... sigh ...*/
                var r = 0.0
                var i = 0
                var n = 0
                while (n < 12) {
                    val factor = 2.0 * n + 1.0
                    val z = Math.pow(limit, factor)
                    val zz = MathFunctionCalculator.factorial(n.toLong()) * factor
                    val zzz = z / zz
                    if (i % 2 != 0)
                        r = r - zzz
                    else
                        r = r + zzz
                    i++
                    n++
                }
                result = r * (2.0 / Math.sqrt(Math.PI))
            } else
                result = erf_try1(limit)

            if (neg)
                result *= -1.0

            if (!java.lang.Double.isNaN(upper_limit)) {    // Erf(upper)-Erf(lower)
                val result2 = calcErf(arrayOf(operands[1]))
                if (result2 is PtgNumber)
                    result = result2.doubleVal - result
                else
                    return PtgErr(PtgErr.ERROR_VALUE)
            }

            return PtgNumber(result)

        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    /***
     * Calculate the error function erf.
     * @param x the argument
     * @return the value erf(x)
     */
    fun erf(x: Double): Double {
        //if (type==1) {
        return erfAS(x)
        //}
        //return erfCody(x);
    }

    /***
     * Calculate the remaining error function erfc.
     * @param x the argument
     * @return the value erfc(x)
     */
    fun erfc(x: Double): Double {
        //if (type==1) {
        return erfcAS(x)
        //}
        //return erfcCody(x);
    }

    /***
     * Internal helper method to calculate the error function at value x.
     * This code is  based on a Fortran implementation from
     * [W. J. Cody](http://www.netlib.org/specfun/erf).
     * Refactored by N.Wulff for Java.
     *
     * Cody vs AS algiorithm
     *
     * @param x the argument
     * @return the approximation of erf(x)
     */
    private fun erfCody(x: Double): Double {
        var result = 0.0
        val y = Math.abs(x)
        if (firstCall) {
            firstCall = false
        }
        if (y <= THRESHOLD) {
            result = x * calcLower(y)
        } else {
            result = calcUpper(y)
            result = 0.5 - result + 0.5
            if (x < 0) {
                result = -result
            }
        }
        return result
    }

    /***
     * Internal helper method to calculate the erfc functions.
     * This code is  based on a Fortran implementation from
     * [W. J. Cody](http://www.netlib.org/specfun/erf).
     * Refactored by N.Wulff for Java.
     *
     * @param x the argument
     * @return the approximation erfc(x)
     */
    private fun erfcCody(x: Double): Double {
        var result = 0.0
        val y = Math.abs(x)
        if (firstCall) {
            firstCall = false
        }
        if (y <= THRESHOLD) {
            result = x * calcLower(y)
            result = 1 - result
        } else {
            result = calcUpper(y)
            if (x < 0) {
                result = 2.0 - result
            }
        }
        return result
    }

    /***
     * Internal helper method to calculate the erf/erfc functions.
     * This code is  based on a Fortran implementation from
     * [W. J. Cody](http://www.netlib.org/specfun/erf).
     * Refactored by N.Wulff for Java.
     *
     * @param y the value y=abs(x)<=THRESHOLD
     * @return the series expansion
     */
    private fun calcLower(y: Double): Double {
        val result: Double
        var ySq: Double
        var xNum: Double
        var xDen: Double
        ySq = 0.0
        if (y > X_SMALL)
            ySq = y * y
        xNum = ERF_A[4] * ySq
        xDen = ySq
        for (i in 0..2) {
            xNum = (xNum + ERF_A[i]) * ySq
            xDen = (xDen + ERF_B[i]) * ySq
        }
        result = (xNum + ERF_A[3]) / (xDen + ERF_B[3])
        return result
    }

    /***
     * Internal helper method to calculate the erf/erfc functions.
     * This code is  based on a Fortran implementation from
     * [W. J. Cody](http://www.netlib.org/specfun/erf).
     * Refactored by N.Wulff for Java.
     *
     * @param y the value y=abs(x)>THRESHOLD
     * @return the series expansion
     */
    private fun calcUpper(y: Double): Double {
        var result: Double
        var ySq: Double
        var xNum: Double
        var xDen: Double
        if (y <= 4.0) {
            xNum = ERF_C[8] * y
            xDen = y
            for (i in 0..6) {
                xNum = (xNum + ERF_C[i]) * y
                xDen = (xDen + ERF_D[i]) * y
            }
            result = (xNum + ERF_C[7]) / (xDen + ERF_D[7])
        } else {
            result = 0.0
            if (y >= X_HUGE) {
                result = SQRPI / y
            } else {
                ySq = 1.0 / (y * y)
                xNum = ERF_P[5] * ySq
                xDen = ySq
                for (i in 0..3) {
                    xNum = (xNum + ERF_P[i]) * ySq
                    xDen = (xDen + ERF_Q[i]) * ySq
                }
                result = ySq * (xNum + ERF_P[4]) / (xDen + ERF_Q[4])
                result = (SQRPI - result) / y
            }
        }
        ySq = Math.round(y * 16.0) / 16.0
        val del = (y - ySq) * (y + ySq)
        result = Math.exp(-ySq * ySq) * Math.exp(-del) * result
        return result
    }

    /***
     * Calculate the error function at value x.
     * AS 7.1.5/7.1.26
     * @param x the argument
     * @return the value erf(x)
     */
    private fun erfAS(x: Double): Double {
        if (firstCall) {
            firstCall = false
        }
        if (x < 0) {
            return -erfAS(-x)
        }
        return if (x < 2) {
            erfSeries(x)
        } else erfRational(x)
    }

    /***
     * Calculate the remaining erfc error function at value x.
     * AS 7.1.5/7.1.26
     * @param x the argument
     * @return the value erfc(x)
     */
    private fun erfcAS(x: Double): Double {
        if (firstCall) {
            firstCall = false
        }
        return 1 - erfAS(x)
    }

    /***
     * Series expansion from A&S 7.1.5.
     *
     * @param x the argument
     * @return erf(x)
     */
    private fun erfSeries(x: Double): Double {
        val eps = 1e-8 // we want only ~1.E-7
        val kmax = 50 // this can be reached with ~30-40
        var an: Double
        var ak = x
        var erfo: Double
        var erf = ak
        var k = 1
        do {
            erfo = erf
            ak *= -x * x / k
            an = ak / (2.0 * k + 1.0)
            erf += an
        } while (!hasConverged(erf, erfo, eps, ++k, kmax))
        return TSQPI * erf
    }

    /**
     * Indicate if an iterative algorithm has RELATIVE converged.
     * <hr></hr>
     * **Note**:
     * HasConverged throws an ArithmeticException if more than max calls
     * have been made. Choose hasReacherAccuracy if this is not desired.
     * <hr></hr>
     *
     * @param xn  the actual argument x[n]
     * @param xo  the older argument x[n-1]
     * @param eps the accuracy to reach
     * @param n   the actual iteration counter
     * @param max the maximal number of iterations
     * @return flag indicating if accuracy is reached.
     */
    fun hasConverged(xn: Double, xo: Double,
                     eps: Double, n: Int, max: Int): Boolean {
        if (hasReachedAccuracy(xn, xo, eps)) {
            return true
        }
        if (n >= max) {
            throw ArithmeticException()
        }
        return false
    }

    /**
     * Indicate if xn and xo have the relative/absolute accuracy epsilon.
     * In case that the true value is less than one this is based
     * on the absolute difference, otherwise on the relative difference:
     * <pre>
     * 2*|x[n]-x[n-1]|/|x[n]+x[n-1]| < eps
    </pre> *
     *
     * @param xn  the actual argument x[n]
     * @param xo  the older argument x[n-1]
     * @param eps accuracy to reach
     * @return flag indicating if accuracy is reached.
     */
    fun hasReachedAccuracy(xn: Double, xo: Double,
                           eps: Double): Boolean {
        val z = Math.abs(xn + xo) / 2
        var error = Math.abs(xn - xo)
        if (z > 1) {
            error /= z
        }
        return error <= eps
    }

    /***
     * Rational approximation A&S 7.1.26 with accuracy 1.5E-7.
     *
     * @param x the argument
     * @return erf(x)
     */
    private fun erfRational(x: Double): Double {
        /*  coefficients for A&S 7.1.26. */
        val a = doubleArrayOf(.254829592, -.284496736, 1.421413741, -1.453152027, 1.061405429)
        /*  constant for A&S 7.1.26 */
        val p = .3275911
        val erf: Double
        var r = 0.0
        val t = 1.0 / (1 + p * x)
        for (i in 4 downTo 0) {
            r = a[i] + r * t
        }
        erf = 1 - t * r * Math.exp(-x * x)
        return erf
    }
    /*
	ERFC
	 Returns the complementary error function
*/

    /**
     * GESTEP
     * Tests whether a number is greater than a threshold value
     */
    internal fun calcGEStep(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "GESTEP")

        val number = operands[0].doubleVal
        var step = 0.0
        if (operands.size > 1)
            step = operands[1].doubleVal
        var result = 0
        if (number >= step)
            result = 1

        val pnum = PtgNumber(result.toDouble())
        if (DEBUG) Logger.logInfo("Result from GESTEP= " + pnum.string)
        return pnum
    }

    /**
     * HEX2BIN
     * Converts a hexadecimal number to binary
     */
    internal fun calcHex2Bin(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "HEX2BIN")
        val hString = operands[0].string.trim { it <= ' ' }
        // 10 digits (40 bits) at most
        if (hString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        var dec: Long
        var bString: String
        try {
            dec = java.lang.Long.parseLong(hString, 16)
            // must det. manually if binary string is negative because parseInt/parseLong does not
            // handle two's complement input!!!
            if (dec >= 549755813888L)
            // 2^39 (= signed 40 bits)
                dec -= 1099511627776L        // 2^40 (= signed 41 bits)
            bString = java.lang.Long.toBinaryString(dec)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        if (dec < -512 /*0xFFFFFFFE00*/ || dec > 0x1FF || places < 0) return PtgErr(PtgErr.ERROR_NUM)
        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            bString = bString.substring(Math.max(bString.length - 10, 0))
        } else if (places > 0) {
            if (bString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            bString = "0000000000$bString"    // maximum= 10 places
            bString = bString.substring(bString.length - places)
        }
        val pstr = PtgStr(bString)
        if (DEBUG) Logger.logInfo("Result from HEX2BIN= " + pstr.string)
        return pstr
    }

    /**
     * HEX2DEC
     * Converts a hexadecimal number to decimal
     */
    internal fun calcHex2Dec(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "HEX2DEC")
        val hString = operands[0].string.trim { it <= ' ' }
        // 10 digits (40 bits) at most
        if (hString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        var dec: Long
        val oString: String
        try {
            dec = java.lang.Long.parseLong(hString, 16)
            // must det. manually if binary string is negative because parseInt/parseLong does not
            // handle two's complement input!!!
            if (dec >= 549755813888L)
            // 2^39 (= signed 40 bits)
                dec -= 1099511627776L        // 2^40 (= signed 41 bits)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val pnum = PtgNumber(dec.toDouble())
        if (DEBUG) Logger.logInfo("Result from HEX2DEC= " + pnum.`val`)
        return pnum
    }

    /**
     * HEX2OCT
     * Converts a hexadecimal number to octal
     */
    internal fun calcHex2Oct(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "HEX2OCT")
        val hString = operands[0].string.trim { it <= ' ' }
        // 10 digits (40 bits) at most
        if (hString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        var dec: Long
        var oString: String
        try {
            dec = java.lang.Long.parseLong(hString, 16)
            // must det. manually if binary string is negative because parseInt/parseLong does not
            // handle two's complement input!!!
            if (dec >= 549755813888L)
            // 2^39 (= signed 40 bits)
                dec -= 1099511627776L        // 2^40 (= signed 41 bits)
            oString = java.lang.Long.toOctalString(dec)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        if (dec < -536870912L /*0xFFE0000000*/ || dec > 0x1FFFFFFF || places < 0) return PtgErr(PtgErr.ERROR_NUM)
        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            oString = oString.substring(Math.max(oString.length - 10, 0))
        } else if (places > 0) {
            if (oString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            oString = "0000000000$oString"    // maximum= 10 places
            oString = oString.substring(oString.length - places)
        }
        val pstr = PtgStr(oString)
        if (DEBUG) Logger.logInfo("Result from HEX2OCT= " + pstr.string)
        return pstr
    }

    /**
     * imParseComplexNumber
     *
     *
     * used in the following Imaginary-based formulas to parse a complex number into its
     * real and imaginary components.
     * Throws a numberformat exception if the complex number is not in the format of:
     * real + imaginary
     * real - imaginary
     * imaginary			(real= 0)
     * real				(imaginary= 0)
     * where imaginary coefficient n= ni or nj
     * *
     */
    @Throws(NumberFormatException::class)
    private fun imParseComplexNumber(complexNumber: String): Complex {
        var complexNumber = complexNumber
        val c = Complex()
        if (complexNumber.length > 0) {
            try {
                var i = complexNumber.length
                if (complexNumber.substring(i - 1, i) == "i" || complexNumber.substring(i - 1, i) == "j") {
                    c.suffix = complexNumber.substring(i - 1, i)
                    i -= 2
                    while (i >= 0 && !(complexNumber.substring(i, i + 1) == "+" || complexNumber.substring(i, i + 1) == "-"))
                        i--
                    if (i < 0) { // case of "#i" or "#j" i.e. no real and no sign
                        complexNumber = "+$complexNumber"
                        i++
                    }
                    // get imaginary coefficient + sign
                    var s = complexNumber.substring(i, complexNumber.length - 1)
                    if (s.length == 1) {    // only a sign; means that the coefficient==1 eg. real-j or real+i
                        s += "1"
                    }
                    c.imaginary = java.lang.Double.parseDouble(s)
                }
                if (i > 0)
                    c.real = java.lang.Double.parseDouble(complexNumber.substring(0, i))
            } catch (e: Exception) {
                throw NumberFormatException()
            }

        }
        return c
    }

    /**
     * imGetStr
     *
     * @param operands
     * @return double formatted as an integer if no precision, otherwise rounds to 15
     */
    private fun imGetExcelStr(d: Double, precision: Int = 15): String {
        var d = d
        val s: String
        if (d == d.toInt().toDouble()) {
            return if (d.toInt() == 1) "" else d.toInt().toString()
        }
        // round to precision - default= 15
        val r = Math.pow(10.0, precision.toDouble())
        d *= r
        d = Math.round(d).toDouble()
        d /= r
        return d.toString()
    }

    /**
     * IMABS
     * Returns the absolute value (modulus) of a complex number
     */
    internal fun calcImAbs(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "IMABS")
        val complexString = StringTool.allTrim(operands[0].string)

        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // Absolute of a complex number is:
        // square root( real^2 + imaginary^2)
        val result = Math.sqrt(Math.pow(c.real, 2.0) + Math.pow(c.imaginary, 2.0))

        val pnum = PtgNumber(result)
        if (DEBUG) Logger.logInfo("Result from IMABS= " + pnum.string)
        return pnum
    }

    /**
     * IMAGINARY
     * Returns the imaginary coefficient of a complex number
     */
    internal fun calcImaginary(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "Imaginary")
        val complexString = StringTool.allTrim(operands[0].string)

        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val pnum = PtgNumber(c.imaginary)
        if (DEBUG) Logger.logInfo("Result from IMAGINARY= " + pnum.string)
        return pnum
    }

    /**
     * IMARGUMENT
     * Returns the argument theta, an angle expressed in radians
     */
    internal fun calcImArgument(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImArgument")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val result = Math.atan(c.imaginary / c.real)

        val pnum = PtgNumber(result)
        if (DEBUG) Logger.logInfo("Result from IMARGUMENT= " + pnum.string)
        return pnum
    }

    /**
     * IMCONJUGATE
     * Returns the complex conjugate of a complex number
     */
    internal fun calcImConjugate(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImCongugate")
        val complexString = StringTool.allTrim(operands[0].string)

        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val congugate: String
        if (c.real != 0.0 && c.imaginary != 0.0)
            congugate = imGetExcelStr(c.real) + (if (c.imaginary < 0) "+" else "-") + imGetExcelStr(Math.abs(c.imaginary)) + c.suffix
        else if (c.real == 0.0)
            congugate = imGetExcelStr(Math.abs(c.imaginary)) + c.suffix
        else
            congugate = imGetExcelStr(c.real)
        val pstr = PtgStr(congugate)
        if (DEBUG) Logger.logInfo("Result from IMCONGUGATE= " + pstr.string)
        return pstr
    }

    /**
     * IMCOS
     * Returns the cosine of a complex number
     */
    internal fun calcImCos(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImCos")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // cos(a + bi)= cos(a)cosh(b) - sin(a)sinh(b)i
        val cosh = (Math.pow(Math.E, c.imaginary) + Math.pow(Math.E, c.imaginary * -1)) / 2
        val sinh = (Math.pow(Math.E, c.imaginary) - Math.pow(Math.E, c.imaginary * -1)) / 2
        val a = Math.cos(c.real) * cosh
        val b = Math.sin(c.real) * sinh

        val imCos: String
        if (b < 0)
            imCos = imGetExcelStr(a) + "+" + imGetExcelStr(Math.abs(b)) + c.suffix
        else
            imCos = imGetExcelStr(a) + "-" + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imCos)
        if (DEBUG) Logger.logInfo("Result from IMCOS= " + pstr.string)
        return pstr
    }

    /**
     * IMDIV
     * Returns the quotient of two complex numbers, defined as (deep breath):
     *
     *
     * (r1 + i1j)/(r2 + i2j) == ( r1r2 + i1i2 + (r2i1 - r1i2)i ) / (r2^2 + i2^2)
     */
    internal fun calcImDiv(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "IMDIV")
        val complexString1 = StringTool.allTrim(operands[0].string)
        val complexString2 = StringTool.allTrim(operands[1].string)

        val c1: Complex
        val c2: Complex

        try {
            c1 = imParseComplexNumber(complexString1)
            c2 = imParseComplexNumber(complexString2)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val divisor = Math.pow(c2.real, 2.0) + Math.pow(c2.imaginary, 2.0)
        val a = c1.real * c2.real + c1.imaginary * c2.imaginary
        val b = c1.imaginary * c2.real - c1.real * c2.imaginary
        val c = a / divisor
        val d = b / divisor

        val imDiv: String
        if (d > 0)
            imDiv = imGetExcelStr(c) + "+" + imGetExcelStr(d) + "i"
        else
            imDiv = imGetExcelStr(c) + "-" + imGetExcelStr(d) + "i"

        val pstr = PtgStr(imDiv)
        if (DEBUG) Logger.logInfo("Result from IMDIV= " + pstr.string)
        return pstr
    }

    /**
     * IMEXP
     * Returns the exponential of a complex number
     */
    internal fun calcImExp(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImExp")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // Exponential of a complex number x+yi = e^x(cos(y) + sin(y)i)
        val e_x = Math.pow(Math.E, c.real)
        val a = e_x * Math.cos(c.imaginary)
        val b = e_x * Math.sin(c.imaginary)

        val imExp: String
        if (b > 0)
            imExp = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c.suffix
        else
            imExp = imGetExcelStr(a) + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imExp)
        if (DEBUG) Logger.logInfo("Result from IMEXP= " + pstr.string)
        return pstr
    }

    /**
     * IMLN
     * Returns the natural logarithm of a complex number
     */
    internal fun calcImLn(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImLn")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // Natural log of a complex number=
        // 	IMLN(x + yi)= ln(sqrt(x2+y2)) + atan(y/x)i
        var a = Math.pow(c.real, 2.0) + Math.pow(c.imaginary, 2.0)
        a = Math.sqrt(a)
        a = Math.log(a)
        val b = Math.atan(c.imaginary / c.real)

        val imLn: String
        if (b > 0)
            imLn = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c.suffix
        else
            imLn = imGetExcelStr(a) + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imLn)
        if (DEBUG) Logger.logInfo("Result from IMLN= " + pstr.string)
        return pstr
    }

    /**
     * IMLOG10
     * Returns the base-10 logarithm of a complex number
     */
    internal fun calcImLog10(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImLog10")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // Natural log of a complex number=
        // 	IMLN(x + yi)= ln(sqrt(x2+y2)) + atan(y/x)i
        var a = Math.pow(c.real, 2.0) + Math.pow(c.imaginary, 2.0)
        a = Math.sqrt(a)
        a = Math.log(a)
        var b = Math.atan(c.imaginary / c.real)
        // now, convert to base 10 log:
        val logE = Math.log(Math.E) / Math.log(10.0)
        a = a * logE
        b = b * logE

        val imLog10: String
        if (b > 0)
            imLog10 = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c.suffix
        else
            imLog10 = imGetExcelStr(a) + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imLog10)
        if (DEBUG) Logger.logInfo("Result from IMLOG10= " + pstr.string)
        return pstr
    }

    /**
     * IMLOG2
     * Returns the base-2 logarithm of a complex number
     */
    internal fun calcImLog2(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImLog2")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // Natural log of a complex number=
        // 	IMLN(x + yi)= ln(sqrt(x2+y2)) + atan(y/x)i
        var a = Math.pow(c.real, 2.0) + Math.pow(c.imaginary, 2.0)
        a = Math.sqrt(a)
        a = Math.log(a)
        var b = Math.atan(c.imaginary / c.real)
        // now, convert to base 2 log:
        val logE = Math.log(Math.E) / Math.log(2.0)
        a = a * logE
        b = b * logE
        // TODO: Results only correct to 8th precision: WHY???
        val imLog2: String
        if (b > 0)
            imLog2 = imGetExcelStr(a, 8) + "+" + imGetExcelStr(b, 8) + c.suffix
        else
            imLog2 = imGetExcelStr(a, 8) + imGetExcelStr(b, 8) + c.suffix

        val pstr = PtgStr(imLog2)
        if (DEBUG) Logger.logInfo("Result from IMLOG2= " + pstr.string)
        return pstr
    }

    /**
     * IMPOWER
     * Returns a complex number raised to any power
     */
    internal fun calcImPower(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImPower")
        val complexString = StringTool.allTrim(operands[0].string)
        val n = operands[1].doubleVal
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // A complex number (x + yi) raised to a power n is:
        // sqrt(x^2 + y^2)*cos(n*atan(y/x)) + sqrt(x^2 + y^2)*sin(n*atan(y/x))i
        var r = Math.pow(c.real, 2.0) + Math.pow(c.imaginary, 2.0)
        r = Math.sqrt(r)
        r = Math.pow(r, n)
        val t = Math.atan(c.imaginary / c.real)
        val a = r * Math.cos(n * t)
        val b = r * Math.sin(n * t)

        val imPower: String
        if (b > 0)
            imPower = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c.suffix
        else
            imPower = imGetExcelStr(a) + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imPower)
        if (DEBUG) Logger.logInfo("Result from IMPOWER= " + pstr.string)
        return pstr
    }

    /**
     * IMPRODUCT
     * Returns the product of two complex numbers
     */
    internal fun calcImProduct(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }

        if (DEBUG) debugOperands(operands, "IMPRODUCT")
        val ops = PtgCalculator.getAllComponents(operands)
        val complexStrings = arrayOfNulls<String>(ops.size)
        for (i in ops.indices)
            complexStrings[i] = StringTool.allTrim(ops[i].string)

        val c = arrayOfNulls<Complex>(complexStrings.size)
        for (i in complexStrings.indices) {
            try {
                c[i] = imParseComplexNumber(complexStrings[i])
            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_NUM)
            }

        }

        // basically, linear binomial multiplication over n terms
        // (a + bi)(c + di) = (ac-bd) + (ad+bc)i   for n terms
        for (i in 1 until c.size) {
            val a = c[0].real
            val b = c[0].imaginary
            c[0].real = a * c[i].real - b * c[i].imaginary
            c[0].imaginary = a * c[i].imaginary + b * c[i].real
        }

        // Format Result
        val imSum: String
        if (c[0].imaginary > 0)
            imSum = imGetExcelStr(c[0].real) + "+" + imGetExcelStr(c[0].imaginary) + c[0].suffix
        else
            imSum = imGetExcelStr(c[0].real) + imGetExcelStr(c[0].imaginary) + c[0].suffix

        val pstr = PtgStr(imSum)
        if (DEBUG) Logger.logInfo("Result from IMSPRODUCT= " + pstr.string)
        return pstr
    }

    /**
     * IMREAL
     * Returns the real coefficient of a complex number
     */
    internal fun calcImReal(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "IMREAL")
        val complexString = StringTool.allTrim(operands[0].string)

        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val pnum = PtgNumber(c.real)
        if (DEBUG) Logger.logInfo("Result from IMREAL= " + pnum.string)
        return pnum
    }

    /**
     * IMSIN
     * Returns the sine of a complex number
     */
    internal fun calcImSin(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImSin")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // sin(a + bi)= sin(a)cosh(b) + cos(a)sinh(b)i (Excel doc is wrong!)
        val cosh = (Math.pow(Math.E, c.imaginary) + Math.pow(Math.E, c.imaginary * -1)) / 2
        val sinh = (Math.pow(Math.E, c.imaginary) - Math.pow(Math.E, c.imaginary * -1)) / 2
        val a = Math.sin(c.real) * cosh
        val b = Math.cos(c.real) * sinh

        val imSin: String
        if (b > 0)
            imSin = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c.suffix
        else
            imSin = imGetExcelStr(a) + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imSin)
        if (DEBUG) Logger.logInfo("Result from IMSIN= " + pstr.string)
        return pstr
    }

    /**
     * IMSQRT
     * Returns the square root of a complex number
     */
    internal fun calcImSqrt(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ImSqrt")
        val complexString = StringTool.allTrim(operands[0].string)
        val c: Complex
        try {
            c = imParseComplexNumber(complexString)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // The square root of a complex number (x + yi) is:
        // sqrt(sqrt(x^2 + y^2))*cos(atan(y/x)/2) +
        //   sqrt(sqrt(x^2 + y^2))*sin(atan(y/x)/2)i
        var r = Math.pow(c.real, 2.0) + Math.pow(c.imaginary, 2.0)
        r = Math.sqrt(r)
        r = Math.sqrt(r)
        val t = Math.atan(c.imaginary / c.real)
        val a = r * Math.cos(t / 2)
        val b = r * Math.sin(t / 2)

        val imSqrt: String
        if (b > 0)
            imSqrt = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c.suffix
        else
            imSqrt = imGetExcelStr(a) + imGetExcelStr(b) + c.suffix

        val pstr = PtgStr(imSqrt)
        if (DEBUG) Logger.logInfo("Result from IMSQRT= " + pstr.string)
        return pstr
    }


    /**
     * IMSUB
     * Returns the difference of two complex numbers
     */
    internal fun calcImSub(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "IMSUB")
        val complexString1 = StringTool.allTrim(operands[0].string)
        val complexString2 = StringTool.allTrim(operands[1].string)

        val c1: Complex
        val c2: Complex
        try {
            c1 = imParseComplexNumber(complexString1)
            c2 = imParseComplexNumber(complexString2)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        // basically, linear binomial subtraction:
        // (a + bi) - (c + di)= (a-c) + (b-d)i
        val a = c1.real - c2.real
        val b = c1.imaginary - c2.imaginary

        val imSub: String
        if (b > 0)
            imSub = imGetExcelStr(a) + "+" + imGetExcelStr(b) + c1.suffix    // should have the same suffix
        else
            imSub = imGetExcelStr(a) + imGetExcelStr(b) + c1.suffix

        val pstr = PtgStr(imSub)
        if (DEBUG) Logger.logInfo("Result from IMSUB= " + pstr.string)
        return pstr
    }

    /**
     * IMSUM
     * Returns the sum of complex numbers
     */
    internal fun calcImSum(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }

        if (DEBUG) debugOperands(operands, "IMSUM")
        val ops = PtgCalculator.getAllComponents(operands)
        val complexStrings = arrayOfNulls<String>(ops.size)
        for (i in ops.indices)
            complexStrings[i] = StringTool.allTrim(ops[i].string)

        val c = arrayOfNulls<Complex>(complexStrings.size)
        for (i in complexStrings.indices) {
            try {
                c[i] = imParseComplexNumber(complexStrings[i])
            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_NUM)
            }

        }

        // basically, linear binomial addition over n terms
        // (a + bi)+(c + di) = (a+c) + (b+d)i   for n terms
        for (i in 1 until c.size) {
            c[0].real = c[0].real + c[i].real
            c[0].imaginary = c[0].imaginary + c[i].imaginary
        }

        // Format Result
        val imSum: String
        if (c[0].imaginary > 0)
            imSum = imGetExcelStr(c[0].real) + "+" + imGetExcelStr(c[0].imaginary) + c[0].suffix
        else
            imSum = imGetExcelStr(c[0].real) + imGetExcelStr(c[0].imaginary) + c[0].suffix

        val pstr = PtgStr(imSum)
        if (DEBUG) Logger.logInfo("Result from IMSUM= " + pstr.string)
        return pstr
    }

    /**
     * OCT2BIN
     * Converts an octal number to binary
     */
    internal fun calcOct2Bin(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "OCT2BIN")
        val l = operands[0].doubleVal.toLong()    // avoid sci notation
        val oString = l.toString().trim { it <= ' ' }
        // 10 digits at most (=30 bits)
        if (oString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        var dec: Long
        var bString: String
        try {
            dec = java.lang.Long.parseLong(oString, 8)
            // must det. manually if binary string is negative because parseInt/parseLong does not
            // handle two's complement input!!!
            if (dec >= 536870912L)
            // 2^29 (= 30 bits, signed)
                dec -= 1073741824L        // 2^30 (= 31 bits, signed)
            bString = java.lang.Long.toBinaryString(dec)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        if (dec < -512 /*7777777000*/ || dec > 511 || places < 0) return PtgErr(PtgErr.ERROR_NUM)
        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            bString = bString.substring(Math.max(bString.length - 10, 0))
        } else if (places > 0) {
            if (bString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            bString = "0000000000$bString"    // maximum= 10 places
            bString = bString.substring(bString.length - places)
        }
        val pstr = PtgStr(bString)
        if (DEBUG) Logger.logInfo("Result from OCT2BIN= " + pstr.string)
        return pstr
    }

    /**
     * OCT2DEC
     * Converts an octal number to decimal
     */
    internal fun calcOct2Dec(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "OCT2DEC")
        val l = operands[0].doubleVal.toLong() // avoid sci notation
        val oString = l.toString().trim { it <= ' ' }
        // 10 digits at most (=30 bits)
        if (oString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        var dec: Long
        try {
            dec = java.lang.Long.parseLong(oString, 8)
            // must det. manually if binary string is negative because parseInt/parseLong does not
            // handle two's complement input!!!
            if (dec >= 536870912L)
            // 2^29 (= 30 bits, signed)
                dec -= 1073741824L        // 2^30 (= 31 bits, signed)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        val pnum = PtgNumber(dec.toDouble())
        if (DEBUG) Logger.logInfo("Result from OCT2DEC= " + pnum.`val`)
        return pnum
    }

    /**
     * OCT2HEX
     * Converts an octal number to hexadecimal
     */
    internal fun calcOct2Hex(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "OCT2HEX")
        val l = operands[0].doubleVal.toLong() // avoid sci notation
        val oString = l.toString().trim { it <= ' ' }
        // 10 digits at most (=30 bits)
        if (oString.length > 10) return PtgErr(PtgErr.ERROR_NUM)
        var places = 0
        if (operands.size > 1)
            places = operands[1].intVal

        var dec: Long
        var hString: String
        try {
            dec = java.lang.Long.parseLong(oString, 8)
            // must det. manually if binary string is negative because parseInt/parseLong does not
            // handle two's complement input!!!
            if (dec >= 536870912L)
            // 2^29 (= 30 bits, signed)
                dec -= 1073741824L        // 2^30 (= 31 bits, signed)
            hString = java.lang.Long.toHexString(dec).toUpperCase()
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_NUM)
        }

        if (dec < 0) {    // truncate to 10 digits automatically (should already be two's complement)
            hString = hString.substring(Math.max(hString.length - 10, 0))
        } else if (places > 0) {
            if (hString.length > places) return PtgErr(PtgErr.ERROR_NUM)
            hString = "0000000000$hString"    // maximum= 10 bits
            hString = hString.substring(hString.length - places)
        }

        val pstr = PtgStr(hString)
        if (DEBUG) Logger.logInfo("Result from OCT2HEX= " + pstr.string)
        return pstr
    }

    /*
	SQRTPI
	 Returns the square root of (number * PI)
*/
    internal fun debugOperands(operands: Array<Ptg>, f: String) {
        if (DEBUG) {
            Logger.logInfo("Operands for $f")
            for (i in operands.indices) {
                val s = operands[i].string
                if (operands[i] !is PtgMissArg) {
                    val v = operands[i].value.toString()
                    Logger.logInfo("\tOperand[$i]=$s $v")
                } else
                    Logger.logInfo("\tOperand[$i]=$s is Missing")
            }
        }
    }

}

internal class Complex {
    var real: Double = 0.toDouble()
    var imaginary: Double = 0.toDouble()
    var suffix: String

    constructor() {
        this.suffix = "i"
    }

    constructor(r: Double, i: Double) {
        this.real = r
        this.imaginary = i
        this.suffix = "i"
    }
}	

