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
package io.starter.OpenXLS

import io.starter.toolkit.CompatibleBigDecimal
import io.starter.toolkit.Logger
import io.starter.toolkit.ResourceLoader
import io.starter.toolkit.StringTool

import java.util.ArrayList
import java.util.Date
import java.util.IllegalFormatConversionException
import java.util.StringTokenizer

//import java.text.SimpleDateFormat;


/**
 * OpenXLS helper methods. <br></br>
 * Contains helpful methods to ease use of the OpenXLS toolkit. <br></br>
 *
 *
 * "http://starter.io">Starter Inc.
 *
 * @see ByteTools
 */

class ExcelTools : java.io.Serializable {
    companion object {

        private const val serialVersionUID = 7622857355626065370L

        /**
         * Formats a double in the standard OpenXLS (General) format. Up to
         * 99999999999 is expressed in standard notation. Above that is formatted in
         * scientific notation
         *
         *
         * In addition, Excel precision of 9 digits is maintained
         *
         *
         *
         *
         * returns a number formatted in Excel's General format, (assuming a wide
         * enough column width - see below) example:
         * formatNumericNotation(1234567890123) returns "1.23457E+12"
         *
         *
         * Information on NOTATION_STANDARD_EXCEL (i.e. Excel's General Format): //
         * Excel will show as many decimal places that the text item has room for,
         * it won't use a thousands separator, and if the // number can't fit, Excel
         * uses a scientific number format. // RULES: // 1- Assuming the column is
         * wide enough numbers will only be displayed in the scientific format when
         * they contain more than 10 digits. // 2- If you enter a number into a cell
         * and thre is not enough room to display all the digits // then the number
         * will either be displayed in scientific format or will not be displayed at
         * all, meaning that ##### will appear. // The exact precision of the
         * scientific format will depend on the width of the actual cell.
         *
         * @param fpnum
         * @return String formatted number
         */
        fun getNumberAsString(fpnum: Double): String {
            // Ensure precision and number of digits ala Excel
            // double issues - use BigDecimal
            var bd = java.math.BigDecimal(fpnum)
            val scale = bd.scale()
            if (Math.abs(fpnum) > 0.000000001 && scale > 9) {
                bd = bd.setScale(9, java.math.RoundingMode.HALF_UP)
            } else if (scale > 9)
                bd = java.math.BigDecimal(fpnum, java.math.MathContext(5, java.math.RoundingMode.HALF_UP))
            bd = bd.stripTrailingZeros()
            var s = bd.toPlainString()
            val len = s.length
            // If larger than 11 characters, truncate string
            if (len > 11 && fpnum > 0 || len > 12) { // must deal with exponents and such as well
                if (scale == 0) {
                    s = java.math.BigDecimal(bd.toString(), java.math.MathContext(6, java.math.RoundingMode.HALF_UP)).toString()
                } else if (bd.toString().indexOf("E") == -1) {
                    s = java.math.BigDecimal(bd.toString(), java.math.MathContext(10, java.math.RoundingMode.HALF_UP)).toString()
                    while (s.length > 0 && s[s.length - 1] == '0')
                        s = s.substring(0, s.length - 1)
                    if (s.endsWith("."))
                        s = s.substring(0, s.length - 1)
                } else { // 5 + E+XX + sign
                    s = java.math.BigDecimal(bd.toString(), java.math.MathContext(5, java.math.RoundingMode.HALF_UP)).toString()
                }
            }
            return s
        }

        /**
         * static version of getFormattedStringVal; given an object value, a
         * valid Excel format pattern, return the formatted string value.
         *
         * @param Object  o
         * @param String  pattern	if General or "" returns string value
         * @param boolean isInteger	if General pattern, attempt to use integer value (rather than double)
         */
        fun getFormattedStringVal(o: Any?, pattern: String?/*, boolean isInteger*/): String {
            var o = o
            var pattern = pattern
            if (o == null) o = ""

            var isInteger = false
            isInteger = o is Int || o is Double && o.toInt().toDouble() == o.toDouble()

            if (pattern == null || pattern == "" || pattern.equals("GENERAL", ignoreCase = true)) {
                if (isInteger)
                    return java.lang.Double.valueOf(o.toString()).toInt().toString()
                else {    // general double numbers have default precision ...
                    try {
                        val d = Double(o.toString())
                        return ExcelTools.getNumberAsString(java.lang.Double.valueOf(o.toString()))    // handles default precision
                    } catch (e: NumberFormatException) {
                    }

                    return o.toString()
                }
            } else if (pattern == "000-00-0000") { // special case for SSN format ... sigh ...
                try {
                    Double(o.toString())    // if it can't be converted to a number, return original string (tis what excel does)
                    var s = o.toString()
                    while (s.length < 9)
                    // tis what excel does ...
                        s = '0' + s
                    return s.substring(0, 3) + "-" + s.substring(3, 5) + "-" + s.substring(5)
                } catch (e: Exception) {
                    return o.toString()
                }

            }


            /** try to determine if the format is numeric (+currency) or date  */
            var isNumeric = false
            var isDate = false
            var isString = false

            /** excel formats can have up to 4 parts:  <positive>;<negative>;<zero>;<text> </text></zero></negative></positive> */
            val pats = pattern.split(";".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()    // assign the correct pattern according to double or string value

            val tester = StringTool.convertPatternExtractBracketedExpression(pats[0])
            if (tester.matches(".*(((y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}).*)+).*".toRegex())) {
                isDate = true
                pats[0] = tester // ignore locale and other info for dates ...
            }
            if (!isDate) {
                var idx = pats.size - 1    // default with string
                try {
                    val d = Double(o.toString())
                    isNumeric = true
                    if (d > 0)
                    // 1st expression is for + numbers
                        idx = 0
                    else if (pats.size > 1 && d < 0)
                    // 2nd is for - numbers
                        idx = 1
                    else if (pats.size > 2 && d == 0.0)
                    // 3rd for 0
                        idx = 2
                    pattern = StringTool.convertPatternFromExcelToStringFormatter(pats[idx], d < 0)    // get correct format for String.format formatter
                } catch (e: NumberFormatException) {    // 4th for text (non-numeric)
                    if (pats.size > 3)
                        idx = 3
                    isString = true
                    pattern = StringTool.convertPatternFromExcelToStringFormatter(pats[idx], false)    // get correct format for String.format formatter
                }

            } else {
                pattern = pats[0]
                pattern = StringTool.convertDatePatternFromExcelToStringFormatter(pattern!!)    // get correct format for SimpleDateFormat
            }


            if (isString) {    // use string portion of format, if any
                try {
                    return String.format(pattern!!, o)
                } catch (e: IllegalFormatConversionException) {
                    return o.toString()
                }

            }
            if (isNumeric) {
                try {
                    var d = Double(o.toString())
                    if (!java.lang.Double.isNaN(d)) {
                        d = Math.abs(d)    // negative number intricacies have been handled in convertPattern method
                        // ugly, but has to be done ...
                        if (pattern!!.indexOf("%%") != -1)
                        // convert to percent
                            d *= 100.0
                        // special case of "@" -- integers converted to doubles format incorrectly ...
                        return if (pattern == "%s") {
                            o.toString()
                        } else String.format(pattern, d)
                    }
                } catch (e: Exception) {
                }

                return o.toString()
            }
            if (isDate) {
                try {
                    WorkBookHandle.simpledateformat.applyPattern(pattern!!)
                } catch (ex: Exception) {
                    return o.toString()
                }

                try {
                    return WorkBookHandle.simpledateformat.format(DateConverter.getCalendarFromNumber(o).time)
                    /*// KSC: TESTING
Date d= DateConverter.getCalendarFromNumber(o).getTime();
return WorkBookHandle.simpledateformat.format(d);*/
                } catch (e: NumberFormatException) {
                    try {
                        return WorkBookHandle.simpledateformat.format(Date(o.toString()).time)
                    } catch (i: IllegalArgumentException) {
                        if (o is Number)
                            Logger.logWarn("Unable to format date in $pattern")
                    }

                } catch (e: IllegalArgumentException) {
                    if (o is Number)
                        Logger.logWarn("Unable to format date in $pattern")
                }

            }
            // otherwise
            return o.toString()
        }


        /**
         * A FAIL FAST implementation for finding whether a cell string address
         * falls within a set of row/col range coordinates.
         *
         *
         * Sep 21, 2010
         *
         * @param rng      the range you want to test
         * @param rowFirst in the target range
         * @param rowLast  in the target range
         * @param colFirst in the target range
         * @param colLast  in the target range
         * @return
         */
        fun isInRange(rng: String, rowFirst: Int, rowLast: Int,
                      colFirst: Int, colLast: Int): Boolean {
            val sh = io.starter.OpenXLS.ExcelTools.getRowColFromString(rng)

            // the guantlet
            if (sh[1] < colFirst)
                return false
            if (sh[1] > colLast)
                return false
            return if (sh[0] < rowFirst) false else sh[0] <= rowLast
// passes!

        }

        /**
         * returns true if range intersects with range2
         *
         * @param rng
         * @param rc
         * @return
         */
        fun intersects(rng: String, rc: IntArray): Boolean {
            val rc2 = ExcelTools.getRangeCoords(rng)
            return (rc[0] >= rc2[0] && rc[2] <= rc2[2] && rc[1] >= rc2[1]
                    && rc[3] <= rc2[3])
        }

        /**
         * returns true if address is before the range coordinates defined by rc
         *
         * @param rc  row col of address
         * @param rng int[] coordinates as: row0, col0, row1, col1
         * @return true if address is before the range coordinates
         */
        fun isBeforeRange(rc: IntArray, rng: IntArray): Boolean {
            return rc[0] < rng[0] || rc[0] == rng[0] && rc[1] < rng[1]
        }

        /**
         * returns true if address is before the range coordinates defined by rc
         *
         * @param rc  row col of address
         * @param rng int[] coordinates as: row0, col0, row1, col1
         * @return true if address is before the range coordinates
         */
        fun isAfterRange(rc: IntArray, rng: IntArray): Boolean {
            return rc[0] > rng[2] || rc[0] == rng[2] && rc[1] > rng[3]
        }

        /**
         * Takes an input Object and attempts to convert to numeric Objects of the
         * highest precision possible.
         *
         *
         * This method is useful for avoiding the Excel warnings
         * "Number Stored As Text" when storing string data that contains numbers.
         *
         *
         * NOTE: this method is useful for ensuring that Formula references contain
         * true numeric values as not all String numbers are properly interpreted in
         * Formula engines, and can silently fail.
         *
         *
         * For this reason, always use numeric, non-string values to calculated
         * cells.
         *
         * @param input
         * @return
         */
        fun getObject(`in`: Any): Any {
            // do not record -- only called from other methods
            if (`in` !is String) {
                return `in`
            }

            val input = `in`.toString()
            var ret: Any = input // default is the original string

            try {
                ret = Double(input)
                return ret
            } catch (ex: NumberFormatException) {
                try {
                    ret = Float(input)
                    return ret
                } catch (ex2: NumberFormatException) {
                    try {
                        ret = Integer.valueOf(input)
                        return ret
                    } catch (ex3: NumberFormatException) {
                        // ret Is set outside of loop incase no match is ever found.
                    }

                }

            }

            // list of formatting chars to check for
            val fmtlist = arrayOf(arrayOf("$", ","), arrayOf(",", ","), arrayOf("%", ","))

            // strip the formatting
            for (t in fmtlist.indices) {
                if (input.indexOf(fmtlist[t][0]) > -1) { // contains!
                    var converted = StringTool.replaceText(input, fmtlist[t][0],
                            "") // strip first token (ie: '$')
                    converted = StringTool
                            .replaceText(converted, fmtlist[t][1], "") // strip
                    // second
                    // token
                    // (ie: ',')
                    try {
                        ret = Double(converted)
                        return ret
                    } catch (ex: NumberFormatException) {
                        try {
                            ret = Float(converted)
                            return ret
                        } catch (ex2: NumberFormatException) {
                            try {
                                ret = Integer.valueOf(converted)
                                return ret
                            } catch (ex3: NumberFormatException) {
                                // ret Is set outside of loop incase no match is
                                // ever found.
                            }

                        }

                    }

                }

            }

            return ret
        }

        /**
         * convert twips to pixels
         *
         *
         *
         *
         * In addition to a calculated size unit derived from the average size of
         * the default characters 0-9, Excel uses the 'twips' measurement which is
         * defined as:
         *
         *
         * 1 twip = 1/20 point or 20 twips = 1 point 1 twip = 1/567 centimeter or
         * 567 twips = 1 centimeter 1 twip = 1/1440 inch or 1440 twips = 1 inch
         *
         *
         *
         *
         * 1 pixel = 0.75 points 1 pixel * 1.3333 = 1 point 1 twip * 20 = 1 point
         *
         * @param pixels
         * @return twips
         */
        fun getPixels(twips: Float): Float {
            val points = twips / 20
            return points * 1.3333333f
        }

        /**
         * convert pixels to twips
         *
         *
         *
         *
         * In addition to a calculated size unit derived from the average size of
         * the default characters 0-9, Excel uses the 'twips' measurement which is
         * defined as:
         *
         *
         * 1 pixel = 0.75 points 1 pixel * 1.3333 = 1 point 1 twip * 20 = 1 point
         *
         * @param pixels
         * @return twips
         */
        fun getTwips(pixels: Float): Float {
            val points = pixels * .75f
            return points * 20
        }

        /**
         * get recordy byte def as a String
         *
         *
         * public static String getRecordByteDef(XLSRecord rec){ byte[] b =
         * rec.read(); StringBuffer sb = new StringBuffer("byte[] rbytes = {");
         * for(int t = 0;t<b.length></b.length>;t++){
         *
         *
         * Byte thisb = new Byte(b[t]);
         *
         *
         * sb.append(thisb.toString() + ", "); } sb.append("};"); return
         * sb.toString(); }
         */

        val logDate: String
            get() = Date(System.currentTimeMillis()).toString()

        /**
         * tracks minimal info container for counters -> start time, last time,
         * start mem, last mem
         *
         * @param info
         * @param perfobj
         */
        fun benchmark(info: String, perfobj: Any) {
            val rt = Runtime.getRuntime()
            var p: LongArray? = null
            var lasttime = 0L
            var lastmem = 0L
            if (System.getProperties()[perfobj.toString()] != null) {
                p = System.getProperties()[perfobj.toString()] as LongArray
                lasttime = p[1]
                lastmem = p[3]
                p[1] = System.currentTimeMillis()
                p[3] = rt.freeMemory()
                val elapsedsec = (p[1] - lasttime).toDouble()
                var usedmem = (lastmem - p[3]).toDouble() // - lastmem;
                if (usedmem < 0)
                    usedmem *= -1.0
                Logger.logInfo("$logDate $info")
                Logger.logInfo(" time: $elapsedsec millis")
                Logger.logInfo(" mem: $usedmem bytes.")
            } else {
                p = LongArray(4)
                p[0] = System.currentTimeMillis()
                p[1] = System.currentTimeMillis()
                p[2] = rt.freeMemory()
                p[3] = rt.freeMemory()
                lasttime = p[1]
                lastmem = p[3]
                System.getProperties()[perfobj.toString()] = p
            }

        }

        /**
         * get the bytes from a Vector of objects public byte[]
         * getBytesFrom(CompatibleVector objs){
         *
         *
         * for(int t = 0;t<objs.size></objs.size>();t++){
         *
         *
         * }
         *
         *
         *
         *
         * }
         */

        internal var alpharr = charArrayOf('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')
        /*
     * NOT USED ANYMORE -- see getAlphaVal -- OK to remove??
     */
        val ALPHASDELETE = arrayOf("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ", "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ", "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ")

        /**
         * get the Excel-style Column alphabetical representation of an integer
         * (0-based).
         *
         *
         * for example: 0 = A 26= AA 701= ZZ 702= AAA 16383= XFD (max)
         */
        fun getAlphaVal(i: Int): String {
            var i = i
            var ret = ""
            var leftover = 0
            if (i > 701) { // has 3rd digit
                var z = i / 676 - 1 // -1 to account for 0-based
                if (i % 676 < 26) { // then "leftover" is actually 2nd digit
                    z--
                    leftover = 676
                }
                ret = ExcelTools.alpharr[z].toString()
                i = i % 676
                i += leftover
            }
            if (i > 25) { // has 2nd digit
                val z = i / 26 - 1 // -1 to account for 0-based
                ret = ret + ExcelTools.alpharr[z]
                i = i % 26
            }

            // bbennett: this raises AIOOB if i = -1. In this situation we don't
            // care about alpha because it will just be in the message of the
            // exception that flags cell as non existent.
            ret += if (i < 0) Integer.toString(i) else ExcelTools.alpharr[i].toString()

            return ret
        }

        /**
         * get the int value of the Excel-style Column alpha representation.
         *
         * @param String column name
         * @return int the 0-based column number
         */
        fun getIntVal(c: String): Int {
            var c = c
            c = c.toUpperCase()
            if (c.length > 3)
            // max col value= XFD in Excel 2007
                return -1
            var i = c.length - 1
            var ret = 0
            while (i >= 0) { // process least to most-sigificant dig
                var z = 0
                val cc = c[i]
                while (z < alpharr.size && cc != alpharr[z++])
                ;
                z *= Math.pow(26.0, (c.length - i - 1).toDouble()).toInt() // 1-based col for computing
                ret += z
                i--
            }
            // make 0-based
            ret--
            return ret
        }

        /**
         * Parses an Excel cell address into row and column integers.
         *
         * @param address the address to parse, either A1 or R1C1
         * @return int[2]: [0] row index, [1] column index
         * @throws IllegalArgumentException if the argument is not a valid address
         */
        fun getRowColFromString(address: String): IntArray {
            var address = address

            if (address.indexOf("$") > -1) {
                address = StringTool.strip(address, "$")
            }
            if (address.indexOf("!") > -1) {
                address = address.substring(address.indexOf("!") + 1)
            }
            if (address.indexOf(":") > -1)
                return getRangeRowCol(address)

            val adrchars = address.toCharArray()
            var row = 0
            var col = 0
            var charpos = -1
            var numpos = -1
            var r1c1 = false
            for (i in adrchars.indices) {
                if (Character.isDigit(adrchars[i])) {
                    if (numpos == -1)
                    // its a number
                        numpos = i
                } else if (charpos == -1) {
                    charpos = i
                    if (numpos >= 0) { // we have already set number and we now have
                        // a nondigit - R1C1 style
                        r1c1 = true
                        // break, it's all over!
                        break
                    }
                }
            }

            if (r1c1) { // it's a single cell ref
                try {
                    // there's an R and a C, not adjacent
                    if (address.toUpperCase().indexOf("R") == 0) { // startwith R
                        val rx = address.substring(1, address.toUpperCase()
                                .indexOf("C"))
                        val cx = address.substring(address.toUpperCase()
                                .indexOf("C") + 1)
                        row = Integer.parseInt(rx)
                        col = Integer.parseInt(cx)
                    }
                } catch (e: NumberFormatException) {
                    throw IllegalArgumentException("illegal R1C1 address '"
                            + address + "'")
                }

            } else {
                row = 0 // -1 below
                col = -1
                if (charpos == 0 && numpos > 0) {
                    val colval = address.substring(0, numpos)
                    col = getIntVal(colval)
                    if (col < 0)
                        throw IllegalArgumentException("illegal column value '"
                                + colval + "' in address '" + address + "'")
                }
                if (numpos >= 0) {
                    row = Integer.parseInt(address.substring(numpos))
                    if (row < 1)
                        throw IllegalArgumentException(
                                "row may not be negative in address '" + address
                                        + "'")
                } else { // it's a wholecol ref
                    col = getIntVal(address)
                    if (col < 0)
                        throw IllegalArgumentException("illegal column value '"
                                + address + "' in address '" + address + "'")
                }
            }

            return intArrayOf(row - 1, col)
        }

        /**
         * Parses an Excel cell range and returns the addresses as an int array. The
         * range may not be qualified with sheet names. Strip them with
         * [.stripSheetNameFromRange] before calling this method. If the
         * argument is a single cell address it will be returned for both bounds.
         *
         * @param range the range to parse
         * @return int[4]: [0] first row, [1] first column, [2] second row, [3]
         * second column
         * @throws IllegalArgumentException if the addresses are invalid
         */
        fun getRangeRowCol(range: String): IntArray {
            val colon = range.indexOf(":")

            val firstloc: String
            val lastloc: String

            if (colon > -1) {
                firstloc = range.substring(0, colon)
                lastloc = range.substring(colon + 1)
            } else {
                firstloc = range
                lastloc = range
            }

            val result = IntArray(4)
            var temp: IntArray

            temp = getRowColFromString(firstloc)
            System.arraycopy(temp, 0, result, 0, 2)

            temp = getRowColFromString(lastloc)
            System.arraycopy(temp, 0, result, 2, 2)

            return result
        }

        /**
         * Takes an int array representing a row and column and formats it as a cell
         * address.
         *
         *
         * The index is zero-based.
         *
         *
         * [0][0] is "A1" [1][1] is "B2" [2][2] is "C3"
         *
         * @param int[] the numeric range to convert
         * @return String the string representation of the range
         */
        fun formatLocation(rowCol: IntArray): String {
            val sb = StringBuffer(getAlphaVal(rowCol[1]))
            sb.append(rowCol[0] + 1)

            // handle ranges
            if (rowCol.size > 3) {
                // 20090807: KSC: only a range if 1st is != 2nd cell :)
                if (rowCol[0] == rowCol[2] && rowCol[1] == rowCol[3])
                // it's a single address
                    return sb.toString()
                sb.append(":")
                sb.append(getAlphaVal(rowCol[3]))
                sb.append(rowCol[2] + 1)
            }

            return sb.toString()
        }

        /**
         * Takes an int array representing a row and column and formats it as a cell
         * address, taking into account relative or absolute refs
         *
         *
         * The index is zero-based.
         *
         *
         * [0][0] is "A1", $A1, A$1 or $A$1 depending upon bRelRow or bRelCol [1][1]
         * is "B2", $B1, B$1 or B$1 depending upon bRelRow or bRelCol [2][2] is
         * "C3", $C1, C$1 or $C$1 depending upon bRelRow or bRelCol
         *
         * @param int[]   the numeric range to convert
         * @param bRelRow if true, no "$"s are added, relative row reference
         * @param bRelCol if true, no "$"s are added, relative col reference
         * @return String the string representation of the range
         */
        fun formatLocation(s: IntArray, bRelRow: Boolean,
                           bRelCol: Boolean): String {
            val sb = StringBuffer(if (bRelCol) "" else "$")
            if (s[1] > -1)
            // account for WholeRow/WholeCol references
                sb.append(getAlphaVal(s[1]))

            if (s[0] > -1)
            // account for WholeRow/WholeCol references
                sb.append((if (bRelRow) "" else "$") + (s[0] + 1))
            // 20090906 KSC: handle ranges
            if (s.size > 3) {
                if (s[0] == s[2] && s[1] == s[3])
                // it's a single address
                    return sb.toString()
                sb.append(":")
                sb.append(if (bRelCol) "" else "$")
                sb.append(getAlphaVal(s[3]))
                sb.append((if (bRelRow) "" else "$") + (s[2] + 1))
            }

            return sb.toString()
        }

        /**
         * Takes an array of four shorts and formats it as a cell range.
         *
         *
         * IE [0][3][1][4] would be "A2:B3"
         *
         * @param int[] the numeric range to convert
         * @return String the string representation of the range
         */
        fun formatRange(s: IntArray): String {
            if (s.size != 4)
                return "incorrect array size in ExcelTools.formatLocation"
            val temp = IntArray(2)
            temp[0] = s[1]
            temp[1] = s[0]
            val firstcell = formatLocation(temp)
            temp[0] = s[3]
            temp[1] = s[2]
            val lastcell = formatLocation(temp)
            return "$firstcell:$lastcell"
        }

        /**
         * format a range as a string, range in format of [r][c][r1][c1]
         *
         * @param s
         * @return String representation of the integers as a range, ie A1:B4
         */
        fun formatRangeRowCol(s: IntArray): String {
            if (s.size != 4)
                return "incorrect array size in ExcelTools.formatLocation"
            val temp = IntArray(2)
            temp[0] = s[0]
            temp[1] = s[1]
            val firstcell = formatLocation(temp)
            temp[0] = s[2]
            temp[1] = s[3]
            val lastcell = formatLocation(temp)
            return "$firstcell:$lastcell"
        }

        /**
         * format a range as a string, range in format of [r][c][r1][c1] including
         * relative address state
         *
         * @param s
         * @param bRelAddresses contains relative row and col state for each rcr1c1
         * @return String representation of the integers as a range, ie A1:B4
         */
        fun formatRangeRowCol(s: IntArray, bRelAddresses: BooleanArray): String {
            if (s.size != 4)
                return "incorrect array size in ExcelTools.formatLocation"
            val temp = IntArray(2)
            temp[0] = s[0]
            temp[1] = s[1]
            val firstcell = formatLocation(temp, bRelAddresses[0],
                    bRelAddresses[1])
            temp[0] = s[2]
            temp[1] = s[3]
            val lastcell = formatLocation(temp, bRelAddresses[2],
                    bRelAddresses[3])
            // unfortunately, no formatLocation can do this. This is
            // formattingRANGERowCol
            // if (firstcell.equals(lastcell)) return firstcell; // 20090309 KSC:
            return "$firstcell:$lastcell"
        }

        /**
         * Transforms a string to an array of ints for evaluation purposes. For
         * example, acdc == [0][2][3][2]
         */
        fun transformStringToIntVals(trans: String): IntArray {
            val intarr = IntArray(trans.length)
            for (i in 0 until trans.length) {
                val c = trans[i]
                for (x in alpharr.indices) {
                    if (c.toString().equals(
                                    alpharr[x].toString(), ignoreCase = true)) {
                        intarr[i] = x
                    }
                }
            }
            return intarr
        }

        /**
         * Formats a string representation of a numeric value as a string in the
         * specified notation:
         *
         * @param int <br></br>
         * NOTATION_STANDARD = 0, <br></br>
         * NOTATION_SCIENTIFIC = 1, <br></br>
         * NOTATION_SCIENTIFIC_EXCEL = 2, <br></br>
         * EXTENXLS_NOTATION = 3
         *
         *
         *
         *
         * example: formatNumericNotation(1.23456E5, 0) returns a "123456"
         * example: formatNumericNotation(123456, 1) returns "1.23456E5"
         * example: formatNumericNotation(123456, 2) returns "1.23456E+5"
         * example: formatNumericNotation(123456, 3) returns "1.23456E+5"
         */
        fun formatNumericNotation(num: String, notationType: Int): String {
            var num = num
            // if (notationType > 2)return null;
            var negative = false
            if (num.substring(0, 1) == "-") {
                negative = true
                num = num.substring(1)
            }
            var preString: String
            val postString: String
            var fullString: String? = ""
            when (notationType) {
                0 // NOTATION_STANDARD
                -> {
                    val i = num.indexOf("E")
                    if (i == -1) { // just return
                        if (num.substring(num.length - 2) == ".0") {
                            num = num.substring(0, num.length - 2)
                        }
                        return if (negative) {
                            "-$num"
                        } else num
                    }
                    preString = num.substring(0, i)
                    var outNumD = CompatibleBigDecimal(preString)
                    var exp = ""
                    if (num.indexOf("+") == -1) {
                        exp = num.substring(i + 1)
                    } else {
                        exp = num.substring(i + 2)
                    }
                    val expNum = Integer.valueOf(exp).toInt()
                    outNumD = CompatibleBigDecimal(outNumD.movePointRight(expNum))
                    // outNumD = outNumD.multiply(new CompatibleBigDecimal(Math.pow(10,
                    // expNum)));
                    // Logger.logInfo(String.valueOf(outNumD));
                    // outNum = Math.r

                    // check if we should be returning a whole number or a decimal
                    val moveLen = num.indexOf("E") - num.indexOf(".") - 1
                    if (expNum >= moveLen) {
                        return if (negative) {
                            "-" + Math.round(outNumD.toDouble())
                        } else {
                            Math.round(outNumD.toDouble()).toString()
                        }
                    }
                    val args = arrayOfNulls<Any>(0)
                    // args[0] = outNumD;
                    val res = ResourceLoader.executeIfSupported(outNumD, args,
                            "toPlainString")
                    if (res != null) {
                        fullString = res.toString()
                    } else {
                        fullString = outNumD.toCompatibleString()
                    }
                }
                1 // NOTATION_SCIENTIFIC
                -> if (num.indexOf("E") != -1 && num.indexOf("+") == -1) {
                    fullString = num
                } else if (num.indexOf("+") != -1) {
                    preString = num.substring(0, num.indexOf("+"))
                    postString = num.substring(num.indexOf("+") + 1)
                    return preString + postString
                } else if (num.indexOf(".") != -1) {
                    val pos = num.indexOf(".")
                    preString = (num.substring(0, 1) + "."
                            + num.substring(1, num.indexOf(".")))
                    var d = CompatibleBigDecimal(num)
                    if (d.toDouble() < 1 && d.toDouble() != 0.0) {
                        // it is a very small value, ie 1.0E-10
                        var counter = 0
                        while (d.toDouble() < 1) {
                            d = CompatibleBigDecimal(d.movePointRight(1))
                            counter++
                        }
                        return d.toCompatibleString() + "E-" + counter
                    }
                    postString = num.substring(num.indexOf(".") + 1)
                    fullString = preString + postString
                    fullString = fullString + "E" + (pos - 1)
                } else {
                    preString = num.substring(0, 1) + "."
                    if (num.length > 1) {
                        preString += num.substring(1)
                    } else {
                        preString += "0"
                    }
                    fullString = preString + "E" + (num.length - 1)
                }
                2 // NOTATION_SCIENTIFIC_EXCEL
                -> if (num.indexOf("E") != -1 && num.indexOf("+") != -1) {
                    fullString = num
                } else if (num.indexOf("E") != -1) {
                    preString = num.substring(0, num.indexOf("E") + 1)
                    postString = "+" + num.substring(num.indexOf("E") + 1)
                    fullString = preString + postString
                } else if (num.indexOf(".") != -1) {
                    val pos = num.indexOf(".")
                    var d = CompatibleBigDecimal(num)
                    if (d.toDouble() < 1 && d.toDouble() != 0.0) {
                        // it is a very small value, ie 1.0E-10
                        var counter = 0
                        while (d.toDouble() < 1) {
                            d = CompatibleBigDecimal(d.movePointRight(1))
                            counter++
                        }
                        return d.toCompatibleString() + "E-" + counter
                    }
                    preString = (num.substring(0, 1) + "."
                            + num.substring(1, num.indexOf(".")))
                    postString = num.substring(num.indexOf(".") + 1)
                    fullString = preString + postString
                    fullString = fullString + "E+" + (pos - 1)
                } else {
                    preString = num.substring(0, 1) + "."
                    if (num.length > 1) {
                        preString += num.substring(1)
                    } else {
                        preString += "0"
                    }
                    fullString = preString + "E+" + (num.length - 1)
                }
                else -> return num
            }
            if (negative)
                fullString = "-" + fullString!!
            return fullString
        }

        /**
         * Return an array of cell handles specified from the string passed in.
         *
         *
         * Note that a CellHandle cannot exist for an empty cell, so the cells
         * retrieved in this manner will be blank cells, not empty cells.
         *
         * @param cellstr - a comma delimited String representing cells and cell ranges,
         * example "A1,A5,A6,B1:B5" would return cells A1, A5, A6, B1,
         * B2, B3, B4, B5
         * @param sheet   the worksheet containing the cells.
         * @return CellHandle[]
         */
        fun getCellHandlesFromSheet(strRange: String,
                                    sheet: WorkSheetHandle): Array<CellHandle> {
            var retCells: Array<CellHandle>
            val cellTokenizer = StringTokenizer(strRange, ",")
            val cells = ArrayList<CellHandle>()
            do {
                val element = cellTokenizer.nextElement() as String
                if (element.indexOf(":") != -1) {
                    val aRange = CellRange(sheet.sheetName + "!"
                            + strRange, sheet.workBook, true)
                    cells.addAll(aRange.cellList)
                } else {
                    var aCell: CellHandle? = null
                    try {
                        aCell = sheet.getCell(element)
                    } catch (ce: Exception) {
                        aCell = sheet.add(null, element)
                    }

                    if (aCell != null)
                        cells.add(aCell)
                }
            } while (cellTokenizer.hasMoreElements())
            retCells = arrayOfNulls(cells.size)
            retCells = cells.toTypedArray<CellHandle>()
            return retCells
        }

        /**
         * Strip sheet name(s) from range string can be Sheet1!AB:Sheet!BC or
         * Sheet!AB:BC or AB:BC or Sheet1:Sheet2!A1:A2
         *
         * @param address or range String
         * @return 1st sheetname
         *
         *
         * Ok, this is a strange method. It returns a string array of the
         * following format 0 - sheetname1 1 - cell address or range (what
         * if there are 2?) 2 - sheetname2 3 - external link 1 ?? some ooxml
         * record 4 - external link 2
         */
        fun stripSheetNameFromRange(address: String): Array<String> {
            var address = address
            var sheetname: String? = null
            var sheetname2: String? = null
            var m = address.indexOf('!')
            if (m > -1) {
                if (address.substring(0, m).indexOf(":") == -1)
                    sheetname = address.substring(0, m)
                else {
                    val z = address.indexOf(":")
                    sheetname = address.substring(0, z)
                    sheetname2 = address.substring(z + 1, m)
                }
            }
            address = address.substring(m + 1)
            val n = address.indexOf('!') // see if 2nd sheet name exists
            if (n > -1 && address != "#REF!") {
                m = address.indexOf(':')
                sheetname2 = address.substring(m + 1, n)
                m = address.indexOf(':')
                address = address.substring(0, m + 1) + address.substring(n + 1)
            }
            // 20090323 KSC: handle external references (OOXML-Specific format of
            // [#]SheetName!Ref where # denotes ExternalLink workbook
            var exLink1: String? = null
            var exLink2: String? = null
            if (sheetname != null && sheetname.indexOf('[') >= 0) { // External
                // OOXML
                // reference
                exLink1 = sheetname.substring(sheetname.indexOf('['))
                exLink1 = exLink1!!.substring(0, exLink1.indexOf(']') + 1)
                sheetname = StringTool.replaceText(sheetname, exLink1, "")
                if (sheetname == "")
                    sheetname = null // possible to have address in form of =
                // [#]!Name or range
            }
            if (sheetname2 != null && sheetname2.indexOf('[') >= 0) { // External
                // OOXML
                // reference
                exLink2 = sheetname2.substring(sheetname2.indexOf('['))
                exLink2 = exLink2!!.substring(0, exLink2.indexOf(']') + 1)
                sheetname2 = StringTool.replaceText(sheetname2, exLink2, "")
                if (sheetname2 == "")
                    sheetname2 = null // possible to have address in form of =
                // [#]!Name or range
            }
            // return new String[]{sheetname, address, sheetname2};
            return arrayOf<String>(sheetname, address, sheetname2, exLink1, exLink2) // 20090323
            // KSC:
            // add
            // any
            // external
            // link
            // info
        }

        /**
         * return the first and last coords of a range in int form + the number of
         * cells in the range range is in the format of Sheet
         */
        fun getRangeCoords(range: String): IntArray {
            var numrows = 0
            var numcols = 0
            var numcells = 0
            val coords = IntArray(5)
            var temprange = range
            // figure out the sheet bounds using the range string
            temprange = ExcelTools.stripSheetNameFromRange(temprange)[1]
            var startcell = ""
            var endcell = ""
            val lastcolon = temprange.lastIndexOf(":")
            endcell = temprange.substring(lastcolon + 1)
            if (lastcolon == -1)
            // no range
                startcell = endcell
            else
                startcell = temprange.substring(0, lastcolon)
            startcell = StringTool.strip(startcell, "$")
            endcell = StringTool.strip(endcell, "$")

            // get the first cell's coordinates
            var charct = startcell.length
            while (charct > 0) {
                if (!Character.isDigit(startcell[--charct])) {
                    charct++
                    break
                }
            }
            val firstcellrowstr = startcell.substring(charct)
            var firstcellrow = -1
            try {
                firstcellrow = Integer.parseInt(firstcellrowstr)
            } catch (e: NumberFormatException) { // could be a whole-col-style ref
            }

            val firstcellcolstr = startcell.substring(0, charct).trim({ it <= ' ' })
            val firstcellcol = ExcelTools.getIntVal(firstcellcolstr)
            // get the last cell's coordinates
            charct = endcell.length
            while (charct > 0) {
                if (!Character.isDigit(endcell[--charct])) {
                    charct++
                    break
                }
            }
            val lastcellrowstr = endcell.substring(charct)
            var lastcellrow = -1
            try {
                lastcellrow = Integer.parseInt(lastcellrowstr)
            } catch (e: NumberFormatException) { // could be a whole-col-style ref
            }

            val lastcellcolstr = endcell.substring(0, charct)
            val lastcellcol = ExcelTools.getIntVal(lastcellcolstr)
            numrows = lastcellrow - firstcellrow + 1
            numcols = lastcellcol - firstcellcol + 1
            /*
         * if(numrows == 0)numrows =1; if(numcols == 0)numcols =1;
         */
            numcells = numrows * numcols
            if (numcells < 0)
                numcells *= -1 // handle swapped cells ie: "B1:A1"

            coords[0] = firstcellrow
            coords[1] = firstcellcol
            coords[2] = lastcellrow
            coords[3] = lastcellcol
            coords[4] = numcells
            // Trap errors in range
            // if (firstcellrow < 0 || lastcellrow < 0 || firstcellcol < 0 ||
            // lastcellcol < 0)
            // Logger.logErr("ExcelTools.getRangeCoords: Error in Range " + range);
            return coords
        }
    }
}