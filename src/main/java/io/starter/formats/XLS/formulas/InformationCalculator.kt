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
import io.starter.toolkit.Logger

import java.text.DecimalFormat


/*
    InformationCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/

object InformationCalculator {


    /**
     * CELL
     * Returns information about the formatting, location, or contents of a cell
     * The CELL function returns information about the formatting, location, or contents of a cell.
     *
     *
     * CELL(info_type, [reference])
     * info_type  Required. A text value that specifies what type of cell information you want to return. The following list shows the possible values of the info_type argument and the corresponding results.info_type Returns
     * "address" Reference of the first cell in reference, as text.
     * "col" Column number of the cell in reference.
     * "color" The value 1 if the cell is formatted in color for negative values; otherwise returns 0 (zero).
     * "contents" Value of the upper-left cell in reference; not a formula.
     * "filename" Filename (including full path) of the file that contains reference, as text. Returns empty text ("") if the worksheet that contains reference has not yet been saved.
     * "format" Text value corresponding to the number format of the cell. The text values for the various formats are shown in the following table. Returns "-" at the end of the text value if the cell is formatted in color for negative values. Returns "()" at the end of the text value if the cell is formatted with parentheses for positive or all values.
     * "parentheses" The value 1 if the cell is formatted with parentheses for positive or all values; otherwise returns 0.
     * "prefix" Text value corresponding to the "label prefix" of the cell. Returns single quotation mark (') if the cell contains left-aligned text, double quotation mark (") if the cell contains right-aligned text, caret (^) if the cell contains centered text, backslash (\) if the cell contains fill-aligned text, and empty text ("") if the cell contains anything else.
     * "protect" The value 0 if the cell is not locked; otherwise returns 1 if the cell is locked.
     * "row" Row number of the cell in reference.
     * "type" Text value corresponding to the type of data in the cell. Returns "b" for blank if the cell is empty, "l" for label if the cell contains a text constant, and "v" for value if the cell contains anything else.
     * "width" Column width of the cell, rounded off to an integer. Each unit of column width is equal to the width of one character in the default font size.
     *
     *
     * reference  Optional. The cell that you want information about. If omitted, the information specified in the info_type argument is returned for the last cell that was changed. If the reference argument is a range of cells, the CELL function returns the information for only the upper left cell of the range.
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun calcCell(operands: Array<Ptg>): Ptg {
        val type = operands[0].value.toString().toLowerCase()
        var ref: PtgRef? = null
        var cell: BiffRec? = null
        if (operands.size > 1) {
            ref = operands[1] as PtgRef
            try {
                cell = ref.parentRec!!.workBook!!.getCell(ref.locationWithSheet)
            } catch (e: CellNotFoundException) {
                try {
                    var sh: String? = null
                    try {
                        sh = ref.sheetName
                    } catch (we: WorkSheetNotFoundException) {
                    }

                    if (sh == null) sh = ref.parentRec!!.sheet!!.sheetName
                    cell = ref.parentRec!!.workBook!!.getWorkSheetByName(sh).addValue(null, ref.location)
                } catch (ex: Exception) {
                    return PtgErr(PtgErr.ERROR_VALUE)
                }

            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_VALUE)
            }

            //  If ref param is omitted, the information specified in the info_type argument
            // is returned for the last cell that was changed
        } else if (type != "filename")
        // no ref was passed in and option is not "filename"
        // We cannot determine which is the "last cell" they are referencing;
            throw FunctionNotSupportedException("Worsheet function CELL with no reference parameter is not supported")
        else
        // filename option can use any biffrec ...
            cell = operands[0].parentRec

        // at this point both ref (PtgRef) and r (BiffRec) should be valid
        try {
            if (type == "address") {
                val newref = ref
                newref!!.clearLocationCache()
                newref.fColRel = false        // make absolute
                newref.fRwRel = false
                return PtgStr(newref.location)
            } else if (type == "col") {
                return PtgNumber((ref!!.intLocation!![1] + 1).toDouble())
            } else if (type == "color") {    // The value 1 if the cell is formatted in color for negative values; otherwise returns 0 (zero).
                val s = cell!!.formatPattern
                return if (s.indexOf(";[Red") > -1) PtgNumber(1.0) else PtgNumber(0.0)
            } else if (type == "contents") {// Value of the upper-left cell in reference; not a formula.
                return PtgStr(cell!!.stringVal)
            } else if (type == "filename") {
                var f = cell!!.workBook.fileName
                val sh = cell.sheet.sheetName
                val i = f.lastIndexOf(java.io.File.separatorChar.toInt())
                f = f.substring(0, i + 1) + "[" + f.substring(i + 1)
                f += "]$sh"
                return PtgStr(f)
            } else if (type == "format") {    // Text value corresponding to the number format of the cell. The text values for the various formats are shown in the following table. Returns "-" at the end of the text value if the cell is formatted in color for negative values. Returns "()" at the end of the text value if the cell is formatted with parentheses for positive or all values.
                val s = cell!!.formatPattern
                var ret = "G"    // default?
                if (s == "General" ||
                        s == "# ?/?" ||
                        s == "# ??/??") {
                    ret = "G"
                } else if (s == "0") {
                    ret = "F0"
                } else if (s == "#,##0") {
                    ret = ",0"
                } else if (s == "0.00") {
                    ret = "F2"
                } else if (s == "#,##0.00") {
                    ret = ", 2"
                } else if (s == "$#,##0_);($#,##0)") {
                    ret = "C0"
                } else if (s == "$#,##0_);[Red]($#,##0)") {
                    ret = "C0-"
                } else if (s == "$#,##0.00_);($#,##0.00)") {
                    ret = "C2"
                } else if (s == "$#,##0.00_);[Red]($#,##0.00)") {
                    ret = "C2-"
                } else if (s == "0%") {
                    ret = "P0"
                } else if (s == "0.00%") {
                    ret = "P2"
                } else if (s == "0.00E+00") {
                    ret = "S2"
                    //					   m/d/yy or m/d/yy h:mm or mm/dd/yy 	"D4"
                } else if (s == "m/d/yy" ||
                        s == "m/d/yy h:mm" ||
                        s == "mm/dd/yy" ||
                        s == "mm-dd-yy") {        // added last to accomodate Excel's regional short date setting (format #14)
                    ret = "D4"
                } else if (s == "d-mmm-yy" || s == "dd-mmm-yy") {
                    ret = "D1"
                } else if (s == "d-mmm" || s == "dd-mmm") {
                    ret = "D2"
                } else if (s == "mmm-yy") {
                    ret = "D3"
                } else if (s == "mm/dd") {
                    ret = "D5"
                } else if (s == "h:mm AM/PM") {
                    ret = "D7"
                } else if (s == "h:mm:ss AM/PM") {
                    ret = "D6"
                } else if (s == "h:mm") {
                    ret = "D9"
                } else if (s == "h:mm:ss") {
                    ret = "D8"
                }
                return PtgStr(ret)
            } else if (type == "parentheses") {
                val s = cell!!.formatPattern
                return if (s.startsWith("(")) PtgNumber(1.0) else PtgNumber(0.0)
            } else if (type == "prefix") {
                // TODO: THIS IS NOT CORRECT - EITHER INFORM USER OR ??
                // DOESN'T APPEAR TO MATCH EXCEL
                //Text value corresponding to the "label prefix" of the cell.
                // Returns single quotation mark (') if the cell contains left-aligned text, double quotation mark (") if the cell contains right-aligned text,
                // caret (^) if the cell contains centered text, backslash (\) if the cell contains fill-aligned text, and empty text ("") if the cell contains anything else.
                val al = cell!!.xfRec.horizontalAlignment
                if (al == FormatConstants.ALIGN_LEFT)
                    return PtgStr("'")
                if (al == FormatConstants.ALIGN_CENTER)
                    return PtgStr("^")
                if (al == FormatConstants.ALIGN_RIGHT)
                    return PtgStr("\"")
                return if (al == FormatConstants.ALIGN_FILL) PtgStr("\\") else PtgStr("")
            } else if (type == "protect") {
                return if (cell!!.xfRec.isLocked) PtgNumber(1.0) else PtgNumber(0.0)
            } else if (type == "row") {
                return PtgNumber((ref!!.intLocation!![0] + 1).toDouble())
            } else if (type == "type") {
                //Text value corresponding to the type of data in the cell.
                // Returns "b" for blank if the cell is empty,
                //"l" for label if the cell contains a text constant, and
                // "v" for value if the cell contains anything else.
                if ((cell as XLSRecord).isBlank)
                    return PtgStr("b")
                return if (cell is Labelsst) PtgStr("l") else PtgStr("v")
            } else if (type == "width") {
                var n = 0
                n = cell!!.sheet.getColInfo(cell.colNumber.toInt())!!.colWidthInChars
                return PtgNumber(n.toDouble())
            }
        } catch (e: Exception) {
            Logger.logWarn("CELL: unable to calculate: $e")
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * ERROR.TYPE
     * Returns a number corresponding to an error type
     * If error_val is
     * ERROR.TYPE returns
     *
     *
     * #NULL!	 1
     * #DIV/0! 2
     * #VALUE! 3
     * #REF!	 4
     * #NAME?	 5
     * #NUM!	 6
     * #N/A	 7
     * Anything else	 #N/A
     */
    internal fun calcErrorType(operands: Array<Ptg>): Ptg {
        val o = operands[0].value
        val s = o.toString()
        if (s.equals("#NULL!", ignoreCase = true)) return PtgInt(1)
        if (s.equals("#DIV/0!", ignoreCase = true)) return PtgInt(2)
        if (s.equals("#VALUE!", ignoreCase = true)) return PtgInt(3)
        if (s.equals("#REF!", ignoreCase = true)) return PtgInt(4)
        if (s.equals("#NAME?", ignoreCase = true)) return PtgInt(5)
        if (s.equals("#NUM!", ignoreCase = true)) return PtgInt(6)
        return if (s.equals("#N/A", ignoreCase = true)) PtgInt(7) else PtgErr(PtgErr.ERROR_NA)
    }

    /**
     * INFO
     * Returns information about the current operating environment
     * INFO(type_text)
     *
     *
     * NOTE: Several options are incomplete:
     * "osversion"	-- only valid for Windows versions
     * "system" 	-- only valid for Windows and Mac
     * "release"	-- incomplete
     * "origin"	-- does not return R1C1 format
     */
    internal fun calcInfo(operands: Array<Ptg>?): Ptg {
        // validate
        if (operands == null || operands.size == 0 || operands[0].parentRec == null)
            return PtgErr(PtgErr.ERROR_VALUE)
        val type_text = operands[0].string
        val ret = ""
        if (type_text == "directory")
        // Path of the current directory or folder
            return PtgStr(System.getProperty("user.dir").toLowerCase() + "\\")
        else if (type_text == "numfile")
        // number of active worksheets in the current workbook
        // TODO: what is correct definition of "Active Worksheets"  - hidden state doesn't seem to affect"
            return PtgNumber(operands[0].parentRec.workBook!!.numWorkSheets.toDouble())
        else if (type_text == "origin") {        /* Returns the absolute cell reference of the top and leftmost
	   											    cell visible in the window, based on the current scrolling
	   											    position, as text prepended with "$A:".
	   											    This value is intended for for Lotus 1-2-3 release 3.x compatibility.
	   											    The actual value returned depends on the current reference
	   											    style setting. Using D9 as an example, the return value would be:
	   											    A1 reference style   "$A:$D$9".
													R1C1 reference style  "$A:R9C4"
	    										*/
            // TODO: FINISH R1C1 reference style
            var cell = operands[0].parentRec.sheet!!.window2!!.topLeftCell
            for (i in cell.length - 1 downTo 0) {
                if (!Character.isDigit(cell[i])) {
                    cell = cell.substring(0, i + 1) + "$" + cell.substring(i + 1)
                    break
                }
            }
            cell = "\$A:$$cell"
            return PtgStr(cell)
        } else if (type_text == "osversion") {    //Current operating system version, as text.
            // see end of file for os info
            val osversion = System.getProperty("os.version")
            val n = System.getProperty("os.name")    // Windows Vista
            var os = ""
            // TODO:  need a list of osversions to compare to!  have know idea for mac, linux ...
            if (n.startsWith("Windows")) {
                val v = Double(osversion)
                os = "Windows (32-bit) "
                if (v >= 5)
                    os += "NT "
                val df = DecimalFormat("##.00")
                os += df.format(v)
            } // otherwise have NO idea as cannot find any info on net
            else
                os += osversion
            return PtgStr(os)
        } else if (type_text == "recalc") {    //Current recalculation mode; returns "Automatic" or "Manual".
            return if (operands[0].parentRec.workBook!!.recalculationMode == 0) PtgStr("Manual") else PtgStr("Automatic")
        } else if (type_text == "release") {    //Version of Microsoft Excel, as text.
            // TODO: Finish!  97= 8.0, 2000= 9.0, 2002 (XP)= 10.0, 2003= 11.0, 2007= 12.0
            Logger.logWarn("Worksheet Function INFO(\"release\") is not supported")
            return PtgStr("")
        } else if (type_text == "system")
        // Name of the operating environment: Macintosh = "mac" Windows = "pcdos"
        // TODO: linux?  ****************
            return if (System.getProperty("os.name").indexOf("Windows") >= 0)
                PtgStr("pcdos")
            else
                PtgStr("mac")
        else if (type_text == "memavail" ||
                type_text == "memused" ||
                type_text == "totmem")
            return PtgErr(PtgErr.ERROR_NA)// In previous versions of Microsoft Office Excel, the "memavail", "memused", and "totmem" type_text values, returned memory information.
        // These type_text values are no longer supported and now return a #N/A error value.
        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * ISBLANK
     * ISBLANK determines if the cell referenced is blank, and returns
     * a boolean ptg based off that
     */
    internal fun calcIsBlank(operands: Array<Ptg>): Ptg {
        val allops = PtgCalculator.getAllComponents(operands)
        for (i in allops.indices) {
            // 20081120 KSC: blanks are handled differently now as Excel counts blank cells as 0's
            /*Object o = allops[i].getValue();
			if (o != null) return new PtgBool(false);
			*/
            if (!allops[i].isBlank)
                return PtgBool(false)
        }
        return PtgBool(true)
    }

    /**
     * ISERROR
     * Value refers to any error value
     * (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
     * Usage@ ISERROR(value)
     * Return@ PtgBool
     */
    internal fun calcIserror(operands: Array<Ptg>): Ptg {
        if (operands[0] is PtgErr) {
            return PtgBool(true)
        }
        val errorstr = arrayOf("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!")
        val o = operands[0].value
        val opval = o.toString()
        for (i in errorstr.indices) {
            if (opval.equals(errorstr[i], ignoreCase = true)) {
                return PtgBool(true)
            }
        }
        return PtgBool(false)
    }

    /**
     * ISERR
     * Returns TRUE if the value is any error value except #N/A
     */
    internal fun calcIserr(operands: Array<Ptg>): Ptg {
        val errorstr = arrayOf("#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!")
        if (operands.size != 1) return PtgErr(PtgErr.ERROR_VALUE)
        val o = operands[0].value
        val opval = o.toString()
        for (i in errorstr.indices) {
            if (opval.equals(errorstr[i], ignoreCase = true)) {
                return PtgBool(true)
            }
        }
        return PtgBool(false)
    }

    /**
     * ISEVEN(number)
     *
     *
     * Number   is the value to test. If number is not an integer, it is truncated.
     *
     *
     * Remarks
     * If number is nonnumeric, ISEVEN returns the #VALUE! error value.
     * Examples
     * ISEVEN(-1) equals FALSE
     * ISEVEN(2.5) equals TRUE
     * ISEVEN(5) equals FALSE
     *
     *
     * author: John
     */
    internal fun calcIsEven(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        val allops = PtgCalculator.getAllComponents(operands)
        if (allops.size > 1) return PtgBool(false)
        val o = operands[0].value
        if (o != null) {
            try {    // KSC: mod for different number types + mod typo
                if (o is Int) {
                    val s = o.toInt()
                    return if (s < 0) PtgBool(false) else PtgBool(s % 2 == 0)
                } else if (o is Float) {
                    val s = o.toFloat()
                    return if (s < 0) PtgBool(false) else PtgBool(s % 2 == 0f)
                } else if (o is Double) {
                    val s = o.toDouble()
                    return if (s < 0) PtgBool(false) else PtgBool(s % 2 == 0.0)
                }
            } catch (e: Exception) {
            }

        }
        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * ISLOGICAL
     * Returns TRUE if the value is a logical value
     */
    internal fun calcIsLogical(operands: Array<Ptg>): Ptg {
        val allops = PtgCalculator.getAllComponents(operands)
        if (allops.size > 1) return PtgBool(false)
        val o = operands[0].value
        // unfortunately we need to know the difference between
        // "true" and true, if it's a reference this can be difficult
        try {
            val b = o as Boolean
            return PtgBool(true)
        } catch (e: ClassCastException) {
        }

        return PtgBool(false)
    }

    /**
     * ISNUMBER
     * Returns TRUE if the value is a number
     */
    internal fun calcIsNumber(operands: Array<Ptg>): Ptg {
        val allops = PtgCalculator.getAllComponents(operands)
        if (allops.size > 1) return PtgBool(false)
        val o = operands[0].value
        try {
            val f = o as Float
            return PtgBool(true)
        } catch (e: ClassCastException) {
            try {
                val d = o as Double
                return PtgBool(true)
            } catch (ee: ClassCastException) {
                try {
                    val ii = o as Int
                    return PtgBool(true)
                } catch (eee: ClassCastException) {
                }

            }

        }

        return PtgBool(false)
    }

    /**
     * ISNONTEXT
     * Returns TRUE if the value is not text
     */
    internal fun calcIsNonText(operands: Array<Ptg>): Ptg {
        val allops = PtgCalculator.getAllComponents(operands)
        if (allops.size > 1) return PtgBool(false)
        // blanks return true for this test
        if (allops[0].isBlank)
            return PtgBool(true)
        val o = operands[0].value
        if (o != null) {
            try {
                val s = o as String
                return PtgBool(false)
            } catch (e: ClassCastException) {
            }

        }
        return PtgBool(true)
    }

    /**
     * ISNA
     * Value refers to the #N/A
     * (value not available) error value.
     *
     *
     * usage@ ISNA(value)
     * return@ PtgBool
     */
    internal fun calcIsna(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        if (operands[0] is PtgErr) {
            val per = operands[0] as PtgErr
            if (per.errorType == PtgErr.ERROR_NA) {
                return PtgBool(true)
            }
        } else if (operands[0].isReference) {
            val o = operands[0].value
            if (o.toString().equals(PtgErr(PtgErr.ERROR_NA).toString(), ignoreCase = true)) {
                return PtgBool(true)
            }
        }
        return PtgBool(false)
    }

    /**
     * NA
     * Returns the error value #N/A
     */
    internal fun calcNa(operands: Array<Ptg>): Ptg {
        return PtgErr(PtgErr.ERROR_NA)
    }

    /**
     * ISTEXT
     * Returns TRUE if the value is text
     */
    internal fun calcIsText(operands: Array<Ptg>): Ptg {
        val allops = PtgCalculator.getAllComponents(operands)
        if (allops.size > 1) return PtgBool(false)
        val o = operands[0].value
        if (o != null) {
            try {
                val s = o as String
                return PtgBool(true)
            } catch (e: ClassCastException) {
            }

        }
        return PtgBool(false)
    }


    /**
     * ISODD
     * Returns TRUE if the number is odd
     *
     *
     * author: John
     */
    internal fun calcIsOdd(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        val allops = PtgCalculator.getAllComponents(operands)
        if (allops.size > 1) return PtgBool(false)
        val o = operands[0].value
        if (o != null) {
            try {    // KSC: mod for different number types + mod typo
                if (o is Int) {
                    val s = o.toInt()
                    return if (s < 0) PtgBool(false) else PtgBool(s % 2 != 0)
                } else if (o is Float) {
                    val s = o.toFloat()
                    return if (s < 0) PtgBool(false) else PtgBool(s % 2 != 0f)
                } else if (o is Double) {
                    val s = o.toDouble()
                    return if (s < 0) PtgBool(false) else PtgBool(s % 2 != 0.0)
                }
            } catch (e: Exception) {
            }

        }
        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * ISREF
     * Returns TRUE if the value is a reference
     */
    internal fun calcIsRef(operands: Array<Ptg>): Ptg {
        return if (operands[0].isReference) {
            PtgBool(true)
        } else PtgBool(false)
    }

    /**
     * N
     * Returns a value converted to a number.
     *
     *
     * Syntax
     *
     *
     * N(value)
     *
     *
     * Value   is the value you want converted. N converts values listed in the following table.
     *
     *
     * If value is or refers to
     * N returns
     *
     *
     * A number
     * That number
     *
     *
     * A date, in one of the built-in date formats available in Microsoft Excel
     * The serial number of that date --- Note that to us, this is just a number, the date
     * format is just that, a format.
     *
     *
     * TRUE
     * 1
     *
     *
     * Anything else
     * 0
     */
    internal fun calcN(operands: Array<Ptg>): Ptg {
        val o = operands[0].value
        if (o is Double || o is Int || o is Float || o is Long) {
            val d = Double(o.toString())
            return PtgNumber(d)
        }
        if (o is Boolean) {
            val bo = o.booleanValue()
            if (bo) return PtgInt(1)
        }
        return PtgInt(0)
    }

    /**
     * TYPE
     * Returns a number indicating the data type of a value
     * Value   can be any Microsoft Excel value, such as a number, text, logical value, and so on.
     * If value is TYPE returns
     * Number 1
     * Text 2
     * Logical value 4
     * Error value 16
     * Array 64
     */
    internal fun calcType(operands: Array<Ptg>): Ptg {
        if (operands[0] is PtgArray)
            return PtgNumber(64.0)    // avoid value calc for arrays
        else if (operands[0] is PtgErr)
            return PtgNumber(16.0)

        // otherwise, test value of operand
        val value = operands[0].value
        var type = 0
        if (value is String)
            type = 2
        else if (value is Number)
            type = 1
        else if (value is Boolean)
            type = 4
        return PtgNumber(type.toDouble())
    }

}

/*
 * known INFO function operating systems: 
 * TODO: need complete list
Windows Vista   Windows (32-bit) NT 6.00
Windows XP 		Windows (32-bit) NT 5.01
Windows2000 	Windows (32-bit) NT 5.00
Windows98 		Windows (32-bit) 4.10
Windows95 		Windows (32-bit) 4.00
*/

/*
Linux 	2.0.31 	x86 	IBM Java 1.3
Linux 	(*) 	i386 	Sun Java 1.3.1, 1.4 or Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	x86_64 	Blackdown Java; note x86_64 might change to amd64; (*) os.version depends on Linux Kernel version
Linux 	(*) 	sparc 	Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	ppc 	Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	armv41 	Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	i686 	GNU Java Compiler (GCJ); (*) os.version depends on Linux Kernel version
Linux 	(*) 	ppc64 	IBM Java 1.3; (*) os.version depends on Linux Kernel version
Mac OS 	7.5.1 	PowerPC 	
Mac OS 	8.1 	PowerPC 	
Mac OS 	9.0, 9.2.2 	PowerPC 	MacOS 9.0: java.version=1.1.8, mrj.version=2.2.5; MacOS 9.2.2: java.version=1.1.8 mrj.version=2.2.5
Mac OS X 	10.1.3 	ppc 	
Mac OS X 	10.2.6 	ppc 	Java(TM) 2 Runtime Environment, Standard Edition (build 1.4.1_01-39)
Java HotSpot(TM) Client VM (build 1.4.1_01-14, mixed mode)
Mac OS X 	10.2.8 	ppc 	using 1.3 JVM: java.vm.version=1.3.1_03-74, mrj.version=3.3.2; using 1.4 JVM: java.vm.version=1.4.1_01-24, mrj.version=69.1
Mac OS X 	10.3.1, 10.3.2, 10.3.3, 10.3.4 	ppc 	JDK 1.4.x
Mac OS X 	10.3.8 	ppc 	Mac OS X 10.3.8 Server; using 1.3 JVM: java.vm.version=1.3.1_03-76, mrj.version=3.3.3; using 1.4 JVM: java.vm.version=1.4.2-38; mrj.version=141.3
Windows 95 	4.0 	x86 	
Windows 98 	4.10 	x86 	Note, that if you run Sun JDK 1.2.1 or 1.2.2 Windows 98 identifies itself as Windows 95.
Windows Me 	4.90 	x86 	
Windows NT 	4.0 	x86 	
Windows 2000 	5.0 	x86 	
Windows XP 	5.1 	x86 	Note, that if you run older Java runtimes Windows XP identifies itself as Windows 2000.
Windows 2003 	5.2 	x86 	java.vm.version=1.4.2_06-b03; Note, that Windows Server 2003 identifies itself only as Windows 2003.
Windows CE 	3.0 build 11171 	arm 	Compaq iPAQ 3950 (PocketPC 2002)
OS/2 	20.40 	x86 	
Solaris 	2.x 	sparc 	
SunOS 	5.7 	sparc 	Sun Ultra 5 running Solaris 2.7
SunOS 	5.8 	sparc 	Sun Ultra 2 running Solaris 8
SunOS 	5.9 	sparc 	Java(TM) 2 Runtime Environment, Standard Edition (build 1.4.0_01-b03)
Java HotSpot(TM) Client VM (build 1.4.0_01-b03, mixed mode)
MPE/iX 	C.55.00 	PA-RISC 	
HP-UX 	B.10.20 	PA-RISC 	JDK 1.1.x
HP-UX 	B.11.00 	PA-RISC 	JDK 1.1.x
HP-UX 	B.11.11 	PA-RISC 	JDK 1.1.x
HP-UX 	B.11.11 	PA_RISC 	JDK 1.2.x/1.3.x; note Java 2 returns PA_RISC and Java 1 returns PA-RISC
HP-UX 	B.11.00 	PA_RISC 	JDK 1.2.x/1.3.x
HP-UX 	B.11.23 	IA64N 	JDK 1.4.x
HP-UX 	B.11.11 	PA_RISC2.0 	JDK 1.3.x or JDK 1.4.x, when run on a PA-RISC 2.0 system
HP-UX 	B.11.11 	PA_RISC 	JDK 1.2.x, even when run on a PA-RISC 2.0 system
HP-UX 	B.11.11 	PA-RISC 	JDK 1.1.x, even when run on a PA-RISC 2.0 system
AIX 	5.2 	ppc64 	sun.arch.data.model=64
AIX 	4.3 	Power 	
AIX 	4.1 	POWER_RS 	
OS/390 	390 	02.10.00 	J2RE 1.3.1 IBM OS/390 Persistent Reusable VM
FreeBSD 	2.2.2-RELEASE 	x86 	
Irix 	6.3 	mips 	
Digital Unix 	4.0 	alpha 	
NetWare 4.11 	4.11 	x86 	
OSF1 	V5.1 	alpha 	Java 1.3.1 on Compaq (now HP) Tru64 Unix V5.1
OpenVMS 	V7.2-1 	alpha 	Java 1.3.1_1 on OpenVMS 7.2
*/