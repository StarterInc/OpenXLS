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

import io.starter.OpenXLS.DateConverter
import io.starter.OpenXLS.ExcelTools
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.FormatConstants
import io.starter.formats.XLS.Formula
import io.starter.formats.XLS.XLSConstants
import io.starter.formats.XLS.Xf
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool

import java.io.UnsupportedEncodingException
import java.text.DecimalFormat
import java.text.Format
import java.util.Date

/**
 * TextCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 *
 *
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 */
object TextCalculator {

    /**
     * ASC function
     * For Double-byte character set (DBCS) languages, changes full-width (double-byte) characters to half-width (single-byte) characters.
     * ASC(text)
     * Text   is the text or a reference to a cell that contains the text you want to change. If text does not contain any full-width letters, text is not changed.
     * NOTE: in order to use this and other DBCS Methods in Excel,
     * the input language must be set to a DBCS language such as Japanese
     * Otherwise, the ASC function does nothing (apparently)
     */
    internal fun calcAsc(operands: Array<Ptg>?): Ptg {
        if (operands == null || operands[0] == null)
            return PtgErr(PtgErr.ERROR_VALUE)
        // determine if Excel's language is set up for DBCS; if not, returns normal string
        val bk = operands[0].parentRec.workBook
        if (bk!!.defaultLanguageIsDBCS()) { // otherwise just returns normal string
            var strbytes = getUnicodeBytesFromOp(operands[0])
            if (strbytes == null)
                strbytes = operands[0].value.toString().toByteArray()
            try {
                return PtgStr(String(strbytes!!, XLSConstants.UNICODEENCODING))
            } catch (e: Exception) {
            }

        }
        return PtgStr(operands[0].value.toString())
    }

    /*BAHTTEXT function
	Converts a number to Thai text and adds a suffix of "Baht."
	 */

    /**
     * CHAR
     * Returns the character specified by the code number
     */
    internal fun calcChar(operands: Array<Ptg>): Ptg {
        val o = operands[0].value
        val s = Byte(o.toString())
        if (s.toInt() > 255 || s.toInt() < 1) return PtgCalculator.error
        val b = ByteArray(1)
        b[0] = s
        var str = ""
        try {
            str = String(b, XLSConstants.DEFAULTENCODING)
        } catch (e: UnsupportedEncodingException) {
        }

        return PtgStr(str)
    }

    /**
     * CLEAN
     * Removes all nonprintable characters from text. Use CLEAN on text
     * imported from other applications that contains characters that may
     * not print with your operating system. For example, you can use
     * CLEAN to remove some low-level computer code that is frequently
     * at the beginning and end of data files and cannot be printed.
     *
     *
     * Syntax
     *
     *
     * CLEAN(text)
     *
     *
     * Text   is any worksheet information from which you want to remove nonprintable characters.
     *
     *
     * The CLEAN function was designed to remove the first 32 nonprinting characters in the 7-bit ASCII code (values 0 through 31) from text. In the Unicode character set (Unicode: A character encoding standard developed by the Unicode Consortium. By using more than one byte to represent each character, Unicode enables almost all of the written languages in the world to be represented by using a single character set.), there are additional nonprinting characters (values 127, 129, 141, 143, 144, and 157). By itself, the CLEAN function does not remove these additional nonprinting characters.
     */
    internal fun calcClean(operands: Array<Ptg>): Ptg {
        var retString = ""
        try {
            val o = operands[0].value
            val s = o.toString()
            for (i in 0 until s.length) {
                val c = s[i].toInt()
                if (c >= 32)
                    retString += c.toChar()
            }
        } catch (e: Exception) {
        }

        return PtgStr(retString)
    }


    /**
     * CODE
     * Returns a numeric code for the first character in a text string
     */
    internal fun calcCode(operands: Array<Ptg>): Ptg {
        val o = operands[0].value
        val s = o.toString()
        var b: ByteArray? = null
        try {
            b = s.toByteArray(charset(XLSConstants.DEFAULTENCODING))
        } catch (e: UnsupportedEncodingException) {
        }

        val i = Integer.valueOf(b!![0].toInt())
        return PtgInt(i!!.toInt())
    }


    /**
     * CONCATENATE
     * Joins several text items into one text item
     */
    internal fun calcConcatenate(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        val allops = PtgCalculator.getAllComponents(operands)
        var s = ""
        for (i in allops.indices) {
            s += allops[i].value.toString()
        }
        val str = PtgStr(s)
        str.parentRec = operands[0].parentRec
        return str
    }

    /**
     * DOLLAR
     * Converts a number to text, using currency format. Can
     * have a separate operand to determine POP.
     */
    internal fun calcDollar(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        var pop = 0
        if (operands.size > 1) {
            pop = operands[1].intVal
        }
        var d = operands[0].doubleVal
        d = d * Math.pow(10.0, pop.toDouble())
        d = Math.round(d).toDouble()
        d = d / Math.pow(10.0, pop.toDouble())
        val res = "$$d"
        return PtgStr(res)
    }

    /**
     * EXACT
     * Checks to see if two text values are identical
     */
    internal fun calcExact(operands: Array<Ptg>): Ptg {
        if (operands.size != 2) return PtgCalculator.error
        val s1 = operands[0].value.toString()
        val s2 = operands[1].value.toString()
        return if (s1 == s2) PtgBool(true) else PtgBool(false)
    }

    /**
     * FIND
     * Finds one text value within another (case-sensitive)
     */
    internal fun calcFind(operands: Array<Ptg>): Ptg {
        var instring = ""
        var wholestr = ""
        var start = 0
        if (operands.size < 2) return PtgErr(PtgErr.ERROR_VALUE)
        if (operands.size == 3) {
            start = operands[2].intVal - 1
        }
        val o = operands[0].value
        val oo = operands[1].value

        if (o == null || oo == null) return PtgErr(PtgErr.ERROR_VALUE)

        instring = o.toString()
        wholestr = oo.toString()
        // note this uses a starting position to search for the string,
        // but does not account for that starting position in respects to it's
        // result.  Pretty strange
        var i = wholestr.indexOf(instring, start)
        if (i != -1) {
            i = wholestr.indexOf(instring)
            i++
            return PtgInt(i)
        }
        return PtgErr(PtgErr.ERROR_VALUE)

    }

    /**
     * FINDB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, FINDB counts each character as 1.
     * NOTES:  search is case sensitive and doesn't allow for wildcards
     */
    internal fun calcFindB(operands: Array<Ptg>?): Ptg {
        if (operands == null || operands.size < 2 || operands[0] == null)
            return PtgErr(PtgErr.ERROR_VALUE)
        // determine if Excel's language is set up for DBCS; if not, returns normal string
        val bk = operands[0].parentRec.workBook
        if (!bk!!.defaultLanguageIsDBCS()) { // otherwise just use calcFind
            return calcFind(operands)
        }
        var startnum = 0
        if (operands.size > 2)
            startnum = operands[2].intVal
        val strToFind = getUnicodeBytesFromOp(operands[0])
        val str = getUnicodeBytesFromOp(operands[1])
        var index = -1
        if (strToFind == null || strToFind.size == 0 || str == null || startnum < 0 || str.size < startnum)
            return PtgInt(startnum)
        var i = startnum
        while (i < str.size && index == -1) {
            if (strToFind[0] == str[i]) {
                index = i
                var j = 0
                while (j < strToFind.size && i + j < str.size && index == i) {
                    if (strToFind[j] != str[i + j]) {
                        index = -1 // start over
                        break
                    }
                    j++
                }
            }
            i++
        }
        if (index == -1)
        // not found
            PtgErr(PtgErr.ERROR_VALUE)
        return PtgInt(index + 1)    // return 1-based index of found bytes
    }

    /**
     * FIXED
     * Formats a number as text with a fixed number of decimals
     */
    internal fun calcFixed(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        var nocommas = false
        if (operands.size == 3) {
            val boo = operands[2].value as Boolean
            nocommas = boo.booleanValue()
        }
        var dub = operands[0].doubleVal
        if (dub == java.lang.Double.NaN) dub = 0.0
        val pop = operands[1].intVal
        dub = dub * Math.pow(10.0, pop.toDouble())
        dub = Math.round(dub).toDouble()
        dub = dub / Math.pow(10.0, pop.toDouble())
        var res = dub.toString()
        if (pop == 0) {
            if (res.indexOf(".") > -1) {
                res = res.substring(0, res.indexOf("."))
                return PtgStr(res)
            }
        }
        // pad w/zeros if need be.
        if (res.indexOf(".") == -1 && pop > 0) {
            res = "$res.0"
        }
        var mantissa = res.substring(res.indexOf("."))
        while (mantissa.length <= pop) {
            res += 0
            mantissa = res.substring(res.indexOf("."))
        }
        if (nocommas || dub < 999.99) {
            return PtgStr(res)
        }

        val e = res.indexOf(".")
        var mant = res.substring(e)
        val begin = res.substring(0, e)
        var counter = 0
        val s = begin.length
        // this adds the commas;
        var v = 0
        while (v < s) {
            val ch = begin.substring(s - v - 1, s - v)
            mant = ch + mant
            v++
            if (counter == 2 && v != s) mant = ",$mant"
            counter++
            if (counter == 3) counter = 0
        }
        return PtgStr(mant)

    }

    /**
     * JIS function
     * The function described in this Help topic converts half-width (single-byte)
     * letters within a character string to full-width (double-byte) characters.
     * The name of the function (and the characters that it converts) depends upon your language settings.
     * For Japanese, this function changes half-width (single-byte) English letters or
     * katakana within a character string to full-width (double-byte) characters.
     * JIS(text)
     * Text   is the text or a reference to a cell that contains the text you want to change. If text does not contain any half-width English letters or katakana, text is not changed.
     *
     *
     * TODO: STRING ENCODING IS NOT CORRECT **************
     */
    /*
  * encoding info:
	Shift_JIS	DBCS		16-bit Japanese encoding (Note that you must use an underscore character (_), not a hyphen (-) in the name in CFML attributes.)
	(same as MS932)
	EUC-KR		DBCS		16-bit Korean encoding
	UCS-2		DBCS		Two-byte Unicode encoding
	UTF-8		MBCS		Multibyte Unicode encoding. ASCII is 7-bit; non-ASCII characters used in European and many Middle Eastern languages are two-byte; and most Asian characters are three-byte
*/
    internal fun calcJIS(operands: Array<Ptg>?): Ptg {
        if (operands == null || operands[0] == null)
            return PtgErr(PtgErr.ERROR_VALUE)
        // determine if Excel's language is set up for DBCS; if not, returns normal string
        val bk = operands[0].parentRec.workBook
        if (bk!!.defaultLanguageIsDBCS()) { // otherwise just returns normal string
            var strbytes = getUnicodeBytesFromOp(operands[0])
            if (strbytes == null)
                strbytes = operands[0].value.toString().toByteArray()
            try {
                return PtgStr(String(strbytes!!, "Shift_JIS"))
            } catch (e: Exception) {
            }

        }
        return PtgStr(operands[0].value.toString())
    }

    /**
     * LEFT
     * Returns the leftmost characters from a text value
     */
    internal fun calcLeft(operands: Array<Ptg>): Ptg {
        var numchars = 1
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        if (operands[0] is PtgErr)
            return PtgErr(PtgErr.ERROR_NA)    // 'tis what excel does
        if (operands.size == 2) {
            if (operands[1] is PtgErr)
                return PtgErr(PtgErr.ERROR_VALUE)
            numchars = operands[1].intVal
        }
        val o = operands[0].value ?: return PtgStr("")
        val str = o.toString()
        if (str == null || numchars > str.length)
            return PtgStr("")    // 20081202 KSC: Don't error out if not enough chars ala Excel
        val res = str.substring(0, numchars)
        return PtgStr(res)
    }

    /**
     * LEFTB counts each double-byte character as 2 when you have enabled the editing of a
     * language that supports DBCS and then set it as the default language.
     * Otherwise, LEFTB counts each character as 1.
     */
    internal fun calcLeftB(operands: Array<Ptg>): Ptg {
        val bk = operands[0].parentRec.workBook
        if (bk!!.defaultLanguageIsDBCS()) {// otherwise just returns normal string
            var numchars = 1
            if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
            if (operands.size == 2) {
                if (operands[1] is PtgErr)
                    return PtgErr(PtgErr.ERROR_VALUE)
                try {
                    numchars = operands[1].intVal
                    val b = ByteArray(numchars)
                    System.arraycopy(getUnicodeBytesFromOp(operands[0])!!, 0, b, 0, numchars)
                    return PtgStr(String(b, XLSConstants.UNICODEENCODING))
                } catch (e: Exception) {
                    return PtgErr(PtgErr.ERROR_VALUE)
                }

            }
        }
        return calcLeft(operands)
    }

    /**
     * LEN
     * Returns the number of characters in a text string
     */
    internal fun calcLen(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        val s = operands[0].value.toString()
        return PtgInt(s.length)
    }

    /**
     * LENB counts each double-byte character as 2 when you have enabled the editing of
     * a language that supports DBCS and then set it as the default language.
     * Otherwise, LENB counts each character as 1.
     */
    internal fun calcLenB(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        val bk = operands[0].parentRec.workBook
        if (bk!!.defaultLanguageIsDBCS())
        // otherwise just returns normal string
            return PtgInt(getUnicodeBytesFromOp(operands[0])!!.size)
        val s = operands[0].value.toString()
        return PtgInt(s.length)
    }

    /**
     * LOWER
     * Converts text to lowercase
     */
    internal fun calcLower(operands: Array<Ptg>): Ptg {
        if (operands.size > 1) return PtgCalculator.error
        var s = operands[0].value.toString()
        s = s.toLowerCase()
        return PtgStr(s)
    }

    /**
     * MID
     * Returns a specific number of characters from a text string starting at the position you specify
     */
    internal fun calcMid(operands: Array<Ptg>): Ptg {
        var s: String? = operands[0].value.toString()
        if (s == null || s == "") return PtgStr("")    //  Don't error out if "" ala Excel
        if (operands[1] is PtgErr || operands[2] is PtgErr) return PtgErr(PtgErr.ERROR_VALUE)
        val start = operands[1].intVal - 1
        var len = operands[2].intVal
        if (len < 0) {
            len = start + len
        }
        if (s.length < start) return PtgStr("")
        if (start == -1) return PtgErr(PtgErr.ERROR_VALUE)
        s = s.substring(start)
        if (len > s.length) return PtgStr(s)
        s = s.substring(0, len)
        return PtgStr(s)
    }
    /*
     * MIDB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, MIDB counts each character as 1.
     */
    /*
  * PHONETIC function
	Extracts the phonetic (furigana) characters from a text string.
	PHONETIC(reference)
	Reference   is a text string or a reference to a single cell or a range of cells that contain a furigana text string.
  */

    /**
     * PROPER
     * Capitalizes the first letter in each word of a text value
     */
    internal fun calcProper(operands: Array<Ptg>): Ptg {
        var s = operands[0].value.toString()
        s = StringTool.proper(s)
        return PtgStr(s)
    }

    /**
     * REPLACE
     * Replaces characters within text
     */
    internal fun calcReplace(operands: Array<Ptg>): Ptg {
        val origstr = operands[0].value.toString()
        val start = operands[1].intVal
        val repamount = operands[2].intVal
        val repstr = operands[3].value.toString()
        val begin = origstr.substring(0, start - 1)
        val end = origstr.substring(start + repamount - 1)
        val returnstr = begin + repstr + end
        return PtgStr(returnstr)

    }
    /*
     * REPLACEB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, REPLACEB counts each character as 1.
     */

    /**
     * REPT
     * Repeats text a given number of times
     */
    internal fun calcRept(operands: Array<Ptg>): Ptg {
        val origstr = operands[0].value.toString()
        val numtimes = operands[1].intVal
        var retstr = ""
        for (i in 0 until numtimes) {
            retstr += origstr
        }
        return PtgStr(retstr)
    }

    /**
     * RIGHT
     * Returns the rightmost characters from a text value
     */
    internal fun calcRight(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        val origstr = operands[0].value.toString()
        if (origstr == "") return PtgStr("")
        var numchars = operands[1].intVal
        if (numchars > origstr.length) numchars = origstr.length
        if (numchars < 0) return PtgErr(PtgErr.ERROR_VALUE)
        val res = origstr.substring(origstr.length - numchars)
        return PtgStr(res)
    }
    /*
     * RIGHTB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, RIGHTB counts each character as 1.
     */

    /**
     * SEARCH
     * Finds one text value within another (not case-sensitive)
     */
    internal fun calcSearch(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) return PtgErr(PtgErr.ERROR_VALUE)
        var start = 0
        if (operands.size == 3) {
            start = operands[2].intVal - 1
        }
        val search = operands[0].value.toString().toLowerCase()
        val orig = operands[1].value.toString().toLowerCase()
        val tmp = orig.substring(start).toLowerCase()
        var i = tmp.indexOf(search)
        if (i == -1) return PtgErr(PtgErr.ERROR_VALUE)
        i = orig.indexOf(search)
        i++
        return PtgInt(i)
    }

    /**
     * SEARCHB counts each double-byte character as 2 when you have enabled the editing of a
     * language that supports DBCS and then set it as the default language.
     * Otherwise, SEARCHB counts each character as 1.
     *
     *
     * TODO: THIS IS NOT COMPLETE
     */
    internal fun calcSearchB(operands: Array<Ptg>?): Ptg {
        if (operands == null || operands.size < 2 || operands[0] == null)
            return PtgErr(PtgErr.ERROR_VALUE)
        // determine if Excel's language is set up for DBCS; if not, returns normal string
        val bk = operands[0].parentRec.workBook
        if (!bk!!.defaultLanguageIsDBCS()) { // otherwise just use calcFind
            return calcSearch(operands)
        }
        var startnum = 0
        if (operands.size > 2)
            startnum = operands[2].intVal
        val strToFind = getUnicodeBytesFromOp(operands[0])
        val str = getUnicodeBytesFromOp(operands[1])
        var index = -1
        if (strToFind == null || strToFind.size == 0 || str == null || startnum < 0 || str.size < startnum)
            return PtgInt(startnum)

        val search = operands[0].value.toString().toLowerCase()
        val orig = operands[1].value.toString().toLowerCase()
        val tmp = orig.substring(startnum).toLowerCase()
        index = tmp.indexOf(search)

        if (index == -1)
        // not found
            PtgErr(PtgErr.ERROR_VALUE)
        else
            index *= 2 // count the bytes as double
        return PtgInt(index + 1)    // return 1-based index of found bytes
    }

    /**
     * SUBSTITUTE
     * Substitutes new text for old text in a text string
     */
    internal fun calcSubstitute(operands: Array<Ptg>): Ptg {
        var whichreplace = 0
        if (operands.size < 3) return PtgErr(PtgErr.ERROR_VALUE)
        if (operands.size == 4)
            whichreplace = operands[3].intVal - 1
        val origstr = operands[0].value.toString()
        val srchstr = operands[1].value.toString()
        val repstr = operands[2].value.toString()
        val finalstr = StringTool.replaceText(origstr, srchstr, repstr, whichreplace, true)
        return PtgStr(finalstr)
    }

    /**
     * T
     * According to documentation converts its arguments to text -
     *
     *
     * not really though, it just returns value if they are text
     */
    internal fun calcT(operands: Array<Ptg>): Ptg {
        var res = ""
        try {
            res = operands[0].value as String
        } catch (e: ClassCastException) {
        }

        return PtgStr(res)

    }

    /**
     * TEXT
     * Formats a number and converts it to text
     *
     *
     * Converts a value to text in a specific number format.
     *
     *
     * Syntax
     *
     *
     * TEXT(value,format_text)
     *
     *
     * Value   is a numeric value, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value.
     *
     *
     * Format_text   is a number format in text form from in the Category box on the Number tab in the Format Cells dialog box.
     *
     *
     * Remarks
     *
     *
     * Format_text cannot contain an asterisk (*).
     *
     *
     * Formatting a cell with an option on the Number tab (Cells command, Format menu) changes only the format, not the value. Using the TEXT function converts a value to formatted text, and the result is no longer calculated as a number.
     *
     *
     * Salesperson Sales
     * Buchanan 2800
     * Dodsworth 40%
     *
     *
     * Formula Description (Result)
     * =A2&" sold "&TEXT(B2, "$0.00")&" worth of units." Combines contents above into a phrase (Buchanan sold $2800.00 worth of units.)
     * =A3&" sold "&TEXT(B3,"0%")&" of the total sales." Combines contents above into a phrase (Dodsworth sold 40% of the total sales.)
     */
    internal fun calcText(operands: Array<Ptg>): Ptg {
        if (operands.size != 2)
            return PtgErr(PtgErr.ERROR_VALUE)
        var res: String? = "#ERR!"
        try {
            res = operands[0].value.toString()
        } catch (e: Exception) {
            res = operands[0].toString()
        }

        var fmt = operands[1].toString()
        var fmtx: Format? = null
        // convert a string like "0"
        // to a format pattern like: "##";
        for (t in FormatConstants.NUMERIC_FORMATS.indices) {
            val fmx = FormatConstants.NUMERIC_FORMATS[t][0]
            if (fmx == fmt) {
                fmt = FormatConstants.NUMERIC_FORMATS[t][2]
                fmtx = DecimalFormat(fmt)
            }
        }
        if (fmtx == null) {
            for (t in FormatConstants.CURRENCY_FORMATS.indices) {
                val fmx = FormatConstants.CURRENCY_FORMATS[t][0]
                if (fmx == fmt) {
                    fmt = FormatConstants.CURRENCY_FORMATS[t][2]
                    fmtx = DecimalFormat(fmt)
                }
            }
        }
        if (fmtx != null) {
            try {
                if (res != null && res != "")
                // 20090527 KSC: when cell=="", Excel treats as 0
                    return PtgStr(fmtx.format(Float(res)))
                else if (res != null)
                    return PtgStr(fmtx.format(0))
            } catch (e: Exception) {
                //	            Logger.logWarn("getting formatted string value for :" + res.toString() + " failed: " + e.toString()) ;
                try {
                    return PtgStr(res)    // 20080211 KSC: Double.valueOf(ret.toString());
                } catch (nbe: NumberFormatException) {
                    // who knew? - of course,  functions don't have to return numbers!
                }

            }

        }
        for (x in FormatConstants.DATE_FORMATS.indices) {
            val fmx = FormatConstants.DATE_FORMATS[x][0]
            if (fmx == fmt) {
                fmt = FormatConstants.DATE_FORMATS[x][2]

                try {
                    var d: Date?
                    try {
                        d = DateConverter.getDateFromNumber(Double(res!!))
                    } catch (e: NumberFormatException) {
                        d = DateConverter.getDate(res)    // try to convert a string date
                        if (d == null)
                        // what excel does, if it's an empty date, it reverts to jan 0, 1900
                            d = Date("1/1/1990")
                    }

                    //SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                    try {
                        //sdf = new SimpleDateFormat(fmt);
                        WorkBookHandle.simpledateformat.applyPattern(fmt)
                    } catch (ex: Exception) {
                        Logger.logWarn("Simple Date Format could not parse: $fmt. Returning default.") //not a valid date format
                    }

                    //return new PtgStr(sdf.format(d));
                    return PtgStr(WorkBookHandle.simpledateformat.format(d))
                } catch (e: Exception) {
                    Logger.logErr("Unable to calcText formatting correctly for a date$e")
                }

            }
        }

        // we've been unable to format, try based on the string
        try {
            if (Xf.isDatePattern(fmt)) {
                //fmtx = new SimpleDateFormat( fmt );
                WorkBookHandle.simpledateformat.applyPattern(fmt)
                fmtx = WorkBookHandle.simpledateformat
            } else
                fmtx = DecimalFormat(fmt)

            if (res != null && res != "")
                return PtgStr(fmtx!!.format(Float(res)))
            else if (res != null)
                return PtgStr(fmtx!!.format(0))
        } catch (e: Exception) {
            //Logger.logWarn("getting formatted string value for :" + res.toString() + " failed: " + e.toString()) ;
            try {
                return PtgStr(res)
            } catch (nbe: NumberFormatException) {
            }

        }

        return PtgStr(res)
    }

    /**
     * TRIM
     * According to documentation Trim() removes leading and trailing spaces from the cell value.
     *
     *
     * Actually it removes all spaces except for single spaces
     * between words.
     */
    internal fun calcTrim(operands: Array<Ptg>): Ptg {
        val o = operands[0].value
        var res: String?
        if (o is Double) {
            res = ExcelTools.getNumberAsString(o.toDouble())
        } else {
            res = o.toString()
        }
        if (res == null || res == PtgErr(PtgErr.ERROR_NA).toString())
            return PtgErr(PtgErr.ERROR_NA)
        // first let's remove the beginning and trailing spaces.
        if (res.length > 0) {
            while (res!!.substring(0, 1) == " ") {
                res = res.substring(1)
            }
            while (res!!.substring(res.length - 1) == " ") {
                res = res.substring(0, res.length - 1)
            }
            // now we need to remove double spaces
            while (res!!.indexOf("  ") != -1) {
                val i = res.indexOf("  ")
                val prestring = res.substring(0, i)
                val poststring = res.substring(i + 1)
                res = prestring + poststring
            }
        }
        return PtgStr(res)

    }

    /**
     * UPPER
     * Converts text to uppercase
     */
    internal fun calcUpper(operands: Array<Ptg>): Ptg {
        if (operands.size > 1) return PtgCalculator.error
        var s = operands[0].value.toString()
        s = s.toUpperCase()
        return PtgStr(s)
    }

    /**
     * VALUE
     * Converts a text argument to a number
     */
    internal fun calcValue(operands: Array<Ptg>): Ptg {
        try {
            var s = operands[0].value.toString()
            if (s == "") s = "0" // Excel returns a zero for a blank value if VALUE is called upon it.
            val d = Double(s)
            return PtgNumber(d)
        } catch (e: NumberFormatException) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    /**
     * helper method for all DBCS-related worksheet functions
     *
     * @param op
     * @return
     */
    private fun getUnicodeBytesFromOp(op: Ptg): ByteArray? {
        var strbytes: ByteArray? = null
        if (op is PtgRef) {
            val rec = op.refCells!![0]
            if (rec is io.starter.formats.XLS.Labelsst)
                strbytes = rec.unsharedString!!.readStr()
            else if (rec is Formula) {
                strbytes = op.value.toString().toByteArray()
            } else
            // DEBUGGING- Take out when done
                Logger.logWarn("getUnicodeBytes: Unexpected rec encountered: " + op.javaClass)
        } else if (op is PtgStr) {
            strbytes = ByteArray(op.record.size - 3)
            System.arraycopy(op.record, 3, strbytes, 0, strbytes.size)
        } else {
            // DEBUGGING- Take out when done
            Logger.logWarn("getUnicodeBytes: Unexpected operand encountered: " + op.javaClass)
        }
        return strbytes
    }
}