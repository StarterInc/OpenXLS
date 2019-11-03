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
package io.starter.formats.XLS

import io.starter.OpenXLS.DateConverter
import io.starter.OpenXLS.ExcelTools
import io.starter.OpenXLS.ValidationHandle
import io.starter.formats.XLS.formulas.FormulaParser
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgArea
import io.starter.formats.XLS.formulas.PtgRef
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList
import java.util.Calendar
import java.util.Date
import java.util.Stack


/**
 * **Dv: Data Validity Settings (01BEh)**<br></br>
 *
 *
 * This record is part of the Data Validity Table. It stores data validity settings and a list of cell ranges which
 * contain these settings. The prompt box appears while editing such a cell. The error box appears, if the entered value
 * does not fit the conditions. The data validity settings of a sheet are stored in a sequential list of DV records. This list is
 * preluded by an DVAL record. If a string is empty and the default text should appear in the prompt box or error
 * box, the string must contain a single zero character (string length will be 1).
 *
 * <pre>
 *
 *
 * Offset          Name              Size                Contents
 * --------------------------------------------------------
 * 0               dwDvFlags      4                   Option flags (see below)
 * 4               dTitlePrompt   var                 Title of the prompt box (Unicode string, 16-bit string length)
 * var.            dTitleError      var.                Title of the error box (Unicode string, 16-bit string length)
 * var.            dTextPrompt  var.                Text of the prompt box (Unicode string, 16-bit string length)
 * var.            dTextError      var.                Text of the error box (Unicode string, 16-bit string length)
 * var.            sz1                 2                    Size of the formula data for first condition (sz1)
 * var.            garbage          2                   Not used
 * var.            firstCond       sz1                Formula data for first condition (RPN token array without size field)
 * var.            sz2                 2                   Size of the formula data for second condition (sz2)
 * var.            garbage           2                   Not used
 * var.            secondCond sz2                  Formula data for second condition (RPN token array without size field)
 * var.            cRangeList    var.                 Cell range address list with all affected ranges
 *
 * Option flags field:
 * Bit             Mask                    Name                 Contents
 * --
 * 3-0         0000000FH           ValType             Data type: 00H = Any value
 * 01H = Integer values
 * 02H = Decimal values
 * 03H = User defined list
 * 04H = Date
 * 05H = Time
 * 06H = Text length
 * 07H = Formula
 * 6-4     00000070H               ErrStyle               Error style: 00H = Stop
 * 01H = Warning
 * 02H = Info
 * 7         00000080H              fStrLookup         1 = In list type validity the string list is explicitly given in the formula
 * 8         00000100H            fAllowBlank         1 = Empty cells allowed
 * 9         00000200H             fSuppressCombo 1 = Suppress the drop down arrow in list type validity
 * 18      00040000H             fShowInputMsg     1 = Show prompt box if cell selected
 * 19      00080000H             fShowErrorMsg      1 = Show error box if invalid values entered
 * 23-20 00F00000H             typOperator         Condition operator: 00H = Between
 * 01H = Not between
 * 02H = Equal
 * 03H = Not equal
 * 04H = Greater than
 * 05H = Less than
 * 06H = Greater or equal
 * 07H = Less or equal
 *
</pre> *
 * In list type validity it is possible to enter an explicit string list. This string list is stored as tStr token . The string
 * items are separated by zero characters. There is no zero character at the end of the string list.
 * Example for a string list with the 3 strings A, B, and C: A<00H>B<00H>C (contained in a tStr token, string
 * length is 5).
 */
class Dv : io.starter.formats.XLS.XLSRecord() {
    private var grbit: Int = 0
    private var dTitlePrompt: Unicodestring? = null
    private var dTitleError: Unicodestring? = null
    private var dTextPrompt: Unicodestring? = null
    private var dTextError: Unicodestring? = null
    private var firstCond: Stack<*>? = null
    private var secondCond: Stack<*>? = null
    private var cRangeList: ArrayList<*>? = null
    private val garbageByteOne = ByteArray(2)
    private val garbageByteTwo = ByteArray(2)
    internal var numLocs: Byte = 0
    //private byte[] garbageByteThree = new byte[1];
    // grbit (dwDvFlags) fields
    private var valType: Byte = 0
    private var errStyle: Byte = 0
    private var fStrLookup: Boolean = false
    private var fAllowBlank: Boolean = false
    private var fSuppressCombo: Boolean = false
    private var fShowInputMsg: Boolean = false
    /**
     * Show error box if invalid values entered?
     *
     * @return
     */
    var isShowErrorMsg: Boolean = false
        private set
    private var IMEMode: Short = 0
    private var typOperator: Byte = 0
    private var dirtyflag = false

    // 20090606: made prototype completely blank i.e. no prompt, error text or formulas ...
    //    private byte[] PROTOTYPE_BYTES = {0,   1,  12,   0,   0,   0,   0,  0,   0,   0,  0,   0,   0,  0,   0,   0,  0,   0,   0,   0,
    //      0,   0,  89,  84,  0 };//,   0,   0,   0,   0,   0,  0,   0,   0,   0};
    private val PROTOTYPE_BYTES = byteArrayOf(3, 1, 12, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0)

    // 20090609 KSC: need to store ranges separately as OOXML ranges addresses may exceed 2003 maximum size
    private var ooxmlranges: Array<String>? = null

    /**
     * Return the range of data this Dv refers to as a string array
     *
     *
     * Values are stored as absolute ($) references, but should be displayed
     * as relative
     *
     * @return ptgRef.toString()
     */
    val ranges: Array<String>
        get() {
            if (cRangeList == null && ooxmlranges != null) {
                if (ooxmlranges!!.size > 0)
                    return ooxmlranges
            }
            val s = arrayOfNulls<String>(cRangeList!!.size)
            for (i in s.indices) {
                s[i] = (cRangeList!![i] as PtgArea).location
                s[i] = StringTool.strip(s[i], "$")
            }
            return s

        }

    /**
     * Return the text in the error box
     *
     * @return
     */
    /**
     * Set the text for the error box
     *
     * @param textError
     */
    var errorBoxText: String
        get() = dTextError!!.toString().trim { it <= ' ' }
        set(textError) {
            dTextError!!.updateUnicodeString(textError)
            dirtyflag = true
        }

    /**
     * Return the text in the prompt box
     *
     * @return
     */
    /**
     * Set the text for the prompt box
     *
     * @param text
     */
    var promptBoxText: String
        get() = dTextPrompt!!.toString().trim { it <= ' ' }
        set(text) {
            dTextPrompt!!.updateUnicodeString(text)
            dirtyflag = true
        }

    /**
     * Return the title from the error box
     *
     * @return
     */
    /**
     * Set the title for the error box
     *
     * @param textError
     */
    var errorBoxTitle: String
        get() = dTitleError!!.toString().trim { it <= ' ' }
        set(textError) {
            dTitleError!!.updateUnicodeString(textError)
            dirtyflag = true
        }

    /**
     * Return the title in the prompt box
     *
     * @return
     */
    /**
     * Set the title for the prompt box
     *
     * @param text
     */
    var promptBoxTitle: String
        get() = dTitlePrompt!!.toString().trim { it <= ' ' }
        set(text) {
            dTitlePrompt!!.updateUnicodeString(text)
            dirtyflag = true
        }


    /**
     * Return a byte representing the error style for this DV
     *
     *
     * These map to the static final ints ERROR_* from ValidationHandle
     *
     * @return
     */
    /**
     * Set the error style for this Dv record
     *
     *
     * These map to the static final ints ERROR_* from ValidationHandle
     *
     * @return
     */
    var errorStyle: Byte
        get() = errStyle
        set(errstyle) {
            errStyle = errstyle
            dirtyflag = true

        }

    /**
     * Allow blank cells in the validation area?
     *
     * @return
     */
    /**
     * Allow blank cells in the validation area?
     *
     * @return
     */
    var isAllowBlank: Boolean
        get() = fAllowBlank
        set(allowBlank) {
            fAllowBlank = allowBlank
            updateGrbit()
            dirtyflag = true
        }

    /**
     * Show prompt box if cell selected?
     *
     * @return
     */
    /**
     * Set show prompt box if cell selected?
     * *
     *
     * @return
     */
    var showInputMsg: Boolean
        get() = fShowInputMsg
        set(showInputMsg) {
            fShowInputMsg = showInputMsg
            dirtyflag = true
        }

    /**
     * In list type validity the string list is explicitly given in the formula
     *
     * @return
     */
    /**
     * In list type validity the string list is explicitly given in the formula
     *
     * @return
     */
    var isStrLookup: Boolean
        get() = fStrLookup
        set(strLookup) {
            fStrLookup = strLookup
            dirtyflag = true
        }

    /**
     * Suppress the drop down arrow in list type validity
     *
     * @return
     */
    /**
     * Suppress the drop down arrow in list type validity
     *
     * @return
     */
    var isSuppressCombo: Boolean
        get() = fSuppressCombo
        set(suppressCombo) {
            fSuppressCombo = suppressCombo
            dirtyflag = true
        }

    /**
     * Get the IME mode for this validation
     *
     * @return
     */
    /**
     * set the IME mode for this validation
     *
     * @return
     */
    var imeMode: Short
        get() = IMEMode
        set(mode) {
            IMEMode = mode
            dirtyflag = true
        }

    /**
     * Get the type operator of this validation as a byte.
     *
     *
     * These bytes map to the CONDITION_* static values in
     * ValidationHandle
     *
     * @return
     */
    /**
     * set the type operator of this validation as a byte.
     *
     *
     * These bytes map to the CONDITION_* static values in
     * ValidationHandle
     *
     * @return
     */
    var typeOperator: Byte
        get() = typOperator
        set(typOperator) {
            this.typOperator = typOperator
            dirtyflag = true
        }

    /**
     * generate the proper OOXML to define this Dv
     *
     * @return
     */
    // required
    // any ????
    // TODO: this maps to ???
    //ooxml.append(" type=");
    // default, leave out
    // ooxml.append(" operator=\"between\"");
    // default no need to outut ooxml.append(" errorStyle=\"stop\"");
    //TODO "imeMode"
    // This needs to be better thought out, currently it breaks/strips all changes made to the model, as ranges
    // are not automatically added to ooxml ranges.
    /**if (ooxmlranges!=null) {// have stored OOXML ranges
     * ooxml.append(" sqref=\"");
     * for (int i= 0; i < ooxmlranges.length; i++) {
     * if (i>0) ooxml.append(" ");
     * ooxml.append(ooxmlranges[i]);
     * }
     * ooxml.append("\"");
     * } else {	// 2003-style ranges
     *///}
    /*
	"showDropDown"
*/
    /**
     * imwMode
     * TODO: map options correctly!!  where is the info??
     */// nocontrol
    // DV Lists are delimited by 0 must replace with commas for OOXML use
    // DV Lists are delimited by 0 must replace with commas for OOXML use
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<dataValidation")
            when (valType) {
                0 -> {
                }
                1 -> ooxml.append(" type=\"whole\"")
                2 -> ooxml.append(" type=\"decimal\"")
                3 -> ooxml.append(" type=\"list\"")
                4 -> ooxml.append(" type=\"date\"")
                5 -> ooxml.append(" type=\"time\"")
                6 -> ooxml.append(" type=\"textLength\"")
                7 -> ooxml.append(" type=\"custom\"")
            }
            when (typOperator) {
                0 -> {
                }
                1 -> ooxml.append(" operator=\"notBetween\"")
                2 -> ooxml.append(" operator=\"equal\"")
                3 -> ooxml.append(" operator=\"notEqual\"")
                4 -> ooxml.append(" operator=\"greaterThan\"")
                5 -> ooxml.append(" operator=\"lessThan\"")
                6 -> ooxml.append(" operator=\"greaterThanOrEqual\"")
                7 -> ooxml.append(" operator=\"lessThanOrEqual\"")
            }
            when (errStyle) {
                0 -> {
                }
                1 -> ooxml.append(" errorStyle=\"warning\"")
                2 -> ooxml.append(" errorStyle=\"information\"")
            }
            if (this.errorBoxText != "")
                ooxml.append(" error=\"" + OOXMLAdapter.stripNonAscii(this.errorBoxText) + "\"")
            if (this.errorBoxTitle != "")
                ooxml.append(" errorTitle=\"" + OOXMLAdapter.stripNonAscii(this.errorBoxTitle) + "\"")
            if (this.promptBoxText != "")
                ooxml.append(" prompt=\"" + OOXMLAdapter.stripNonAscii(this.promptBoxText) + "\"")
            if (this.promptBoxTitle != "")
                ooxml.append(" promptTitle=\"" + OOXMLAdapter.stripNonAscii(this.promptBoxTitle) + "\"")
            val ranges = this.ranges
            if (ranges.size > 0) {
                ooxml.append(" sqref=\"")
                for (i in ranges.indices) {
                    if (i > 0) ooxml.append(" ")
                    ooxml.append(ranges[i])
                }
                ooxml.append("\"")
            }

            if (this.isAllowBlank) ooxml.append(" allowBlank=\"1\"")
            if (this.isShowErrorMsg) ooxml.append(" showErrorMessage=\"1\"")
            if (this.showInputMsg) ooxml.append(" showInputMessage=\"1\"")
            when (this.imeMode) {
                0 -> {
                }
                1 -> ooxml.append(" imeMode=\"off\"")
                2 -> ooxml.append(" imeMode=\"on\"")
                3 -> ooxml.append(" imeMode=\"disabled\"")
                4 -> ooxml.append(" imeMode=\"hiragana\"")
                5 -> ooxml.append(" imeMode=\"fullKatakana\"")
                6 -> ooxml.append(" imeMode=\"halfKatakana\"")
                7 -> ooxml.append(" imeMode=\"fullAlpha\"")
                8 -> ooxml.append(" imeMode=\"halfAlpha\"")
                9 -> ooxml.append(" imeMode=\"fullHangul\"")
                10 -> ooxml.append(" imeMode=\"halfHangul\"")
            }
            ooxml.append(">")
            var formula1: String? = this.getFirstCond()
            if (formula1 != null && formula1.length > 0) {
                formula1 = formula1.replace(0.toChar(), ',')
                ooxml.append("<formula1>$formula1</formula1>")
            }
            var formula2: String? = this.getSecondCond()
            if (formula2 != null && formula2.length > 0) {
                formula2 = formula2.replace(0.toChar(), ',')
                ooxml.append("<formula2>$formula2</formula2>")
            }
            ooxml.append("</dataValidation>")
            return ooxml.toString()
        }

    /**
     * Determine if the value passed in is valid for
     * this validation
     *
     * @param value
     * @return
     */
    @Throws(ValidationException::class)
    fun isValid(value: Any?): Boolean {
        var value: Any? = value ?: throw ValidationException(this.errorBoxTitle, this.errorBoxText)

        // TODO: look into whether null is ever a valid value.  for now we assume "no".

        if (!this.isCorrectDataType(value))
            throw ValidationException(this.errorBoxTitle, this.errorBoxText)
        if (value is Date) {
            val d = DateConverter.getXLSDateVal(value as java.util.Date?)
            value = d.toString() + ""
        }
        when (typOperator) {
            ValidationHandle.CONDITION_BETWEEN -> {
                if (isBetween(value!!)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }
            ValidationHandle.CONDITION_NOT_BETWEEN -> {
                if (isNotBetween(value)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }
            ValidationHandle.CONDITION_EQUAL -> {
                if (isEqual(value!!)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }
            ValidationHandle.CONDITION_GREATER_THAN -> {
                if (isGreaterThan(value!!)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }
            ValidationHandle.CONDITION_GREATER_OR_EQUAL -> {
                if (isGreaterOrEqual(value!!)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }
            ValidationHandle.CONDITION_LESS_OR_EQUAL -> {
                if (isLessOrEqual(value!!)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }

            ValidationHandle.CONDITION_LESS_THAN -> {
                if (isGreaterOrEqual(value!!)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }

            ValidationHandle.CONDITION_NOT_EQUAL -> {
                if (isNotEqual(value)) return true
                throw ValidationException(this.errorBoxTitle, this.errorBoxText)
            }
        }
        return true
    }

    /**
     * Validate that the passed in value is between the two values specified in the parameters
     *
     * @param value
     * @return
     */
    private fun isBetween(value: Any): Boolean {
        val s1 = StringTool.strip(FormulaParser.getExpressionString(firstCond!!), "=")
        val s2 = StringTool.strip(FormulaParser.getExpressionString(secondCond!!), "=")
        val formulaStr = "=and($value>$s1,$s2>$value)"
        try {
            val f = FormulaParser.getFormulaFromString(formulaStr, this.workBook!!.getWorkSheetByNumber(0), intArrayOf(1, 1))
            f.setCachedValue(null)
            val o = f.calculateFormula()
            if (o is Boolean) {
                return o.booleanValue()
            }
        } catch (e: Exception) {
            Logger.logErr("Error calculating formula in validation $e")
        }

        return false
    }

    /**
     * Validate that the passed in value is NOT between the two passed in values
     *
     * @param value
     * @return
     */
    private fun isNotBetween(value: Any): Boolean {
        return !isBetween(value)
    }

    /**
     * Validate that the passed in value is equivalant
     *
     * @param value
     * @return
     */
    private fun isEqual(value: Any): Boolean {
        val s1 = StringTool.strip(FormulaParser.getExpressionString(firstCond!!), "=")
        val formulaStr = "=($value=$s1)"
        try {
            val f = FormulaParser.getFormulaFromString(formulaStr, this.workBook!!.getWorkSheetByNumber(0), intArrayOf(1, 1))
            val o = f.calculateFormula()
            if (o is Boolean) {
                return o.booleanValue()
            }
        } catch (e: Exception) {
            Logger.logErr("Error calculating formula in validation $e")
        }

        return false
    }

    /**
     * Validate that the passed in value is NOT between the two passed in values
     *
     * @param value
     * @return
     */
    private fun isNotEqual(value: Any): Boolean {
        return !isEqual(value)
    }

    /**
     * Validate that the passed in value is greater than
     *
     * @param value
     * @return
     */
    private fun isGreaterThan(value: Any): Boolean {
        val s1 = StringTool.strip(FormulaParser.getExpressionString(firstCond!!), "=")
        val formulaStr = "=($value>$s1)"
        try {
            val f = FormulaParser.getFormulaFromString(formulaStr, this.workBook!!.getWorkSheetByNumber(0), intArrayOf(1, 1))
            val o = f.calculateFormula()
            if (o is Boolean) {
                return o.booleanValue()
            }
        } catch (e: Exception) {
            Logger.logErr("Error calculating formula in validation $e")
        }

        return false
    }

    /**
     * Validate that the passed in value is greater than
     *
     * @param value
     * @return
     */
    private fun isGreaterOrEqual(value: Any): Boolean {
        val s1 = StringTool.strip(FormulaParser.getExpressionString(firstCond!!), "=")
        val formulaStr = "=($value>=$s1)"
        try {
            val f = FormulaParser.getFormulaFromString(formulaStr, this.workBook!!.getWorkSheetByNumber(0), intArrayOf(1, 1))
            val o = f.calculateFormula()
            if (o is Boolean) {
                return o.booleanValue()
            }
        } catch (e: Exception) {
            Logger.logErr("Error calculating formula in validation $e")
        }

        return false
    }

    /**
     * Validate that the passed in value is greater than
     *
     * @param value
     * @return
     */
    private fun isLessThan(value: Any): Boolean {
        val s1 = StringTool.strip(FormulaParser.getExpressionString(firstCond!!), "=")
        val formulaStr = "=($value<$s1)"
        try {
            val f = FormulaParser.getFormulaFromString(formulaStr, this.workBook!!.getWorkSheetByNumber(0), intArrayOf(1, 1))
            val o = f.calculateFormula()
            if (o is Boolean) {
                return o.booleanValue()
            }
        } catch (e: Exception) {
            Logger.logErr("Error calculating formula in validation $e")
        }

        return false
    }

    /**
     * Validate that the passed in value is greater than
     *
     * @param value
     * @return
     */
    private fun isLessOrEqual(value: Any): Boolean {
        val s1 = StringTool.strip(FormulaParser.getExpressionString(firstCond!!), "=")
        val formulaStr = "=($value<=$s1)"
        try {
            val f = FormulaParser.getFormulaFromString(formulaStr, this.workBook!!.getWorkSheetByNumber(0), intArrayOf(1, 1))
            val o = f.calculateFormula()
            if (o is Boolean) {
                return o.booleanValue()
            }
        } catch (e: Exception) {
            Logger.logErr("Error calculating formula in validation $e")
        }

        return false
    }


    /**
     * Determines if the value passed in is of the
     * correct data type
     *
     * @return
     */
    fun isCorrectDataType(value: Any): Boolean {
        when (valType) {
            ValidationHandle.VALUE_ANY -> return true
            ValidationHandle.VALUE_DATE -> {
                if (value is Date) return true
                if (value is Calendar) return true
            }
            ValidationHandle.VALUE_DECIMAL -> {
                val possibleDec = value.toString()
                try {
                    Double(possibleDec)
                    return true
                } catch (e: NumberFormatException) {
                    return false
                }

                val possibleFormula = value.toString()
                if (possibleFormula.indexOf("=") == 0) return true
            }
            ValidationHandle.VALUE_FORMULA -> {
                val possibleFormula = value.toString()
                if (possibleFormula.indexOf("=") == 0) return true
            }
            ValidationHandle.VALUE_INTEGER -> {
                val possibleInt = value.toString()
                if (possibleInt.indexOf(".") > -1) return false
                try {
                    val l = Long(possibleInt)  // excel ints go past boundary of java ints, so use long
                    return true
                } catch (e: NumberFormatException) {
                    return false
                }

                return true  //TODO
            }
            ValidationHandle.VALUE_TEXT_LENGTH -> return true
            ValidationHandle.VALUE_TIME -> return true  // TODO
            ValidationHandle.VALUE_USER_DEFINED_LIST -> return true // TODO
        }
        return false
    }

    /**
     * Standard init method
     *
     * @see io.starter.formats.XLS.XLSRecord.init
     */
    override fun init() {
        super.init()

        var offset = 0
        grbit = ByteTools.readInt(this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++))
        var strLen = ByteTools.readShort(this.getByteAt(offset).toInt(), this.getByteAt(offset + 1).toInt())
        var strGrbit = this.getByteAt(offset + 2)
        if (strGrbit and 0x1 == 0x1) {
            strLen *= 2
        }
        strLen += 3
        var namebytes = this.getBytesAt(offset, strLen.toInt())
        offset += strLen.toInt()
        dTitlePrompt = Unicodestring()
        dTitlePrompt!!.init(namebytes!!, false)

        strLen = ByteTools.readShort(this.getByteAt(offset).toInt(), this.getByteAt(offset + 1).toInt())
        strGrbit = this.getByteAt(offset + 2)
        if (strGrbit and 0x1 == 0x1) {
            strLen *= 2
        }
        strLen += 3

        namebytes = this.getBytesAt(offset, strLen.toInt())
        offset += strLen.toInt()
        dTitleError = Unicodestring()
        dTitleError!!.init(namebytes!!, false)

        strLen = ByteTools.readShort(this.getByteAt(offset).toInt(), this.getByteAt(offset + 1).toInt())
        strGrbit = this.getByteAt(offset + 2)
        if (strGrbit and 0x1 == 0x1) {
            strLen *= 2
        }
        strLen += 3
        namebytes = this.getBytesAt(offset, strLen.toInt())
        offset += strLen.toInt()
        dTextPrompt = Unicodestring()
        dTextPrompt!!.init(namebytes!!, false)

        strLen = ByteTools.readShort(this.getByteAt(offset).toInt(), this.getByteAt(offset + 1).toInt())
        strGrbit = this.getByteAt(offset + 2)
        if (strGrbit and 0x1 == 0x1) {
            strLen *= 2
        }
        strLen += 3
        namebytes = this.getBytesAt(offset, strLen.toInt())
        offset += strLen.toInt()
        dTextError = Unicodestring()
        dTextError!!.init(namebytes!!, false)

        val sz1 = ByteTools.readShort(this.getByteAt(offset++).toInt(), this.getByteAt(offset++).toInt()).toInt()
        // unknown bytes
        garbageByteOne[0] = this.getByteAt(offset++)
        garbageByteOne[1] = this.getByteAt(offset++)
        var formulaBytes = this.getBytesAt(offset, sz1)
        firstCond = ExpressionParser.parseExpression(formulaBytes, this)
        offset += sz1

        val sz2 = ByteTools.readShort(this.getByteAt(offset++).toInt(), this.getByteAt(offset++).toInt()).toInt()
        // unknown bytes
        garbageByteTwo[0] = this.getByteAt(offset++)
        garbageByteTwo[1] = this.getByteAt(offset++)
        formulaBytes = this.getBytesAt(offset, sz2)
        secondCond = ExpressionParser.parseExpression(formulaBytes, this)
        offset += sz2


        numLocs = this.getByteAt(offset++)
        cRangeList = ArrayList()
        for (i in 0 until numLocs) {
            var b = ByteArray(1)
            b[0] = 0x0
            b = ByteTools.append(b, this.getBytesAt(offset, 8))
            val p = PtgArea(false)
            p.parentRec = this
            p.init(b)
            cRangeList!!.add(p)
            offset += 8
        }

        // set all the grbit fields
        valType = (grbit and BITMASK_VALTYPE).toByte()
        errStyle = (grbit and BITMASK_ERRSTYLE shr 4).toByte()
        IMEMode = (grbit and BITMASK_MDIMEMODE shr 10).toShort()
        fStrLookup = grbit and BITMASK_FSTRLOOKUP == BITMASK_FSTRLOOKUP
        fAllowBlank = grbit and BITMASK_FALLOWBLANK == BITMASK_FALLOWBLANK
        fSuppressCombo = grbit and BITMASK_FSUPRESSCOMBO == BITMASK_FSUPRESSCOMBO
        fShowInputMsg = BITMASK_FSHOWINPUTMSG == BITMASK_FSHOWINPUTMSG
        isShowErrorMsg = grbit and BITMASK_FSHOWERRORMSG == BITMASK_FSHOWERRORMSG
        typOperator = (grbit and BITMASK_TYPOPERATOR shr 20).toByte()
    }

    /**
     * As most of these records have variable lengths we cannot just update part of
     * the data for the record at a time,  we just don't want to keep up updates.  Rather than this,
     * we'll update the entire record.  To keep processing down, update the record before streaming
     * rather than on each internal record change.
     */
    private fun updateRecord() {
        this.updateGrbit()
        var recbytes = ByteArray(0)

        var tmp = ByteTools.cLongToLEBytes(grbit)
        recbytes = ByteTools.append(tmp, recbytes)
        recbytes = ByteTools.append(dTitlePrompt!!.read(), recbytes)
        recbytes = ByteTools.append(dTitleError!!.read(), recbytes)
        recbytes = ByteTools.append(dTextPrompt!!.read(), recbytes)
        recbytes = ByteTools.append(dTextError!!.read(), recbytes)

        // get the firstCond bytes
        tmp = ByteArray(0)
        for (i in firstCond!!.indices) {
            val o = firstCond!!.elementAt(i)
            val ptg = o as Ptg
            tmp = ByteTools.append(ptg.record, tmp)
        }
        // get the length and add in.
        var sz = tmp.size.toShort()
        recbytes = ByteTools.append(ByteTools.shortToLEBytes(sz), recbytes)
        // add garbage
        recbytes = ByteTools.append(garbageByteOne, recbytes)
        recbytes = ByteTools.append(tmp, recbytes)

        // get the secondCond bytes
        tmp = ByteArray(0)
        for (i in secondCond!!.indices) {
            val o = secondCond!!.elementAt(i)
            val ptg = o as Ptg
            tmp = ByteTools.append(ptg.record, tmp)
        }
        // get the length and add in.
        sz = tmp.size.toShort()
        recbytes = ByteTools.append(ByteTools.shortToLEBytes(sz), recbytes)
        // add garbage
        recbytes = ByteTools.append(garbageByteTwo, recbytes)
        recbytes = ByteTools.append(tmp, recbytes)

        tmp = ByteArray(1)
        if (cRangeList != null) {
            tmp[0] = cRangeList!!.size.toByte()
            recbytes = ByteTools.append(tmp, recbytes)
            for (i in cRangeList!!.indices) {
                tmp = (cRangeList!![i] as PtgArea).record
                val tmp2 = ByteArray(8)
                tmp[0] = 0
                System.arraycopy(tmp, 0, tmp2, 0, tmp2.size)
                recbytes = ByteTools.append(tmp2, recbytes)
            }
            // there is a trailing zero, not sure why...
            if (cRangeList!!.size > 0) {
                tmp = ByteArray(1)
                tmp[0] = 0
                recbytes = ByteTools.append(tmp, recbytes)
            }
        } else if (ooxmlranges != null && ooxmlranges!!.size > 0) {
            tmp[0] = ooxmlranges!!.size.toByte()
            recbytes = ByteTools.append(tmp, recbytes)
            for (i in ooxmlranges!!.indices) {
                val /*Ref*/ p = PtgRef.createPtgRefFromString(this.sheet!!.sheetName + "!" + ooxmlranges!![i], this)
                tmp = p.record
/* replace with above PtgArea pa= new PtgArea();
                    try {
                        pa.setParentRec(this);
                        pa.setLocation(this.getSheet().getSheetName() + "!" + ooxmlranges[i]);
                    } catch (Exception e) {
                        // TODO: handle MAXROWS/MAXCOLS
                    }
                    tmp = pa.getRecord();
                    /**/
                tmp[0] = 0
val tmp2 = ByteArray(8)
System.arraycopy(tmp, 0, tmp2, 0, tmp2.size)
recbytes = ByteTools.append(tmp, recbytes)
}
 // there is a trailing zero, not sure why...
            if (ooxmlranges!!.size > 0)
{
tmp = ByteArray(1)
tmp[0] = 0
recbytes = ByteTools.append(tmp, recbytes)
}
}

this.setData(recbytes)
}


/**
 * update record.
 *
 * @see io.starter.formats.XLS.XLSRecord.preStream
 */
    public override fun preStream() {
if (dirtyflag) this.updateRecord()
}

/**
 * Apply all the grbit fields into the current grbit int
 */
     fun updateGrbit() {
grbit = 0
grbit = grbit or valType.toInt()
grbit = grbit or (errStyle shl 4)
grbit = grbit or (IMEMode shl 10)
if (fStrLookup)
grbit = (grbit or BITMASK_FSTRLOOKUP)

if (fAllowBlank)
grbit = (grbit or BITMASK_FALLOWBLANK)

if (fSuppressCombo)
grbit = (grbit or BITMASK_FSUPRESSCOMBO)

if (fShowInputMsg)
grbit = (grbit or BITMASK_FSHOWINPUTMSG)

if (isShowErrorMsg)
grbit = (grbit or BITMASK_FSHOWERRORMSG)

grbit = grbit or (typOperator shl 20)
}


/**
 * Set the range this Dv refers to.   Pass in a range string, sans worksheet
 * Note that absolute ranges/ptrgrefs are always used, however returning
 * values should not include the dollar sign
 *
 * @param range
 */
     fun setRange(range:String?) {
var range = range
if (range == null)
{    // for creating a dv and adding range info later
cRangeList = null
return
}
if (range!!.indexOf(":") == -1) range = range + ":" + range
val p = PtgArea(range, this, false)
cRangeList = ArrayList()
cRangeList!!.add(p)
dirtyflag = true
}

/**
 * Add a range this Dv refers to.   Pass in a range string, sans worksheet
 *
 * @param range
 */
     fun addRange(range:String) {
if (cRangeList == null) cRangeList = ArrayList()    // 20090605 KSC: Added
 /*int[] i = ExcelTools.getRowColFromString(range);
     	if(i.length==2) {

     	}else {*/
        val p = PtgArea(range, this, false)    // 20090609 KSC: absolute refs if '$' -really should test if row or col
cRangeList!!.add(p)
dirtyflag = true
}

/**
 * Add a range this Dv refers to.   Pass in a range string, sans worksheet
 * May need additional handling for records outside bounds?
 *
 * @param range
 */
     fun addOoxmlRange(range:String?) {
if (cRangeList == null) cRangeList = ArrayList()  // 20090605 KSC: Added
val p = PtgArea(range, this, (range != null) && (range!!.indexOf('$') == -1)) // 20090609 KSC: absolute refs if '$' -really should test if row or col
 //        p.setUseReferenceTracker(false);
        cRangeList!!.add(p)
dirtyflag = true
}

/**
 * Get the first condition of the validation as
 * a string representation
 *
 * @return
 */
     fun getFirstCond():String {
val s = FormulaParser.getExpressionString(firstCond!!)
if (s.substring(0, 1) == "=") return s.substring(1)
return s
}

/**
 * Set the first condition of the validation utilizing
 * a string.  This value must conform to the Value Type of this
 * validation or unexpected results may occur.  For example,
 * entering a string representation of a date here will not work
 * if your validation is an integer...
 *
 *
 * String passed in should be a vaild XLS formula.  Does not need to include the "="
 *
 *
 * Types of conditions
 * Integer values
 * Decimal values
 * User defined list
 * Date
 * Time
 * Text length
 * Formula
 *
 * @return
 */
     fun setFirstCond(firstCond:String?) {
this.firstCond = FormulaParser.getPtgsFromFormulaString(this, firstCond)
dirtyflag = true
}


/**
 * Get the second condition of the validation as
 * a string representation
 *
 * @return
 */
     fun getSecondCond():String {
val s = FormulaParser.getExpressionString(secondCond!!)
if (s.substring(0, 1) == "=") return s.substring(1)
return s
}

/**
 * Set the first condition of the validation utilizing
 * a string.  This value must conform to the Value Type of this
 * validation or unexpected results may occur.  For example,
 * entering a string representation of a date here will not work
 * if your validation is an integer...
 *
 *
 * String passed in should be a vaild XLS formula.  Does not need to include the "="
 *
 *
 * Types of conditions
 * Integer values
 * Decimal values
 * User defined list
 * Date
 * Time
 * Text length
 * Formula
 *
 * @return
 */
     fun setSecondCond(secondCond:String?) {
this.secondCond = FormulaParser.getPtgsFromFormulaString(this, secondCond)
dirtyflag = true
}

/**
 * set show error box if invalid values entered?
 *
 * @return
 */
     fun setShowErrMsg(showErrMsg:Boolean) {
isShowErrorMsg = showErrMsg
dirtyflag = true
}

/**
 * Get the validation type of this Dv as a byte
 *
 *
 * These bytes map to the VALUE_* static values in
 * ValidationHandle
 *
 * @return
 */
     fun getValType():Byte {
return valType
}

/**
 * Set the validation type of this Dv as a byte
 *
 *
 * These bytes map to the VALUE_* static values in
 * ValidationHandle
 *
 * @return
 */
     fun setValType(valtype:Byte) {
valType = valtype
dirtyflag = true
}

/**
 * Determines if the Dv contains the cell address passed in
 *
 * @param range
 * @return
 */
     fun isInRange(celladdy:String):Boolean {
 // FIX broken COLROW
        val rc = ExcelTools.getRowColFromString(celladdy)
for (i in cRangeList!!.indices)
{
if ((cRangeList!!.get(i) as PtgArea).contains(rc)) return true
}
return false
}

companion object {
/**
 * serialVersionUID
 */
    private val serialVersionUID = -7895028832113540094L

 // BITMASK
    private val BITMASK_VALTYPE = 0x0000000F
private val BITMASK_ERRSTYLE = 0x00000070
private val BITMASK_FSTRLOOKUP = 0x00000080
private val BITMASK_FALLOWBLANK = 0x00000100
private val BITMASK_FSUPRESSCOMBO = 0x00000200
private val BITMASK_MDIMEMODE = 0x0003FC00
private val BITMASK_FSHOWINPUTMSG = 0x00040000
private val BITMASK_FSHOWERRORMSG = 0x00080000
private val BITMASK_TYPOPERATOR = 0x00F00000


/**
 * Create a dv record & populate with prototype bytes
 *
 * @return
 */
     fun getPrototype(bk:WorkBook):XLSRecord {
val dv = Dv()
dv.opcode = XLSConstants.DV
dv.setData(dv.PROTOTYPE_BYTES)
dv.workBook = bk
dv.init()
return dv
}
/**
 * OOXML Element:
 * dataValidation (Data Validation)
 * A single item of data validation defined on a range of the worksheet
 *
 * parent: dataValidations (==Dval)
 * children: formula1, formula2
 * attributes:  many
 */

    /**
 * create a new Dv record based on OOXML input
 */
     fun parseOOXML(xpp:XmlPullParser, bs:Boundsheet):Dv {
val dv = bs.createDv(null)
dv.setSheet(bs)
try
{
var eventType = xpp.getEventType()
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
val tnm = xpp.getName()
if (tnm == "dataValidation")
{        // get attributes
for (i in 0 until xpp.getAttributeCount())
{
val n = xpp.getAttributeName(i)
val v = xpp.getAttributeValue(i)
if (n == "allowBlank")
{
dv.isAllowBlank = true
}
else if (n == "error")
{
dv.errorBoxText = v
}
else if (n == "errorStyle")
{    // default= stop
if (v == "information")
dv.errorStyle = 2.toByte()
else if (v == "stop")
dv.errorStyle = 0.toByte()
else if (v == "warning")
dv.errorStyle = 1.toByte()
}
else if (n == "errorTitle")
{
dv.errorBoxTitle = v
}
else if (n == "imeMode")
{    // TODO: what is the correct mapping??????????????????????
if (v == "nocontrol")
{
dv.imeMode = 0.toShort()
}
else if (v == "off")
{
dv.imeMode = 1.toShort()
}
else if (v == "on")
{
dv.imeMode = 2.toShort()
}
else if (v == "disabled")
{
dv.imeMode = 3.toShort()
}
else if (v == "hiragana")
{
dv.imeMode = 4.toShort()
}
else if (v == "fullKatakana")
{
dv.imeMode = 5.toShort()
}
else if (v == "halfKatakana")
{
dv.imeMode = 6.toShort()
}
else if (v == "fullAlpha")
{
dv.imeMode = 7.toShort()
}
else if (v == "halfAlpha")
{
dv.imeMode = 8.toShort()
}
else if (v == "fullHangul")
{
dv.imeMode = 9.toShort()
}
else if (v == "halfHangul")
{
dv.imeMode = 10.toShort()
}
}
else if (n == "operator")
{    // default= "between"
if (v == "between")
dv.typeOperator = 0.toByte()
else if (v == "equal")
dv.typeOperator = 2.toByte()
else if (v == "greaterThan")
dv.typeOperator = 4.toByte()
else if (v == "greaterThanOrEqual")
dv.typeOperator = 6.toByte()
else if (v == "lessThan")
dv.typeOperator = 5.toByte()
else if (v == "lessThanOrEqual")
dv.typeOperator = 7.toByte()
else if (v == "notBetween")
dv.typeOperator = 1.toByte()
else if (v == "notEqual")
dv.typeOperator = 3.toByte()
}
else if (n == "prompt")
{
dv.promptBoxText = v
}
else if (n == "promptTitle")
{
dv.promptBoxTitle = v
}
else if (n == "showDropDown")
{
 //
                            }
else if (n == "showErrorMessage")
{
dv.setShowErrMsg(true)
}
else if (n == "showInputMessage")
{
dv.showInputMsg = true
}
else if (n == "sqref")
{
dv.ooxmlranges = StringTool.splitString(v, " ")
 // 20090609 KSC: cannot add ranges in 2003 format as 2007 addresses can exceed 2003 limits
                                for (z in dv.ooxmlranges!!.indices)
dv.addOoxmlRange(dv.ooxmlranges!![z])
}
else if (n == "type")
{        // required
if (v == "custom")
 // custom formula
                                    dv.setValType(7.toByte())
else if (v == "date")
dv.setValType(4.toByte())
else if (v == "decimal")
dv.setValType(2.toByte())
else if (v == "list")
dv.setValType(3.toByte())
else if (v == "none")
dv.setValType(0.toByte())
else if (v == "textLength")
dv.setValType(6.toByte())
else if (v == "time")
dv.setValType(5.toByte())
else if (v == "whole")
dv.setValType(1.toByte())
}
}
}
else if (tnm == "formula1")
{
dv.setFirstCond(OOXMLAdapter.getNextText(xpp))
}
else if (tnm == "formula2")
{
dv.setSecondCond(OOXMLAdapter.getNextText(xpp))
}
}
else if (eventType == XmlPullParser.END_TAG)
{
val endTag = xpp.getName()
if (endTag == "dataValidation")
{
break
}
}
eventType = xpp.next()
}
}
catch (e:Exception) {
Logger.logErr("OOXMLELEMENT.parseOOXML: " + e.toString())
}

return dv
}
}
}
