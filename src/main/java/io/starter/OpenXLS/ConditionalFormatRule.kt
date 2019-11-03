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

import io.starter.formats.XLS.Cf
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.Formula
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgRef

import java.awt.*

/**
 * ConditionalFormatRule defines a single rule for manipulation of the
 * ConditionalFormat cells in Excel
 *
 *
 * Each ConditionalFormatRule contains one rule and corresponding format data.
 *
 * @see ConditionalFormatHandle
 *
 * @see Handle
 */
class ConditionalFormatRule
/**
 * Create a conditional format rule from a Cf record.
 *
 * @param theCf
 */
(theCf: Cf) : Handle {

    internal var currentCf: Cf? = null

    /**
     * Get the type operator of this ConditionalFormat as a byte.
     *
     *
     * These bytes map to the CONDITION_* static values in ConditionalFormatHandle
     *
     * @return
     */
    /**
     * set the type operator of this ConditionalFormat as a byte.
     *
     *
     * These bytes map to the CONDITION_* static values in ConditionalFormatHandle
     */
    var typeOperator: Byte
        get() = currentCf!!.operator.toByte()
        set(typOperator) = currentCf!!.setOperator(typOperator.toInt())

    /**
     * Get the second condition of the ConditionalFormat as a string representation
     *
     * @return
     */
    val secondCondition: String?
        get() = if (currentCf != null && currentCf!!.formula2 != null) currentCf!!.formula2!!.formulaString else null

    /**
     * retrieves the border colors for the current Conditional Format
     *
     * @return java.awt.Color array of Color objects for each border side (Top,
     * Left, Bottom, Right)
     * @see io.starter.formats.XLS.Cf.getBorderColors
     */
    val borderColors: Array<Color>?
        get() = currentCf!!.borderColors

    /**
     * returns the bottom border line color for the current Conditional Format
     *
     * @return int bottom border line color constant
     * @see io.starter.formats.XLS.Cf.getBorderLineColorBottom
     * @see FormatHandle.COLOR_* constants
     */
    /**
     * @param borderLineColorBottom
     * @see io.starter.formats.XLS.Cf.setBorderLineColorBottom
     */
    var borderLineColorBottom: Int
        get() = currentCf!!.borderLineColorBottom
        set(borderLineColorBottom) {
            currentCf!!.borderLineColorBottom = borderLineColorBottom
        }

    /**
     * returns the left border line color for the current Conditional Format
     *
     * @return int left border line color constant
     * @see io.starter.formats.XLS.Cf.getBorderLineColorLeft
     * @see FormatHandle.COLOR_* constants
     */
    /**
     * @param borderLineColorLeft
     * @see io.starter.formats.XLS.Cf.setBorderLineColorLeft
     */
    var borderLineColorLeft: Int
        get() = currentCf!!.borderLineColorLeft
        set(borderLineColorLeft) {
            currentCf!!.borderLineColorLeft = borderLineColorLeft
        }

    /**
     * returns the right border line color for the current Conditional Format
     *
     * @return int right border line color constant
     * @see io.starter.formats.XLS.Cf.getBorderLineColorRight
     * @see FormatHandle.COLOR_* constants
     */
    val borderLineColorRight: Int
        get() = currentCf!!.borderLineColorRight

    /**
     * returns the top border line color for the current Conditional Format
     *
     * @return int top border line color constant
     * @see io.starter.formats.XLS.Cf.getBorderLineColorTop
     * @see FormatHandle.COLOR_* constants
     */
    /**
     * @param borderLineColorTop
     * @see io.starter.formats.XLS.Cf.setBorderLineColorTop
     */
    var borderLineColorTop: Int
        get() = currentCf!!.borderLineColorTop
        set(borderLineColorTop) {
            currentCf!!.borderLineColorTop = borderLineColorTop
        }

    /**
     * returns the bottom border line style for the current Conditional Format
     *
     * @return int bottom border line style constant
     * @see io.starter.formats.XLS.Cf.getBorderLineStylesBottom
     * @see FormatHandle.BORDER* line style constants
     */
    /**
     * @param borderLineStylesBottom
     * @see io.starter.formats.XLS.Cf.setBorderLineStylesBottom
     */
    var borderLineStylesBottom: Int
        get() = currentCf!!.borderLineStylesBottom
        set(borderLineStylesBottom) {
            currentCf!!.borderLineStylesBottom = borderLineStylesBottom
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getBorderLineStylesLeft
     */
    /**
     * @param borderLineStylesLeft
     * @see io.starter.formats.XLS.Cf.setBorderLineStylesLeft
     */
    var borderLineStylesLeft: Int
        get() = currentCf!!.borderLineStylesLeft
        set(borderLineStylesLeft) {
            currentCf!!.borderLineStylesLeft = borderLineStylesLeft
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getBorderLineStylesRight
     */
    /**
     * @param borderLineStylesRight
     * @see io.starter.formats.XLS.Cf.setBorderLineStylesRight
     */
    var borderLineStylesRight: Int
        get() = currentCf!!.borderLineStylesRight
        set(borderLineStylesRight) {
            currentCf!!.borderLineStylesRight = borderLineStylesRight
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getBorderLineStylesTop
     */
    /**
     * @param borderLineStylesTop
     * @see io.starter.formats.XLS.Cf.setBorderLineStylesTop
     */
    var borderLineStylesTop: Int
        get() = currentCf!!.borderLineStylesTop
        set(borderLineStylesTop) {
            currentCf!!.borderLineStylesTop = borderLineStylesTop
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getBorderSizes
     */
    val borderSizes: IntArray?
        get() = currentCf!!.borderSizes

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getBorderStyles
     */
    val borderStyles: IntArray?
        get() = currentCf!!.borderStyles

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFont
     */
    val font: Font?
        get() = currentCf!!.font

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontColorIndex
     */
    /**
     * @param fontColorIndex
     * @see io.starter.formats.XLS.Cf.setFontColorIndex
     */
    var fontColorIndex: Int
        get() = currentCf!!.fontColorIndex
        set(fontColorIndex) {
            currentCf!!.fontColorIndex = fontColorIndex
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontEscapement
     */
    /**
     * @param fontEscapementFlag
     * @see io.starter.formats.XLS.Cf.setFontEscapement
     */
    var fontEscapement: Int
        get() = currentCf!!.fontEscapement
        set(fontEscapementFlag) {
            currentCf!!.fontEscapement = fontEscapementFlag
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontHeight
     */
    /**
     * @param fontHeight
     * @see io.starter.formats.XLS.Cf.setFontHeight
     */
    var fontHeight: Int
        get() = currentCf!!.fontHeight
        set(fontHeight) {
            currentCf!!.fontHeight = fontHeight
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontOptsCancellation
     */
    /**
     * @param fontOptsCancellation
     * @see io.starter.formats.XLS.Cf.setFontOptsCancellation
     */
    var fontOptsCancellation: Int
        get() = currentCf!!.fontOptsCancellation
        set(fontOptsCancellation) {
            currentCf!!.fontStriken = fontOptsCancellation == Cf.FONT_OPTIONS_CANCELLATION_ON
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontOptsPosture
     */
    /**
     * @param fontOptsPosture
     * @see io.starter.formats.XLS.Cf.setFontOptsPosture
     */
    var fontOptsPosture: Int
        get() = currentCf!!.fontOptsPosture
        set(fontOptsPosture) {
            currentCf!!.fontOptsPosture = fontOptsPosture
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontUnderlineStyle
     */
    /**
     * @param fontUnderlineStyle
     * @see io.starter.formats.XLS.Cf.setFontUnderlineStyle
     */
    var fontUnderlineStyle: Int
        get() = currentCf!!.fontUnderlineStyle
        set(fontUnderlineStyle) {
            currentCf!!.fontUnderlineStyle = fontUnderlineStyle
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFontWeight
     */
    /**
     * @param fontWeight
     * @see io.starter.formats.XLS.Cf.setFontWeight
     */
    var fontWeight: Int
        get() = currentCf!!.fontWeight
        set(fontWeight) {
            currentCf!!.fontWeight = fontWeight
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getForegroundColor
     */
    val foregroundColor: Int
        get() = currentCf!!.foregroundColor

    /**
     * @return
     * @see io.starter.formats.XLS.XLSRecord.getFormatPattern
     */
    val formatPattern: String?
        get() = currentCf!!.formatPattern

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFormula1
     */
    val formula1: Formula?
        get() = currentCf!!.formula1

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getFormula2
     */
    val formula2: Formula?
        get() = currentCf!!.formula2

    /**
     * returns the operator for this Conditional Format Rule <br></br>
     * e.g. "bewteen", "greater than" ...
     *
     * @return
     */
    val operator: String
        get() {
            val op = currentCf!!.operator.toInt()
            return if (op >= 0 && op < OPERATORS.size) OPERATORS[op] else "unknown operator: $op"
        }

    /**
     * returns the type of this Conditional Format <br></br>
     * e.g. "Cell value is" or "Formula value is"
     *
     * @return String Conditional Format Type
     */
    val type: String
        get() = currentCf!!.typeString

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getPatternFillColor
     */
    /**
     * @param patternFillColor
     * @see io.starter.formats.XLS.Cf.setPatternFillColor
     */
    var patternFillColor: Int
        get() = currentCf!!.patternFillColor
        set(patternFillColor) = currentCf!!.setPatternFillColor(patternFillColor, null)

    /**
     * returns the pattern color, if any, as an HTML color String. Includes custom
     * OOXML colors.
     *
     * @return String HTML Color String
     */
    val patternFgColor: String?
        get() = currentCf!!.patternFgColor

    /**
     * returns the pattern color, if any, as an HTML color String. Includes custom
     * OOXML colors.
     *
     * @return String HTML Color String
     */
    val patternBgColor: String?
        get() = currentCf!!.patternBgColor

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getPatternFillColorBack
     */
    /**
     * @param patternFillColorBack
     * @see io.starter.formats.XLS.Cf.setPatternFillColorBack
     */
    var patternFillColorBack: Int
        get() = currentCf!!.patternFillColorBack
        set(patternFillColorBack) {
            currentCf!!.patternFillColorBack = patternFillColorBack
        }

    /**
     * @return
     * @see io.starter.formats.XLS.Cf.getPatternFillStyle
     */
    /**
     * @param patternFillStyle
     * @see io.starter.formats.XLS.Cf.setPatternFillStyle
     */
    var patternFillStyle: Int
        get() = currentCf!!.patternFillStyle
        set(patternFillStyle) {
            currentCf!!.patternFillStyle = patternFillStyle
        }

    /**
     * Get the first condition of the ConditionalFormat as a string representation
     *
     * @return
     */
    val firstCondition: String
        get() = currentCf!!.formula1!!.formulaString

    val conditionalFormatType: Int
        get() = 0

    init {
        currentCf = theCf
    }

    /**
     * evaluates the criteria for this Conditional Format Rule <br></br>
     * if the criteria involves a comparison i.e. equals, less than, etc., it uses
     * the value from the passed in referenced cell to compare with
     *
     * @param Ptg refcell - the Ptg location to obtain cell value from
     * @return boolean true if evaluation of criteria passes
     * @see io.starter.formats.XLS.Cf.evaluate
     */
    fun evaluate(refcell: CellHandle): Boolean {
        val /* Ref */ pr = PtgRef.createPtgRefFromString(refcell.cellAddress, null)
        return currentCf!!.evaluate(pr)
    }

    /**
     * Set the second condition of the ConditionalFormat utilizing a string. This
     * value must conform to the Value Type of this ConditionalFormat or unexpected
     * results may occur. For example, entering a string representation of a date
     * here will not work if your ConditionalFormat is an integer...
     *
     *
     * String passed in should be a vaild XLS formula. Does not need to include the
     * "="
     *
     *
     * Types of conditions Integer values Decimal values User defined list Date Time
     * Text length Formula
     *
     *
     * Be sure that your ConditionalFormat type (getConditionalFormatType()) matches
     * the type of data.
     *
     * @return
     */
    fun setSecondCondition(secondCond: Any) {
        var setval = secondCond.toString()
        if (secondCond is java.util.Date) {
            val d = DateConverter.getXLSDateVal(secondCond)
            setval = d.toString() + ""
        }
        currentCf!!.setCondition2(setval)
    }

    /**
     * Set the first condition of the ConditionalFormat
     *
     *
     * This value must conform to the Value Type of this ConditionalFormat or
     * unexpected results may occur. For example, entering a string representation
     * of a date here will not work if your ConditionalFormat is an integer...
     *
     *
     * A java.util.Date object can also be passed in. This value will be translated
     * into an integer as excel stores dates. If you need to manipulate/retrieve
     * this value later utilize the DateConverter tool to transform the value
     *
     *
     * String passed in should be a vaild XLS formula. Does not need to include the
     * "="
     *
     *
     * Types of conditions Integer values Decimal values User defined list Date Time
     * Text length Formula
     *
     *
     * Be sure that your ConditionalFormat type (getConditionalFormatType()) matches
     * the type of data.
     *
     * @param firstCond = the first condition for the ConditionalFormat
     */
    fun setFirstCondition(firstCond: Any) {
        var setval = firstCond.toString()
        if (firstCond is java.util.Date) {
            val d = DateConverter.getXLSDateVal(firstCond)
            setval = d.toString() + ""
        }
        currentCf!!.setCondition1(setval)
    }

    companion object {
        // static shorts for setting ConditionalFormat type
        val VALUE_ANY: Byte = 0x0
        val VALUE_INTEGER: Byte = 0x1
        val VALUE_DECIMAL: Byte = 0x2
        val VALUE_USER_DEFINED_LIST: Byte = 0x3
        val VALUE_DATE: Byte = 0x4
        val VALUE_TIME: Byte = 0x5
        val VALUE_TEXT_LENGTH: Byte = 0x6
        val VALUE_FORMULA: Byte = 0x7

        // static shorts for setting action on error
        var ERROR_STOP: Byte = 0x0
        var ERROR_WARN: Byte = 0x1
        var ERROR_INFO: Byte = 0x2

        // static shorts for setting conditions on ConditionalFormat
        val CONDITION_BETWEEN: Byte = 0x0
        val CONDITION_NOT_BETWEEN: Byte = 0x1
        val CONDITION_EQUAL: Byte = 0x2
        val CONDITION_NOT_EQUAL: Byte = 0x3
        val CONDITION_GREATER_THAN: Byte = 0x4
        val CONDITION_LESS_THAN: Byte = 0x5
        val CONDITION_GREATER_OR_EQUAL: Byte = 0x6
        val CONDITION_LESS_OR_EQUAL: Byte = 0x7

        var OPERATORS = arrayOf("nocomparison", "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterOrEqual", "lessOrEqual", "beginsWith", "endsWith", "containsText", "notContains")

        /**
         * Get the byte representing the condition type string passed in. Options are'
         * "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan",
         * "greaterOrEqual", "lessOrEqual"
         *
         * @return
         */
        fun getConditionNumber(conditionType: String): Byte {
            for (i in OPERATORS.indices) {
                if (conditionType.equals(OPERATORS[i], ignoreCase = true))
                    return i.toByte()
            }
            return -1
        }

        var VALUE_TYPE = arrayOf("any", "integer", "decimal", "userDefinedList", "date", "time", "textLength", "formula")

        /**
         * Get the byte representing the value type string passed in. Options are'
         * "any", "integer", "decimal", "userDefinedList", "date", "time", "textLength",
         * "formula"
         *
         * @return
         */
        fun getValueNumber(valueType: String): Byte {
            for (i in VALUE_TYPE.indices) {
                if (valueType.equals(VALUE_TYPE[i], ignoreCase = true))
                    return i.toByte()
            }
            return -1
        }
    }
}
