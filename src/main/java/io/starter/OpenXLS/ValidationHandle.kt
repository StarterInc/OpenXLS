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

import io.starter.formats.XLS.Dv
import io.starter.formats.XLS.ValidationException
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool

/**
 * ValidationHandle allows for manipulation of the validation cells in Excel
 *
 *
 * Using the ValidationHandle, the affected range of validations can be
 * modified, along with many of the error messages and actual validation upon
 * the cells. Many of these calls are very self-explanatory and can be found in
 * the api.
 *
 * Some common use cases:
 *
 *  1. Setting/changing the validation data type.
 * <pre>
 * // Change validation to only allow a formula ValidationHandle validator =
 * theSheet.getValidationHandle("A1");
 * validator.setValidationType(ValidationHandle.VALUE_FORMULA); // can use any VALUE static byte here.
</pre> *
 *
 *  1.
 * Setting/changing the validation condition. This requires setting the
 * condition type, and a first and second condition. Also make the page error on
 * invalid entry and set the error text.
 * <pre>
 * // Validate cell is an int between the current values of cell D1 and cell D2.
 * ValidationHandle validator = theSheet.getValidationHandle("A1");
 * validator.setValidationType(ValidationHandle.VALUE_INTEGER);
 * validator.setTypeOperator(ValidationHandle.CONDITION_BETWEEN)// any CONDITION
 * static byte validator.setFirstCondition("D1"); // any valid excel formula,
 * omitting the '=' validator.setSecondCondition("D2");
 * validator.setErrorBoxText
 * ("The value is not between the values of D1 and D2");
 * validator.setShowErrorMessage(true);
</pre> *
 *  1.
 * Change the range the validation is applied to, for instance if one is
 * inseting a number of new rows and wants to grow the range A1:D1 to A1:Z1.
 * <pre>
 * ValidationHandle validator = theSheet.getValidationHandle("A1");
 * validator.setRange("A1:Z1");
</pre> *
 *
 *
 * "http://starter.io">Starter Inc.
 */

class ValidationHandle
/**
 * Create a new Validation for the input cell range.
 *
 * Validations are specific to worksheets.
 *
 *
 * @param cellRange
 * = cell or range of cells, example "A1", "A1:A10"
 *
 * protected ValidationHandle(String cellRange, WorkSheetHandle
 * wsh) {
 *
 * }
 */
/**
 * For internal use only. Creates a Validation Handle based of the Dv passed
 * in.
 *
 * @param dv
 */
(private val myDv: Dv) : Handle {

    /**
     * Return the range of data this ValidationHandle refers to as a string Will
     * not contain worksheet identifier, as ValidationHandles are specific to a
     * worksheet. If the Validation effects multiple ranges they are separated
     * by a space
     *
     * @return ptgRef.toString()
     */
    /**
     * Set the range this ValidationHandle refers to. Pass in a range string,
     * sans worksheet.
     *
     *
     * This range will overwrite all other ranges this ValidationHandle refers
     * to.
     *
     *
     * In order to handle multiple ranges, use the addRange(String range) method
     *
     * @param range = standard excel range without worksheet information ("A1" or
     * "A1:A10")
     */
    var range: String
        get() {
            val s = myDv.ranges
            var out = ""
            for (i in s.indices) {
                if (i > 0)
                    out += " "
                out += s[i]
            }
            return out
        }
        set(range) = myDv.setRange(range)

    /**
     * Get the text from the error box.
     *
     * @return
     */
    /**
     * Set the text for the error box
     *
     * @param textError
     */
    var errorBoxText: String
        get() = myDv.errorBoxText
        set(textError) {
            myDv.errorBoxText = textError
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
        get() = myDv.promptBoxText
        set(text) {
            myDv.promptBoxText = text
        }

    /**
     * Get the title from the error box
     *
     * @return
     */
    /**
     * Set the title for the error box
     *
     * @param textError
     */
    var errorBoxTitle: String
        get() = myDv.errorBoxTitle
        set(textError) {
            myDv.errorBoxTitle = textError
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
        get() = myDv.promptBoxTitle
        set(text) {
            myDv.promptBoxTitle = text
        }

    /**
     * Return a byte representing the error style for this ValidationHandle
     *
     *
     * These map to the static final ints ERROR_* in ValidationHandle
     *
     * @return
     */
    /**
     * Set the error style for this ValidationHandle record
     *
     *
     * These map to the static final ints ERROR_* from ValidationHandle
     *
     * @return
     */
    var errorStyle: Byte
        get() = myDv.errorStyle
        set(errstyle) {
            myDv.errorStyle = errstyle
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
        get() = myDv.imeMode
        set(mode) {
            myDv.imeMode = mode
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
        get() = myDv.isAllowBlank
        set(allowBlank) {
            myDv.isAllowBlank = allowBlank
        }

    /**
     * Get the first condition of the validation as a string representation
     *
     * @return
     */
    val firstCondition: String?
        get() = myDv.firstCond

    /**
     * Get the second condition of the validation as a string representation
     *
     * @return
     */
    val secondCondition: String?
        get() = myDv.secondCond

    /**
     * Show error box if invalid values entered?
     *
     * @return
     */
    /**
     * Set show error box if invalid values entered?
     *
     * @return
     */
    var isShowErrorMsg: Boolean
        get() = myDv.isShowErrorMsg
        set(showErrMsg) = myDv.setShowErrMsg(showErrMsg)

    /**
     * Show prompt box if cell selected?
     *
     * @return
     */
    /**
     * Set show prompt box if cell selected?
     *
     * @param showInputMsg
     */
    var showInputMsg: Boolean
        get() = myDv.showInputMsg
        set(showInputMsg) {
            myDv.showInputMsg = showInputMsg
        }

    /**
     * In list type validity the string list is explicitly given in the formula
     *
     * @return boolean
     */
    /**
     * In list type validity the string list is explicitly given in the formula
     *
     * @param strLookup
     */
    var isStrLookup: Boolean
        get() = myDv.isStrLookup
        set(strLookup) {
            myDv.isStrLookup = strLookup
        }

    /**
     * Suppress the drop down arrow in list type validity
     *
     * @return boolean
     */
    /**
     * Suppress the drop down arrow in list type validity
     */
    var isSuppressCombo: Boolean
        get() = myDv.isSuppressCombo
        set(suppressCombo) {
            myDv.isSuppressCombo = suppressCombo
        }

    /**
     * Get the type operator of this validation as a byte.
     *
     *
     * These bytes map to the CONDITION_* static values in ValidationHandle
     *
     * @return
     */
    /**
     * set the type operator of this validation as a byte.
     *
     *
     * These bytes map to the CONDITION_* static values in ValidationHandle
     */
    var typeOperator: Byte
        get() = myDv.typeOperator
        set(typOperator) {
            myDv.typeOperator = typOperator
        }

    /**
     * Get the validation type of this ValidationHandle as a byte
     *
     *
     * These bytes map to the VALUE_* static values in ValidationHandle
     *
     * @return
     */
    /**
     * Set the validation type of this ValidationHandle as a byte
     *
     *
     * These bytes map to the VALUE_* static values in ValidationHandle
     */
    var validationType: Byte
        get() = myDv.valType
        set(valtype) {
            myDv.valType = valtype
        }

    /**
     * Return an xml representation of the ValidationHandle
     *
     * @return
     */
    val xml: String
        get() {
            val xml = StringBuffer()
            xml.append("<datavalidation")
            xml.append(" type=\"" + VALUE_TYPE[this.validationType] + "\"")
            xml.append(" operator=\"" + CONDITIONS[this.typeOperator] + "\"")
            xml.append(" allowBlank=\"" + (if (this.isAllowBlank) "1" else "0")
                    + "\"")
            xml.append(" showInputMessage=\""
                    + (if (this.showInputMsg) "1" else "0") + "\"")
            xml.append(" showErrorMessage=\""
                    + (if (this.isShowErrorMsg) "1" else "0") + "\"")
            xml.append(" errorTitle=\"" + this.errorBoxTitle + "\"")
            xml.append(" error=\"" + this.errorBoxText + "\"")
            xml.append(" promptTitle=\"" + this.promptBoxTitle + "\"")
            xml.append(" prompt=\"" + this.promptBoxText + "\"")
            try {
                xml.append(" sqref=\"" + this.range + "\"")
            } catch (e: Exception) {
                Logger.logErr(
                        "Problem getting range for ValidationHandle.getXML().", e)
            }

            xml.append(">")
            if (this.firstCondition != null) {
                xml.append("<formula1>")
                xml.append(StringTool.convertXMLChars(this.firstCondition))
                xml.append("</formula1>")
            }
            if (this.secondCondition != null) {
                xml.append("<formula2>")
                xml.append(StringTool.convertXMLChars(this.secondCondition))
                xml.append("</formula2>")
            }
            xml.append("</datavalidation>")
            return xml.toString()
        }

    /**
     * Determine if the value passed in is valid for this validation
     *
     * @param value
     * @return
     */
    @Throws(ValidationException::class)
    fun isValid(value: Any): Boolean {
        return myDv.isValid(value)
    }

    /**
     * Determine if the value passed in is valid for this validation
     *
     * @param value
     * @return
     */
    @Throws(RuntimeException::class)
    fun isValid(value: Any, throwException: Boolean): Boolean {
        if (throwException) {
            try {
                return myDv.isValid(value)
            } catch (e: Exception) {
                try {
                    throw e
                } catch (ex: Exception) {
                    Logger.logErr("Error getting isValid $ex")
                }

            }

        }
        try {
            return myDv.isValid(value)
        } catch (e: Exception) {
            return false
        }

    }

    /**
     * Adds an additional range to the existing ranges in this validationhandle
     *
     * @param range
     */
    fun addRange(range: String) {
        myDv.addRange(range)
    }

    /**
     * Set the first condition of the validation
     *
     *
     * This value must conform to the Value Type of this validation or
     * unexpected results may occur. For example, entering a string
     * representation of a date here will not work if your validation is an
     * integer...
     *
     *
     * A java.util.Date object can also be passed in. This value will be
     * translated into an integer as excel stores dates. If you need to
     * manipulate/retrieve this value later utilize the DateConverter tool to
     * transform the value
     *
     *
     * String passed in should be a vaild XLS formula. Does not need to include
     * the "="
     *
     *
     * Types of conditions Integer values Decimal values User defined list Date
     * Time Text length Formula
     *
     *
     * Be sure that your validation type (getValidationType()) matches the type
     * of data.
     *
     * @param firstCond = the first condition for the validation
     */
    fun setFirstCondition(firstCond: Any) {
        var setval = firstCond.toString()
        if (firstCond is java.util.Date) {
            val d = DateConverter.getXLSDateVal(firstCond)
            setval = d.toString() + ""
        }
        myDv.firstCond = setval
    }

    /**
     * Set the first condition of the validation utilizing a string. This value
     * must conform to the Value Type of this validation or unexpected results
     * may occur. For example, entering a string representation of a date here
     * will not work if your validation is an integer...
     *
     *
     * String passed in should be a vaild XLS formula. Does not need to include
     * the "="
     *
     *
     * Types of conditions Integer values Decimal values User defined list Date
     * Time Text length Formula
     *
     *
     * Be sure that your validation type (getValidationType()) matches the type
     * of data.
     *
     * @return
     */
    fun setSecondCondition(secondCond: Any) {
        var setval = secondCond.toString()
        if (secondCond is java.util.Date) {
            val d = DateConverter.getXLSDateVal(secondCond)
            setval = d.toString() + ""
        }
        myDv.secondCond = setval
    }

    /**
     * Determines if the ValidationHandle contains the cell address passed in
     *
     * @param range
     * @return
     */
    fun isInRange(celladdy: String): Boolean {
        return myDv.isInRange(celladdy)
    }

    companion object {

        // static shorts for setting validation type
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

        // static shorts for setting conditions on validation
        val CONDITION_BETWEEN: Byte = 0x0
        val CONDITION_NOT_BETWEEN: Byte = 0x1
        val CONDITION_EQUAL: Byte = 0x2
        val CONDITION_NOT_EQUAL: Byte = 0x3
        val CONDITION_GREATER_THAN: Byte = 0x4
        val CONDITION_LESS_THAN: Byte = 0x5
        val CONDITION_GREATER_OR_EQUAL: Byte = 0x6
        val CONDITION_LESS_OR_EQUAL: Byte = 0x7

        // static shorts for setting IME modes
        var IME_MODE_NO_CONTROL: Short = 0x0 // No control for IME.
        // (default)
        var IME_MODE_ON: Short = 0x1 // IME is on.
        var IME_MODE_OFF: Short = 0x2 // IME is off.
        var IME_MODE_DISABLE: Short = 0x3// IME is disabled.
        var IME_MODE_HIRAGANA: Short = 0x4// IME is in hiragana input
        // mode.
        var IME_MODE_KATAKANA: Short = 0x5// IME is in full-width katakana
        // input mode.
        var IME_MODE_KATALANA_HALF: Short = 0x6// IME is in half-width
        // katakana input mode.
        var IME_MODE_FULL_WIDTH_ALPHA: Short = 0x7// IME is in full-width
        // alphanumeric input
        // mode
        var IME_MODE_HALF_WIDTH_ALPHA: Short = 0x8// IME is in half-width
        // alphanumeric input
        // mode.
        var IME_MODE_FULL_WIDTH_HANKUL: Short = 0x9// IME is in full-width
        // Hankul input mode
        var IME_MODE_HALF_WIDTH_HANKUL: Short = 0x10// IME is in
        // half-width Hankul
        // input mode.

        var CONDITIONS = arrayOf("between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterOrEqual", "lessOrEqual")

        /**
         * Get the byte representing the condition type string passed in. Options
         * are' "between", "notBetween", "equal", "notEqual", "greaterThan",
         * "lessThan", "greaterOrEqual", "lessOrEqual"
         *
         * @return
         */
        fun getConditionNumber(conditionType: String): Byte {
            for (i in CONDITIONS.indices) {
                if (conditionType.equals(CONDITIONS[i], ignoreCase = true))
                    return i.toByte()
            }
            return -1
        }

        var VALUE_TYPE = arrayOf("any", "integer", "decimal", "userDefinedList", "date", "time", "textLength", "formula")

        /**
         * Get the byte representing the value type string passed in. Options are'
         * "any", "integer", "decimal", "userDefinedList", "date", "time",
         * "textLength", "formula"
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
