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
import io.starter.formats.XLS.Condfmt
import io.starter.toolkit.Logger

import java.util.ArrayList

/**
 * <pre>
 * ConditionalFormatHandle allows for manipulation of the ConditionalFormat cells in Excel
 *
 * Using the ConditionalFormatHandle, the affected range of ConditionalFormats can be modified,
 * along with the formatting applied to the cells when the condition is true.
 *
 * Each ConditionalFormatHandle represents a range of cells and can have a number
 * of formatting rules and formats (ConditionalFormatRule) applied.
 *
 * The ConditionalFormatHandle affected range can either be a contiguous range, or a series of cells and ranges.
 *
 * Each ConditionalFormatRule contains one rule and corresponding format data.
 *
 *
 * Many of these calls are very self-explanatory and can be found in the api.
</pre> *
 */
class ConditionalFormatHandle
/**
 * For internal use only.  Creates a ConditionalFormat Handle based of the Condfmt passed in.
 *
 * @param workBookHandle
 * @param Condfmt
 */
(c: Condfmt,
 /**
  * get the WorkSheetHandle for this ConditionalFormat
  *
  *
  * ConditionalFormats are bound to a specific worksheet and cannot be
  * applied to multiple worksheets
  *
  * @return the WorkSheetHandle for this ConditionalFormat
  */
 val workSheetHandle: WorkSheetHandle) : Handle {


    /**
     * @return Returns the cndfmt.
     */
    /**
     * @param cndfmt The cndfmt to set.
     */
    var cndfmt: Condfmt? = null
        protected set


    /**
     * Get all the rules assocated with this conditional format record
     *
     * @return
     */
    val rules: Array<ConditionalFormatRule>
        get() {
            val cfs = this.cndfmt!!.rules
            val rules = arrayOfNulls<ConditionalFormatRule>(cfs.size)
            for (i in cfs.indices) {
                val cfr = ConditionalFormatRule(cfs[i] as Cf)
                rules[i] = cfr
            }
            return rules
        }


    /**
     * Return the range of data this ConditionalFormatHandle refers to as a string
     * This location is the largest bounding rectangle that all cells utilized in this conditional
     * format can be contained in.
     *
     * @return Encompassing range in the format "A2:B12"
     */
    val encompassingRange: String
        get() {
            val rowcols = cndfmt!!.encompassingRange
            return ExcelTools.formatRangeRowCol(rowcols)
        }


    /**
     * Return a string representing all ranges that this conditional format handle can affect
     *
     * @return range in the format "A2:B3";
     */
    val allAffectedRanges: Array<String>
        get() = cndfmt!!.allRanges


    /**
     * Return an xml representation of the ConditionalFormatHandle
     *
     * @return
     */
    // TODO: more than one cf rule???
    val xml: String
        get() {
            val rules = this.rules
            val xml = StringBuffer()
            xml.append("<dataConditionalFormat")
            xml.append(" type=\"" + ConditionalFormatRule.VALUE_TYPE[rules[0].conditionalFormatType] + "\"")
            xml.append(" operator=\"" + ConditionalFormatRule.OPERATORS[rules[0].typeOperator] + "\"")
            try {
                xml.append(" sqref=\"" + this.encompassingRange + "\"")
            } catch (e: Exception) {
                Logger.logErr("Problem getting range for ConditionalFormatHandle.getXML().", e)
            }

            xml.append(">")
            if (rules[0].firstCondition != null) {
                xml.append("<formula1>")
                xml.append(rules[0].firstCondition)
                xml.append("</formula1>")
            }
            if (rules[0].secondCondition != null) {
                xml.append("<formula2>")
                xml.append(rules[0].secondCondition)
                xml.append("</formula2>")
            }
            xml.append("</dataConditionalFormat>")
            return xml.toString()
        }

    /**
     * returns the formatting for each rule of this Contditional Format Handle
     *
     * @return FormatHandle[]
     */
    //cfm.initCells(this);	// added!
    val formats: Array<FormatHandle>
        get() {
            val fmx = arrayOfNulls<FormatHandle>(this.cndfmt!!.rules.size)
            if (cndfmt!!.formatHandle == null) {
                val cfxe = cndfmt!!.cfxe
                val fz = FormatHandle(cndfmt!!, workSheetHandle.workBook!!, cfxe, null)
            }
            for (t in fmx.indices) {
                cndfmt!!.formatHandle!!.updateFromCF(this.cndfmt!!.rules[t] as Cf, workSheetHandle.workBook)
                fmx[t] = cndfmt!!.formatHandle
            }
            return fmx
        }


    /**
     * evaluates the criteria for this Conditional Format
     * <br></br>if the criteria involves a comparison i.e. equals, less than, etc., it uses the
     * value from the passed in referenced cell to compare with
     *
     *
     * If there are multiple rules in the ConditionalFormat then the first rule that passes will be returned.
     *
     *
     * If no valid rules pass then a null result is given
     *
     * @param CellHandle refcell - the cell to obtain a value from in order for evaluation to occur
     * @return the ConditionalFormatRule that passes.
     * @see io.starter.formats.XLS.Cf.evaluate
     */
    fun evaluate(refcell: CellHandle): ConditionalFormatRule? {
        val rules = this.rules
        for (i in rules.indices) {
            if (rules[i].evaluate(refcell)) return rules[i]
        }
        return null
    }


    init {
        this.cndfmt = c
    }


    /**
     * Determine if the conditional format contains/affects the cell handle passed in
     *
     * @param cellHandle
     * @return
     */
    operator fun contains(cellHandle: CellHandle): Boolean {
        return this.cndfmt!!.contains(cellHandle.intLocation)
    }


    /**
     * Set the range this ConditionalFormatHandle refers to.
     * Pass in a range string, sans worksheet.
     *
     *
     * This range will overwrite all other ranges this ConditionalFormatHandle refers to.
     *
     *
     * In order to handle multiple ranges, use the addRange(String range) method
     *
     * @param range = standard excel range without worksheet information ("A1" or "A1:A10")
     */
    fun setRange(range: String) {
        this.cndfmt!!.resetRange(range)
    }

    /**
     * Determines if the ConditionalFormatHandle contains the cell address passed in
     *
     * @param cellAddress a cell address in the format "A1"
     * @return if the ce
     */
    operator fun contains(celladdy: String): Boolean {
        return this.cndfmt!!.contains(ExcelTools.getRowColFromString(celladdy))
    }

    /**
     * return a string representation of this Conditional Format
     *
     *
     * This method is still incomplete as it only returns data for one rule, and only refers to one range
     */
    // TODO: more than one cf rule???
    override fun toString(): String {
        val rules = this.rules
        var ret = this.encompassingRange + ": " + rules[0].type + " " + rules[0].operator    // range, type + operator
        if (rules[0].formula1 != null)
        // formulas
            ret += " " + rules[0].formula1!!.formulaString.substring(1)
        if (rules[0].formula2 != null)
            ret += " and " + rules[0].formula2!!.formulaString.substring(1)
        // todo: add formats to this
        return ret
    }


    /**
     * Add a cell to this conditional format record
     *
     * @param cellHandle
     */
    fun addCell(cellHandle: CellHandle) {
        if (this.contains(cellHandle)) return
        cndfmt!!.addLocation(cellHandle.cellAddress)
    }

}
