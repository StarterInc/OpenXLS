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
/*
 * XSLConverterTool is a collection of methods that are of use in the xml/xsl conversion for sheetster.
 *
 * These methods are called from XSL to populate various fields correctly
 *
 * Created on Mar 17, 2006
 *
 */
package io.starter.toolkit

import io.starter.OpenXLS.ExcelTools

import java.text.NumberFormat
import java.util.Hashtable

class XSLConverterTool {

    private val styles = Hashtable<String, Style>()
    private var lastCell = "A1"

    /**
     * Gets a date format pattern based off of an ifmt. This applies only to built
     * in dates.
     *
     * @param ifmt from XF record
     * @return date format pattern
     */
    fun getDateFormatPattern(ifmt: String): String {
        if (ifmt == "14") {
            return "m/d/yy"
        } else if (ifmt == "15") {
            return "d-mmm-yy"
        } else if (ifmt == "16") {
            return "d-mmm"
        } else if (ifmt == "17") {
            return "mmm-yy"
        } else if (ifmt == "22") {
            return "m/d/yy h:mm"
        }
        return "m/d/yy"

    }

    /**
     * Gets a calendar format pattern based off of an ifmt. This applies only to
     * built in dates.
     *
     * @param ifmt from XF record
     * @return calendar format pattern
     */
    fun getJsCalendarFormatPattern(ifmt: String): String {
        return if (ifmt == "14") {
            "%m/%d/%Y"
        } else if (ifmt == "15") {
            "%d-%b-%y"
        } else if (ifmt == "16") {
            "%d-%b"
        } else if (ifmt == "17") {
            "%m-%y"
        } else if (ifmt == "22") {
            "%m/%d/%Y %h:%M"
        } else
            "%m/%d/%Y"
    }

    /**
     * Returns a formatted currency string based off the local format and the string
     * format passed in. TODO: implement formatting patterns
     *
     * @param fmt
     * @return
     */
    fun getCurrencyFormat(fmt: String, value: String): String? {
        try {
            val nf = NumberFormat.getCurrencyInstance()
            val d = Double(value)
            return nf.format(d)
        } catch (e: NumberFormatException) {
            return value
        }

    }

    /**
     * Get a style based of a style ID. If the style does not yet exist, create a
     * new one, and add it to the hashtable of styles
     *
     * @param styleId
     * @return Style
     */
    private fun getStyle(styleId: String): Style {
        val o = styles[styleId]
        if (o != null) {
            return o
        }
        val thisStyle = Style(styleId)
        styles[styleId] = thisStyle
        return thisStyle
    }

    /**
     * returns a String populated with cell data for missing cells since the last
     * cell read. Requires the first and last cells to exist for a row.
     *
     * @return html fragment
     */
    fun getPreviousCellData(sheet: String, currCell: String, colspan: String?): String {
        var currCell = currCell
        val returnString = StringBuffer()
        var colSpan = 1
        if (colspan != null && colspan !== "") {
            colSpan = Integer.parseInt(colspan)
        }
        val newCell = ExcelTools.getRowColFromString(currCell)
        val oldCell = ExcelTools.getRowColFromString(lastCell)
        val newCol = newCell[1]
        var oldCol = oldCell[1] + 1
        while (oldCol < newCol) {
            val newAddress = ExcelTools.getAlphaVal(oldCol) + (oldCell[0] + 1)
            returnString.append(getEmptyCellHTML(sheet, newAddress))
            oldCol++
        }
        newCell[1] = newCol + colSpan - 1
        currCell = ExcelTools.formatLocation(newCell)
        lastCell = currCell
        return returnString.toString()
    }

    private fun getEmptyCellHTML(sheet: String, address: String): String {
        return "<td id=\"$address\" > </td>"
    }

    /*******************************************************************************************************
     * DELEGATING METHODS ***********************************************
     */
    fun setStyleColor(styleId: String, color: String) {
        val thisStyle = getStyle(styleId)
        thisStyle.color = color
    }

    fun getStyleColor(styleId: String): String {
        val thisStyle = getStyle(styleId)
        return thisStyle.color
    }

    fun setIsDate(styleId: String, isDate: String) {
        val thisStyle = getStyle(styleId)
        thisStyle.isDate = isDate
    }

    fun getIsDate(styleId: String): String {
        val thisStyle = getStyle(styleId)
        return thisStyle.isDate ?: return "0"
    }

    fun setIsCurrency(styleId: String, isCurrency: String) {
        var isCurrency = isCurrency
        if (isCurrency == "")
            isCurrency = "0"
        val thisStyle = getStyle(styleId)
        thisStyle.isCurrency = isCurrency
    }

    fun getIsCurrency(styleId: String): String {
        val thisStyle = getStyle(styleId)
        return thisStyle.isCurrency ?: return "0"
    }

    fun setFormatId(styleId: String, formatId: String) {
        val thisStyle = getStyle(styleId)
        thisStyle.formatId = formatId
    }

    fun getFormatId(styleId: String): String {
        val thisStyle = getStyle(styleId)
        return thisStyle.formatId ?: return "0"
    }

    fun getFormatPattern(styleId: String): String? {
        val thisStyle = getStyle(styleId)
        return thisStyle.formatPattern
    }

    fun setFormatPattern(styleId: String, pattern: String) {
        val thisStyle = getStyle(styleId)
        thisStyle.formatPattern = pattern
    }

    /**
     * Style holds style information about a certain style in the xsl spreadsheet
     */
    private inner class Style(var id: String?) {
        var formatId: String? = null
        var color = ""
        private val test: String? = null
        var fontFamily: String? = null
        var fontSize: String? = null
        var fontWeight: String? = null
        var fontColor: String? = null
        var textAlign: String? = null
        var isDate: String? = null
        var isCurrency: String? = null
        var formatPattern: String? = null

    }
}
