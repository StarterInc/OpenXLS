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

package io.starter.formats.XLS.charts

import io.starter.formats.XLS.BiffRec
import io.starter.formats.XLS.XLSRecord

import java.util.ArrayList

open class GenericChartObject : XLSRecord(), ChartObject, ChartConstants {

    var chartType = -1        // this will be >=0 when record defines the type of chart
    var chartArr: ArrayList<XLSRecord> = ArrayList()
    override var parentChart: Chart? = null

    open var isStacked: Boolean
        get() = false
        set(bIsStacked) {}

    open var is100Percent: Boolean
        get() = false
        set(bOn) {}

    /**
     * @return String XML representation of this chart-type's options
     */
    open val optionsXML: String
        get() = ""

    override val chartRecords: ArrayList<*>
        get() = chartArr

    /**
     * Get the output array of records, including begin/end records and those of it's children.
     */
    override// 20070712 KSC: missed some recs!
    val recordArray: ArrayList<*>
        get() {
            val outputArr = ArrayList()
            outputArr.add(this)
            val nChart = -1
            for (i in chartArr.indices) {
                if (i == 0) {
                    val b = Begin.prototype as Begin?
                    outputArr.add(b)
                }

                val o = chartArr[i]
                if (o is ChartObject) {
                    val co = o as ChartObject
                    outputArr.addAll(co.recordArray)
                } else {
                    val b = o as BiffRec
                    outputArr.add(b)
                }

                if (i == chartArr.size - 1) {
                    val e = End.prototype as End?
                    outputArr.add(e)
                }
            }
            return outputArr
        }

    open fun setChartOption(op: String, `val`: String): Boolean {
        return false
    }

    open fun hasShadow(): Boolean {
        return false
    }

    /**
     * get chart option common to almost all chart types
     *
     * @param op
     * @return
     */
    open fun getChartOption(op: String): String? {
        if (op == "Stacked") { // Area, Bar, Pie, Line
            return isStacked.toString()
        } else if (op == "Shadow") { // Pie, Area, Bar, Line, Radar, Scatter
            return hasShadow().toString()
        } else if (op == "PercentageDisplay") { // Area, Bar,Line
            return is100Percent.toString()
        }
        return null
    }

    override fun addChartRecord(b: XLSRecord) {
        chartArr.add(b)
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        for (i in chartArr.indices) {
            val r = chartArr[i]
            r.close()
        }
        chartArr.clear()
        parentChart = null
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1919254120575019160L

        /**
         * adds html id and handlers for generic chart svg element
         *
         * @param title id of element
         * @return String html
         */
        fun getScript(id: String?): String {
            var ret = ""
            if (id != null)
                ret = "id='$id' "
            ret += "onmouseover='highLight(evt);' onclick='handleClick(evt);' onmouseout='restore(evt)'"
            return ret
        }
    }
}
