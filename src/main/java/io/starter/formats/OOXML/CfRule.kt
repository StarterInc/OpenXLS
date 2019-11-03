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
package io.starter.formats.OOXML

import io.starter.formats.XLS.OOXMLAdapter
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList
import java.util.HashMap

/**
 * cfRule (Conditional Formatting Rule)
 * This collection represents a description of a conditional formatting rule.
 * <br></br>
 * NOTE: now is used merely to parse the cfRule OOXML element
 * the data is stored in a BIFF-8 Cf object
 *
 *
 * parent:       conditionalFormatting
 * children:     SEQ: formula (0-3), colorScale, dataBar, iconSet
 * attributes:   type, dxfId, priority (REQ), stopIfTrue, aboveAverage,
 * percent, bottom, operator, text, timePeriod, rank, stdDev, equalAverage
 */
//TODO: Finish children colorScale, dataBar, iconSet
class CfRule : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var formulas: ArrayList<String>? = null

    /**
     * get the dxfId (the incremental style associated with this conditional formatting rule)
     *
     * @return
     */
    /**
     * set the dxfId (the incremental style associated with this conditional formatting rule)
     *
     * @param dxfId
     */
    // it's possible to not specify a dxfId
    var dxfId: Int
        get() {
            if (this.attrs != null)
                try {
                    return Integer.valueOf(this.attrs!!["dxfId"]).toInt()
                } catch (e: Exception) {
                }

            return -1
        }
        @Deprecated("")
        set(dxfId) {
            if (this.attrs == null) this.attrs = HashMap()
            attrs!!["dxfId"] = Integer.valueOf(dxfId)!!.toString()
        }

    override// TODO Auto-generated method stub
    val ooxml: String?
        get() = null

    /**
     * get methods
     */
    // valid only when type=="cellIs"
    val operator: String?
        get() = if (attrs != null) attrs!!["operator"] else null

    val type: String
        get() {
            var type: String? = null
            if (attrs != null)
                type = attrs!!["type"]
            if (type == null)
                type = "cellIs"
            return type
        }

    /**
     * returns the text to test in a containsText type of condition
     * <br></br>Only valid for containsText conditions
     *
     * @return
     */
    val containsText: String?
        get() = if (attrs != null) attrs!!["text"] else null

    val formula1: String?
        get() = if (formulas != null && formulas!!.size > 0) formulas!![0] else null

    val formula2: String?
        get() = if (formulas != null && formulas!!.size > 1) formulas!![1] else null

    constructor(attrs: HashMap<String, String>, formulas: ArrayList<String>) {
        this.attrs = attrs
        this.formulas = formulas
    }

    constructor(cf: CfRule) {
        this.attrs = cf.attrs
        this.formulas = cf.formulas
    }

    /**
     * generate OOXML for this cfRule
     * NOW Cf is parsed to obtain OOXML and CfRule is not retained
     *
     * @return
     */
    @Deprecated("")
    fun getOOXML(priority: Int): String {
        val ooxml = StringBuffer()
        ooxml.append("<cfRule")
        // attributes
        // TODO: use passed in priority
        val i = attrs!!.keys.iterator()
        while (i.hasNext()) {
            val key = i.next()
            val `val` = attrs!![key]
            ooxml.append(" $key=\"$`val`\"")
        }
        ooxml.append(">")
        if (formulas != null) {
            for (j in formulas!!.indices) {
                ooxml.append("<formula>" + OOXMLAdapter.stripNonAsciiRetainQuote(formulas!![j]) + "</formula>")
            }
        }
        // TODO: finish children dataBar, colorScale, iconSet
        ooxml.append("</cfRule>")
        return ooxml.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return CfRule(this)
    }

    companion object {

        private val serialVersionUID = 8509907308100079138L

        /**
         * generate a cfRule based on OOXML input stream
         *
         * @param xpp
         * @return
         */
        // TODO: finish children colorScale, dataBar, iconSet
        fun parseOOXML(xpp: XmlPullParser): CfRule {
            val attrs = HashMap<String, String>()
            val formulas = ArrayList<String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cfRule") {     // get attributes: priority, type, operator, dxfId
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "formula") {
                            formulas.add(OOXMLAdapter.getNextText(xpp))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cfRule") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cfRule.parseOOXML: $e")
            }

            return CfRule(attrs, formulas)
        }
    }
}

/**
 * dataBar (Data Bar)
 * Describes a data bar conditional formatting rule.
 * [Example:
 * In this example a data bar conditional format is expressed, which spreads across all cell values in the cell range,
 * and whose color is blue.
 * <dataBar>
 * <cfvo type="min" val="0"></cfvo>
 * <cfvo type="max" val="0"></cfvo>
 * <color rgb="FF638EC6"></color>
</dataBar> *
 * end example]
 * The length of the data bar for any cell can be calculated as follows:
 * Data bar length = minLength + (cell value - minimum value in the range) / 1 (maximum value in the range -
 * minimum value in the range) * (maxLength - minLength),
 * where min and max length are a fixed percentage of the column width (by default, 10% and 90% respectively.)
 * The minimum difference in length (or increment amount) is 1 pixel.
 *
 *
 * parent:       cfRule
 * children:     SEQ: cfvo (2), color (1)
 * attributes:   minLength (def=10), maxLength (def=90), showValue (def=true)
 *
 *
 * colorScale (Color Scale)
 * Describes a graduated color scale in this conditional formatting rule.
 *
 *
 * parent:      cfRule
 * children:    SEQ: cfvo (2+), color (2+)
 *
 *
 * cfvo (Conditional Format Value Object)
 * Describes the values of the interpolation points in a gradient scale.
 *
 *
 * parents: colorScale, dataBar, iconSet
 * children: extLst
 * attributes:  type (REQ), val, gte (def=true)
 */


/**
 * colorScale (Color Scale)
 * Describes a graduated color scale in this conditional formatting rule.
 *
 * parent:      cfRule
 * children:    SEQ: cfvo (2+), color (2+)
 */

/**
 * cfvo (Conditional Format Value Object)
 * Describes the values of the interpolation points in a gradient scale.
 *
 * parents: colorScale, dataBar, iconSet
 * children: extLst
 * attributes:  type (REQ), val, gte (def=true)
 *
 */

