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

import io.starter.OpenXLS.WorkBookHandle
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * dLbls (Data Labels)
 * This element serves as a root element that specifies the settings for the data labels for an entire series or the
 * entire chart. It contains child elements that specify the specific formatting and positioning settings.
 *
 *
 * parent:  chart types, ser
 * children: dLbl (0 or more times), GROUPDLBLS (numFmt, spPr, txPr, dLblPos, showLegendKey, showVal, showCatName, showSerName,showPercent, showBubbleSize, separator, showLeaderLines
 */
// TODO: Finish All Children!!!! leaderLines, numFmt, separator
class DLbls : OOXMLElement {
    private var sp: SpPr? = null
    private var tx: TxPr? = null
    // all of these are child elements with one attribute=val, all default to "1" (true)
    private var showVal = -1
    private var showLeaderLines = -1
    private var showLegendKey = -1
    private var showCatName = -1
    private var showSerName = -1
    private var showPercent = -1
    private var showBubbleSize = -1

    override// TODO: numFmt
    // TODO: dLblPos
    // TODO: separator
    // TODO: leaderLines
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<c:dLbls>")
            if (sp != null) ooxml.append(sp!!.ooxml)
            if (tx != null) ooxml.append(tx!!.ooxml)
            if (showLegendKey != -1) ooxml.append("<c:showLegendKey val=\"$showLegendKey\"/>")
            if (showVal != -1) ooxml.append("<c:showVal val=\"$showVal\"/>")
            if (showCatName != -1) ooxml.append("<c:showCatName val=\"$showCatName\"/>")
            if (showSerName != -1) ooxml.append("<c:showSerName val=\"$showSerName\"/>")
            if (showPercent != -1) ooxml.append("<c:showPercent val=\"$showPercent\"/>")
            if (showBubbleSize != -1) ooxml.append("<c:showBubbleSize val=\"$showBubbleSize\"/>")
            if (showLeaderLines != -1) ooxml.append("<c:showLeaderLines val=\"$showLeaderLines\"/>")
            ooxml.append("</c:dLbls>")
            return ooxml.toString()
        }

    constructor(showVal: Int, showLeaderLines: Int, showLegendKey: Int, showCatName: Int,
                showSerName: Int, showPercent: Int, showBubbleSize: Int, sp: SpPr, tx: TxPr) {
        this.showVal = showVal
        this.showLeaderLines = showLeaderLines
        this.showLegendKey = showLegendKey
        this.showCatName = showCatName
        this.showSerName = showSerName
        this.showPercent = showPercent
        this.showBubbleSize = showBubbleSize
        this.sp = sp
        this.tx = tx
    }

    constructor(showVal: Boolean, showLeaderLines: Boolean, showLegendKey: Boolean, showCatName: Boolean,
                showSerName: Boolean, showPercent: Boolean, showBubbleSize: Boolean, sp: SpPr, tx: TxPr) {
        if (showVal) this.showVal = 1
        if (showLeaderLines) this.showLeaderLines = 1
        if (showLegendKey) this.showLegendKey = 1
        if (showCatName) this.showCatName = 1
        if (showSerName) this.showSerName = 1
        if (showPercent) this.showPercent = 1
        if (showBubbleSize) this.showBubbleSize = 1
        this.sp = sp
        this.tx = tx
    }

    constructor(d: DLbls) {
        this.showVal = d.showVal
        this.showLeaderLines = d.showLeaderLines
        this.showLegendKey = d.showLegendKey
        this.showCatName = d.showCatName
        this.showSerName = d.showSerName
        this.showPercent = d.showPercent
        this.showBubbleSize = d.showBubbleSize
        this.sp = d.sp
        this.tx = d.tx
    }

    /**
     * generate the ooxml necessary to define the data labels for a chart
     * Controls view of Series name, Category Name, Percents, Leader Lines, Bubble Sizes where applicable
     *
     * @return public String getOOXML(ChartFormat cf) {
     * StringBuffer ooxml= new StringBuffer();
     * ooxml.append("<c:dLbls>"); ooxml.append("\r\n");
     * // TODO: c:numFmt, c:spPr, c:txPr
     * if (cf.getChartOption("ShowBubbleSizes")=="1")
     * ooxml.append("<c:showBubbleSize val=\"1\"/>");
     * if (cf.getChartOption("ShowValueLabel")=="1")
     * ooxml.append("<c:showVal val=\"1\"/>");
     * if (cf.getChartOption("ShowLabel")=="1")
     * ooxml.append("<c:showSerName val=\"1\"/>");
     * if (cf.getChartOption("ShowCatLabel")=="1")
     * ooxml.append("<c:showCatName val=\"1\"/>");
     * // Pie specific
     * if (cf.getChartOption("ShowLabelPct")=="1")
     * ooxml.append("<c:showPercent val=\"1\"/>");
     * if (cf.getChartOption("ShowLdrLines")=="true")
     * ooxml.append("<c:showLeaderLines val=\"1\"/>");
     *
     *
     *
     *
     * ooxml.append("</c:showLeaderLines></c:showPercent></c:showCatName></c:showSerName></c:showVal></c:showBubbleSize></c:dLbls>"); ooxml.append("\r\n");
     * return ooxml.toString();
     * }
     */

    override fun cloneElement(): OOXMLElement {
        return DLbls(this)
    }

    /**
     * get methods
     */
    fun showLegendKey(): Boolean {
        return showLegendKey == 1
    }

    fun showVal(): Boolean {
        return showVal == 1
    }

    fun showCatName(): Boolean {
        return showCatName == 1
    }

    fun showSerName(): Boolean {
        return showSerName == 1
    }

    fun showPercent(): Boolean {
        return showPercent == 1
    }

    fun showBubbleSize(): Boolean {
        return showBubbleSize == 1
    }

    fun showLeaderLines(): Boolean {
        return showLeaderLines == 1
    }

    companion object {

        private val serialVersionUID = -3765320144606034211L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var sp: SpPr? = null
            var tx: TxPr? = null
            var showVal = -1
            var showLeaderLines = -1
            var showLegendKey = -1
            var showCatName = -1
            var showSerName = -1
            var showPercent = -1
            var showBubbleSize = -1
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "spPr") {
                            lastTag.push(tnm)
                            sp = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //	        			 sp.setNS("c");
                        } else if (tnm == "txPr") {
                            lastTag.push(tnm)
                            tx = TxPr.parseOOXML(xpp, lastTag, bk).cloneElement() as TxPr
                        } else if (tnm == "showVal") {
                            if (xpp.attributeCount > 0)
                                showVal = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "showLeaderLines") {
                            if (xpp.attributeCount > 0)
                                showLeaderLines = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "showLegendKey") {
                            if (xpp.attributeCount > 0)
                                showLegendKey = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "showCatName") {
                            if (xpp.attributeCount > 0)
                                showCatName = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "showSerName") {
                            if (xpp.attributeCount > 0)
                                showSerName = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "showPercent") {
                            if (xpp.attributeCount > 0)
                                showPercent = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "showBubbleSize") {
                            if (xpp.attributeCount > 0)
                                showBubbleSize = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "dLbls") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("dLbls.parseOOXML: $e")
            }

            return DLbls(showVal, showLeaderLines, showLegendKey, showCatName, showSerName,
                    showPercent, showBubbleSize, sp, tx)
        }
    }
}



