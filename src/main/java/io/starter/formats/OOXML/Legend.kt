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
 * OOXML Legend element
 *
 *
 * parent:  chart
 * children:	layout, legendEntry, legendPos, overlay, spPr, txPr
 */
class Legend : OOXMLElement {
    private var legendpos: String? = null    // actually an element but only contains 1 attribute: val
    private var le: LegendEntry? = null
    private var layout: Layout? = null
    private var shapeProps: SpPr? = null                    // defines the shape properties for the legend for this chart
    private var txpr: TxPr? = null                        // defines text properties
    private var overlay: String? = null

    /**
     * generate the ooxml necessary to display chart legend
     *
     * @return
     */
    override// sequence
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<c:legend>")
            if (legendpos != null) ooxml.append("<c:legendPos val=\"$legendpos\"/>")
            if (le != null) ooxml.append(le!!.ooxml)
            if (layout != null) ooxml.append(layout!!.ooxml)
            if (overlay != null) ooxml.append("<c:overlay val=\"$overlay\"/>")
            if (shapeProps != null) ooxml.append(shapeProps!!.ooxml)
            if (txpr != null) ooxml.append(txpr!!.ooxml)
            ooxml.append("</c:legend>")
            return ooxml.toString()
        }


    constructor(pos: String, overlay: String, l: Layout, le: LegendEntry, sp: SpPr, txpr: TxPr) {
        this.legendpos = pos
        this.le = le
        this.overlay = overlay
        this.layout = l
        this.shapeProps = sp
        this.txpr = txpr
    }

    constructor(l: Legend) {
        this.legendpos = l.legendpos
        this.le = l.le
        this.overlay = l.overlay
        this.layout = l.layout
        this.shapeProps = l.shapeProps
        this.txpr = l.txpr
    }


    /**
     * fill 2003-style legend from OOXML Legend
     *
     * @param l_2003
     */
    fun fill2003Legend(l_2003: io.starter.formats.XLS.charts.Legend) {
        // 0= bottom, 1= corner, 2= top, 3= right, 4= left, 7= not docked
        val pos = arrayOf("b", "tr", "t", "r", "l")
        for (i in pos.indices) {
            if (pos[i] == legendpos) {
                l_2003.legendPosition = i.toShort()
                break
            }
        }
        if (this.hasBox())
            l_2003.addBox()

    }

    override fun cloneElement(): OOXMLElement {
        return Legend(this)
    }

    /**
     * returns true if this legend should have a box around it
     *
     * @return
     */
    fun hasBox(): Boolean {
        return if (this.shapeProps != null) this.shapeProps!!.hasLine() else false
    }

    companion object {

        private val serialVersionUID = 419453456635220517L

        /**
         * parse OOXML legend element
         *
         * @param xpp
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var legendpos: String? = null
            var le: LegendEntry? = null
            var overlay: String? = null
            var layout: Layout? = null
            var shapeProps: SpPr? = null                    // defines the shape properties for the legend for this chart
            var txpr: TxPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "legendPos") {
                            legendpos = xpp.getAttributeValue(0)
                        } else if (tnm == "layout") {
                            lastTag.push(tnm)
                            layout = Layout.parseOOXML(xpp, lastTag) as Layout
                        } else if (tnm == "legendEntry") {
                            lastTag.push(tnm)
                            le = LegendEntry.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "spPr") {
                            lastTag.push(tnm)
                            shapeProps = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //shapeProps.setNS("c");
                        } else if (tnm == "txPr") {
                            lastTag.push(tnm)
                            txpr = TxPr.parseOOXML(xpp, lastTag, bk) as TxPr
                        } else if (tnm == "overlay") {
                            overlay = xpp.getAttributeValue(0)
                        }

                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "legend") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("OOXMLAdapter.parseLegendElement: $e")
            }

            return Legend(legendpos, overlay, layout, le, shapeProps, txpr)
        }

        /**
         * create an OOXML legend from a 2003-vers legend
         *
         * @param l
         */
        fun createLegend(l: io.starter.formats.XLS.charts.Legend): Legend? {

            var ooxmllegend: Legend? = null
            try {
                var sp: SpPr? = null
                sp = SpPr("c")
                sp.setLine(3175, "000000")
                l.fnt
                ooxmllegend = Legend(l.legendPositionString, "1", null, null, sp, null)
            } catch (e: Exception) {
                Logger.logWarn("Error creating 2007+ version Legend: $e")
            }

            return ooxmllegend
        }
    }
}

/**
 * legend Entry element
 *
 *
 * parent: legend
 * attributes;  val= position of legend
 * children:  idx, choice of: delete or txPr
 */
internal class LegendEntry : OOXMLElement {
    private var tx: TxPr? = null
    private var idx = -1
    private var delete: Boolean = false

    override val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<c:legendEntry>")
            tooxml.append("<c:idx val=\"$idx\"/>")
            if (!delete) tooxml.append("<c:delete=\"$delete\">")
            if (tx != null) tooxml.append(tx!!.ooxml)
            tooxml.append("</c:legendEntry>")
            return tooxml.toString()
        }

    constructor(idx: Int, d: Boolean, tx: TxPr) {
        this.idx = idx
        this.delete = d
        this.tx = tx
    }

    constructor(le: LegendEntry) {
        this.idx = le.idx
        this.delete = le.delete
        this.tx = le.tx
    }

    override fun cloneElement(): OOXMLElement {
        return LegendEntry(this)
    }

    companion object {

        private val serialVersionUID = 1859347855337611982L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): LegendEntry {
            var tx: TxPr? = null
            var idx = -1
            var delete = true
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "idx") {
                            idx = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "delete") {
                            delete = java.lang.Boolean.valueOf(xpp.getAttributeValue(0)).booleanValue()
                        } else if (tnm == "txPr") {
                            lastTag.push(tnm)
                            tx = TxPr.parseOOXML(xpp, lastTag, bk) as TxPr
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "legendEntry") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("legendEntry.parseOOXML: $e")
            }

            return LegendEntry(idx, delete, tx)
        }
    }
}
