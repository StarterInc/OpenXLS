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
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.WorkBook
import io.starter.formats.XLS.charts.TextDisp
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.HashMap
import java.util.Stack

/**
 * class holds OOXML title Property used to define chart and axis titles
 */
class Title : OOXMLElement {
    var layout: Layout? = null
        private set
    var spPr: SpPr? = null
        private set
    private var chartText: ChartText? = null // tx
    private val txpr: TxPr? = null // xPr

    /**
     * generate ooxml to define a title
     *
     * @return
     */
    /**
     * tx chart text layout overlay spPr txPr
     */
    override// TODO: overlay
    val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<c:title>")
            if (this.chartText != null)
                tooxml.append(chartText!!.ooxml)
            if (this.layout != null)
                tooxml.append(layout!!.ooxml)
            if (this.spPr != null)
                tooxml.append(spPr!!.ooxml)
            if (this.txpr != null)
                tooxml.append(this.txpr.ooxml)
            tooxml.append("</c:title>")
            return tooxml.toString()
        }

    val title: String?
        get() = if (chartText != null) chartText!!.title else ""

    constructor(ct: ChartText, txpr: TxPr, l: Layout, sp: SpPr) {
        this.layout = l
        this.spPr = sp
        this.chartText = ct
        this.txpr = txpr
    }

    /**
     * for BIFF8 compatibility, create a Title element from the title string
     *
     * @param t
     */
    constructor(t: String) {
        this.chartText = ChartText(t)
        this.spPr = SpPr("c")
        // no spPr
    }

    /**
     * create an OOXML title from a 2003-v TextDisp record
     *
     * @param td
     */
    constructor(td: TextDisp, bk: WorkBook) {
        val para = P(td.getFont(bk)!!, td.toString())
        this.chartText = ChartText(null, para, null)
    }

    fun setLayout(x: Double, y: Double) {
        this.layout = Layout(null, doubleArrayOf(x, y, -1.0, -1.0))
    }

    override fun cloneElement(): OOXMLElement {
        return Title(this.chartText, this.txpr, this.layout, this.spPr)
    }

    /**
     * return the font index for this title
     *
     * @param wb
     * @return
     */
    fun getFontId(wb: WorkBookHandle): Int {
        return if (chartText != null) chartText!!.getFontId(wb) else -1
    }

    companion object {

        private val serialVersionUID = -3889674575558708481L

        /**
         * parse title OOXML element title
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spPr object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var txpr: TxPr? = null
            var ct: ChartText? = null
            var l: Layout? = null
            var sp: SpPr? = null

            /*
         * TextDisp td= (TextDisp) TextDisp.getPrototype(ObjectLink.TYPE_TITLE,
         * str, this.getWorkBook()); this.addChartRecord((BiffRec) td); // add
         * TextDisp title to end of chart recs ... charttitle= td;
         */
            /**
             * contains (in Sequence) layout overlay -- not handled yet spPr tx txPr
             */
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "tx") { // chart text
                            lastTag.push(tnm)
                            ct = ChartText.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "manualLayout") {
                            lastTag.push(tnm)
                            l = Layout.parseOOXML(xpp, lastTag) as Layout
                        } else if (tnm == "spPr") {
                            lastTag.push(tnm)
                            sp = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            // sp.setNS("c");
                        } else if (tnm == "txPr") {
                            lastTag.push(tnm)
                            txpr = TxPr.parseOOXML(xpp, lastTag, bk) as TxPr
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "title") {
                            lastTag.pop() // pop title tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("title.parseOOXML: $e")
            }

            return Title(ct, txpr, l, sp)
        }
    }
}

/**
 * chart text element tx
 */

/**
 * contains either strRef -- contains f, strCache rich -- contains bodyPr,
 * lstStyle, p
 */
internal class ChartText : OOXMLElement {
    var strref: StrRef? = null
    var para: P? = null
    var bpr: BodyPr? = null

    override// chart text
    // it has a rich element
    // TODO: Handle!!!
    // text paragraph
    // it has a strRef element
    val ooxml: String
        get() {
            val cooxml = StringBuffer()
            cooxml.append("<c:tx>")
            if (strref == null) {
                cooxml.append("<c:rich>")
                if (bpr != null)
                    cooxml.append(bpr!!.ooxml)
                else
                    cooxml.append("<a:bodyPr/>")
                cooxml.append("<a:lstStyle/>")
                if (para != null)
                    cooxml.append(para!!.ooxml)
                cooxml.append("</c:rich>")
            } else {
                cooxml.append(strref!!.ooxml)
            }
            cooxml.append("</c:tx>")
            return cooxml.toString()
        }

    val title: String?
        get() = if (para != null) para!!.title else ""

    /**
     * for BIFF8 compatibility, create an OOXML chartText element from title
     * string
     *
     * @param s
     */
    constructor(s: String) {
        this.para = P(s)
    }

    constructor(s: StrRef, para: P, bpr: BodyPr) {
        this.strref = s
        this.para = para
        this.bpr = bpr
    }

    override fun cloneElement(): OOXMLElement {
        return ChartText(this.strref, this.para, this.bpr)
    }


    /**
     * concatenate the 3 levels of text properties and either find an existing
     * font or add new
     *
     * @return
     */
    fun getFontId(wb: WorkBookHandle): Int {
        var textprops: HashMap<*, *> = HashMap<Any, Any>()
        if (bpr != null) {
            /*
             * noAutofit (No AutoFit) §5.1.5.1.2 normAutofit (Normal AutoFit)
             * §5.1.5.1.3 prstTxWarp (Preset Text Warp) §5.1.11.19 scene3d (3D
             * Scene Properties) §5.1.4.1.26 sp3d (Apply 3D shape properties)
             * §5.1.7.12 spAutoFit (Shape AutoFit)
             */
            // TODO: set any of the above??
        }
        if (para != null) {
            textprops = para!!.textProperties
        }
        /*
         * altLang (Alternative Language) b (Bold) bool baseline (Baseline) bmk
         * (Bookmark Link Target) cap (Capitalization) i (Italics) bool kern
         * (Kerning) kumimoji lang (Language ID) spc (Spacing) strike
         * (Strikethrough) sz (Font Size) size u (Underline) underline style
         *
         * PLUS-- fill, line, blipFill, cs/ea/latin font (attribute: typeface)
         * ****
         */
        val w = 400
        val u = 0
        var h = 200.0 // default
        var b = false
        var i = false
        var face = "Arial"
        if (textprops.get("b") != null)
            b = "1" == textprops.get("b")
        if (textprops.get("i") != null)
            i = "1" == textprops.get("i")
        if (textprops.get("latin_typeface") != null)
            face = textprops.get("latin_typeface") as String
        // if (textprops.get("u")!=null)
        // u= textprops.get("u").toString();
        var o: Any? = textprops.get("sz") // Whole points are specified in
        // increments of 100 starting with 100
        // being a point size of 1
        if (o != null)
            h = Font.PointsToFontHeight((Integer.parseInt((o as String?)!!) / 100).toDouble()).toDouble()
        val f = Font(face, w, h.toInt())
        if (b)
            f.bold = true
        if (i)
            f.italic = i
        if (u != 0)
            f.setUnderlineStyle(u.toByte())
        o = textprops.get("vertAlign")
        if (o != null) {
            val s = o as String?
            if (s == "baseline")
                f.script = 0
            else if (s == "superscript")
                f.script = 1
            else if (s == "subscript")
                f.script = 2
        }
        o = textprops.get("strike")
        if (o != null)
            f.stricken = true
        // f.setFontColor(cl);
        return io.starter.OpenXLS.FormatHandle.addFont(f, wb)
    }

    companion object {

        private val serialVersionUID = -1175394918747218776L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): ChartText {
            var para: P? = null
            var bpr: BodyPr? = null
            var s: StrRef? = null

            try { // title->tx->rich->bodyPr lstStyle, p->pPr, r
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "p") { // text Paragraph props - part of rich
                            lastTag.push(tnm)
                            para = P.parseOOXML(xpp, lastTag, bk) as P
                        } else if (tnm == "bodyPr") { // body-level Paragraph props -- part of rich
                            lastTag.push(tnm)
                            bpr = BodyPr.parseOOXML(xpp, lastTag) as BodyPr
                        } else if (tnm == "strRef") {
                            lastTag.push(tnm)
                            s = StrRef.parseOOXML(xpp, lastTag) as StrRef
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "tx") {
                            lastTag.pop() // pop title tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("chartText.parseOOXML: $e")
            }

            return ChartText(s, para, bpr)
        }
    }

}
