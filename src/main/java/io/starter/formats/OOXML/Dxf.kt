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

import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.Font
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Arrays


/**
 * dxf (Formatting) OOXML Element
 *
 *
 * A single dxf record, expressing incremental formatting to be applied.
 * Used for Conditional Formatting, Tables, Sort Conditions, Color filters ...
 *
 *
 * Differntial Formatting:
 * define formatting for all non-cell formatting in the workbook. Whereas xf records fully specify a particular aspect of formatting (e.g., cell borders)
 * by referencing those formatting definitions elsewhere in the Styles part, dxf records specify incremental (or
 * differential) aspects of formatting directly inline within the dxf element. The dxf formatting is to be applied on
 * top of or in addition to any formatting already present on the object using the dxf record.
 *
 *
 * parent:  (StyleSheet styles.xml) dxfs
 * chilren: SEQUENCE:  font, numFmt, fill, alignment, border, protection
 */
// TODO: protection element
class Dxf : OOXMLElement {
    /**
     * returns the Font for ths dxf, if any
     *
     * @return
     */
    /**
     * Sts the Font for this dxf from an existing Font
     *
     * @param f
     */
    var font: Font? = null
    private var numFmt: NumFmt? = null
    private var fill: Fill? = null
    private var alignment: Alignment? = null
    private var border: Border? = null
    private var wbh: WorkBookHandle? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<dxf>")
            if (font != null) ooxml.append(font!!.ooxml)
            if (numFmt != null) ooxml.append(numFmt!!.ooxml)
            if (fill != null) ooxml.append(fill!!.getOOXML(true))
            if (alignment != null) ooxml.append(alignment!!.ooxml)
            if (border != null) ooxml.append(border!!.ooxml)
            ooxml.append("</dxf>")
            return ooxml.toString()
        }

    val borderColors: IntArray?
        get() = if (border != null) border!!.borderColorInts else null

    val borderStyles: IntArray?
        get() {
            if (border != null) {
                val styles = border!!.borderStyles
                return if (Arrays.equals(intArrayOf(0, 0, 0, 0), styles)) null else styles
            }
            return null
        }

    val borderSizes: IntArray?
        get() = if (border != null) border!!.borderSizes else null

    val isStriken: Boolean
        get() = if (font != null) font!!.stricken else false

    val fg: Int
        get() = if (fill != null) fill!!.getFgColorAsInt(wbh!!.workBook!!.theme) else -1

    val fillPatternInt: Int
        get() = if (fill != null) fill!!.fillPatternInt else -1

    val bg: Int
        get() = if (fill != null) fill!!.getBgColorAsInt(wbh!!.workBook!!.theme) else -1

    val bgColorAsString: String?
        get() = if (fill != null) fill!!.getBgColorAsRGB(wbh!!.workBook!!.theme) else null


    val horizontalAlign: String?
        get() = if (alignment != null) alignment!!.getAlignment("horizontal") else null

    val verticalAlign: String?
        get() = if (alignment != null) alignment!!.getAlignment("vertical") else null

    val numberFormat: String?
        get() = if (numFmt != null) numFmt!!.formatId else null

    val fontHeight: Int
        get() = if (font != null) font!!.fontHeight else -1

    val fontWeight: Int
        get() = if (font != null) font!!.fontWeight else -1

    val fontName: String?
        get() = if (font != null) font!!.fontName else null

    val fontColor: Int
        get() = if (font != null) font!!.color else -1

    val isItalic: Boolean
        get() = if (font != null) font!!.italic else false

    val fontUnderline: Int
        get() = if (font != null) font!!.underlineStyle else -1

    /**
     * return a String representation of this Dxf in "style properties" notation
     *
     * @return String representation of this Dxf
     * @see Cf.setStylePropsFromString
     */
    // fill
    // fg is pattern color
    // font
    // note: since this is differential, many of these may not be set
    // TODO: italic, bold
    // borders
    // alignment
    // number format
    val styleProps: String
        get() {
            val props = StringBuffer()
            if (fill != null) {
                props.append("pattern:" + fill!!.fillPatternInt + ";")
                var s = fill!!.getFgColorAsRGB(wbh!!.workBook!!.theme)
                if (s != null)
                    props.append("patterncolor:#$s;")
                s = fill!!.getBgColorAsRGB(wbh!!.workBook!!.theme)
                if (s != null)
                    props.append("background:#$s;")
            }
            if (font != null) {
                if (font!!.fontName != "")
                    props.append("font-name" + font!!.fontName + ";")
                if (font!!.fontWeight > -1)
                    props.append("font-weight:" + font!!.fontWeight + ";")
                if (font!!.fontHeight > -1)
                    props.append("font-Height:" + font!!.fontHeight + ";")
                props.append("font-ColorIndex:" + font!!.color + ";")
                if (font!!.stricken) props.append("font-Striken:" + font!!.stricken + ";")
                if (font!!.italic) props.append("font-italic:" + font!!.italic + ";")
                if (font!!.underlineStyle != 0) props.append("font-UnderlineStyle:" + font!!.underlineStyle + ";")
            }
            if (border != null) {
                val sizes = border!!.borderSizes
                val styles = border!!.borderStyles
                val colors = border!!.borderColors
                props.append("border-top:" + sizes[0] + " " + FormatHandle.BORDER_NAMES[styles[0]] + " " + colors[0] + ";")
                props.append("border-left:" + sizes[1] + " " + FormatHandle.BORDER_NAMES[styles[1]] + " " + colors[1] + ";")
                props.append("border-bottom:" + sizes[2] + " " + FormatHandle.BORDER_NAMES[styles[2]] + " " + colors[2] + ";")
                props.append("border-right:" + sizes[3] + " " + FormatHandle.BORDER_NAMES[styles[3]] + " " + colors[3] + ";")
            }
            if (alignment != null) {
                var s = alignment!!.getAlignment("vertical")
                if (s != null)
                    props.append("alignment-vertical$s;")
                s = alignment!!.getAlignment("horizontal")
                if (s != null)
                    props.append("alignment-horizontal$s;")
            }
            if (numFmt != null) {
                val s = numFmt!!.formatId
                props.append("numberformat:$s;")

            }

            return props.toString()
        }

    constructor(fnt: Font, nf: NumFmt, f: Fill, a: Alignment, b: Border, wbh: WorkBookHandle) {
        this.font = fnt
        this.numFmt = nf
        this.fill = f
        this.alignment = a
        this.border = b
        this.wbh = wbh
    }

    constructor(d: Dxf) {
        this.font = d.font
        this.numFmt = d.numFmt
        this.fill = d.fill
        this.alignment = d.alignment
        this.border = d.border
        this.wbh = d.wbh
    }

    constructor() {}

    /**
     * returns the OOXML Fill element
     *
     * @return
     */
    fun getFill(): Fill? {
        return fill
    }


    /**
     * for BIFF8->OOXML Compatiblity, create a dxf from Cf style info
     */
    fun createFont(w: Int, i: Boolean, ustyle: Int, cl: Int, h: Int) {
        font = Font("", w, h)
        if (w == 700)
            font!!.bold = true    // why doesn't constructor do this?
        if (i) font!!.italic = i
        if (ustyle != 0) font!!.setUnderlineStyle(ustyle.toByte())
        font!!.color = cl
    }

    /**
     * Sets the fill for this dxf from an existing Fill element
     *
     * @param f
     */
    fun setFill(f: Fill) {
        this.fill = f.cloneElement() as Fill
    }

    /**
     * for BIFF8->OOXML Compatiblity, create a dxf from Cf style info
     */
    fun createFill(fs: Int, fg: Int, bg: Int, bk: WorkBookHandle) {
        if (fs < 0 || fs > OOXMLConstants.patternFill.size)
            this.fill = Fill(null, fg, bg, bk.workBook!!.theme)    // meaning it's the default (solid bg) pattern
        else
            this.fill = Fill(OOXMLConstants.patternFill[fs], fg, bg, bk.workBook!!.theme)
    }

    /**
     * for BIFF8->OOXML Compatiblity, create a dxf from Cf style info
     */
    fun createBorder(bk: WorkBookHandle, styles: IntArray, colors: IntArray) {
        border = Border(bk, styles, colors)
    }

    override fun cloneElement(): OOXMLElement {
        return Dxf(this)
    }

    companion object {

        private val serialVersionUID = -5999328795988018131L

        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): OOXMLElement {
            var fnt: Font? = null
            var nf: NumFmt? = null
            var f: Fill? = null
            var a: Alignment? = null
            var b: Border? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "font") {
                            fnt = Font.parseOOXML(xpp, bk)
                        } else if (tnm == "numFmt") {
                            nf = NumFmt.parseOOXML(xpp) as NumFmt
                        } else if (tnm == "fill") {
                            f = Fill.parseOOXML(xpp, true, bk) as Fill
                        } else if (tnm == "alignment") {
                            a = Alignment.parseOOXML(xpp) as Alignment
                        } else if (tnm == "border") {
                            b = Border.parseOOXML(xpp, bk) as Border
                        } else if (tnm == "protection") {
                            // TODO: finish
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "dxf") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("dxf.parseOOXML: $e")
            }

            return Dxf(fnt, nf, f, a, b, bk)
        }
    }
}

