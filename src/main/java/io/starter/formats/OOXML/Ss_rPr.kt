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

import io.starter.OpenXLS.DocumentHandle
import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.Font
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.HashMap

/**
 * rPr (Run Properties for Shared Strings (see SharedStrings.xml))
 *
 *
 * This element represents a set of properties to apply to the contents of this rich text run.
 * This element corresponds to Unicode String formatting runs where a specific font is applied to a sub-section of a string
 *
 *
 * parent:  r (rich text run)
 * children:  many
 */
class Ss_rPr : OOXMLElement {
    private var attrs = HashMap<String, String>()
    private var color: Color? = null

    override// attributes
    // same as true for <b/> <u/> <i/> ...
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<rPr>")
            val i = attrs.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs[key]
                if (`val` == "")
                    ooxml.append("<$key/>")
                else
                    ooxml.append("<$key val=\"$`val`\"/>")
            }
            if (color != null) ooxml.append(color!!.ooxml)
            ooxml.append("</rPr>")
            return ooxml.toString()
        }

    constructor() {

    }

    constructor(attrs: HashMap<String, String>, c: Color) {
        this.attrs = attrs
        this.color = c
    }

    constructor(r: Ss_rPr) {
        this.attrs = r.attrs
        this.color = r.color
    }

    override fun cloneElement(): OOXMLElement {
        return Ss_rPr(this)
    }

    /**
     * retrieve one of the following key values, if already set
     * b, charset, condense, extend, family, i, outline, rFont, scheme, shadow, strike, sz, u, vertAlign
     *
     * @param key
     * @return
     */
    fun getAttr(key: String): String {
        return attrs[key]
    }

    /**
     * set the value for one of the rPr children:
     * b, charset, condense, extend, family, i, outline, rFont, scheme, shadow, strike, sz, u, vertAlign
     *
     *
     * most are boolean string values "0" or "1"
     * except for sz, rFont, family, scheme
     *
     * @param key
     * @param val
     */
    fun setAttr(key: String, `val`: String?) {
        attrs[key] = `val`
    }

    /** family:
     * 0 Not applicable.
     * 1 Roman
     * 2 Swiss
     * SpreadsheetML Reference Material - Styles
     * 2113
     * Value Font Family
     * 3 Modern
     * 4 Script
     * 5 Decorative
     */
    /**
     * charset (Character Set)
     * This element defines the font character set of this font.
     * This field is used in font creation and selection if a font of the given facename is not available on the system.
     * Although it is not required to have around when resolving font facename, the information can be stored for
     * when needed to help resolve which font face to use of all available fonts on a system.
     *
     * int 0-255.
     *
     * The following are some of the possible the character sets:
     * INT
     * Value
     * Character Set
     * 0 ANSI_CHARSET
     * 1 DEFAULT_CHARSET
     * 2 SYMBOL_CHARSET
     * 77 MAC_CHARSET
     * 128 SHIFTJIS_CHARSET
     * 129 HANGEUL_CHARSET
     * 129 HANGUL_CHARSET
     * 130 JOHAB_CHARSET
     * 134 GB2312_CHARSET
     * 136 CHINESEBIG5_CHARSET
     * 161 GREEK_CHARSET
     * 162 TURKISH_CHARSET
     * 163 VIETNAMESE_CHARSET
     * 177 HEBREW_CHARSET
     * 178 ARABIC_CHARSET
     * 186 BALTIC_CHARSET
     * 204 RUSSIAN_CHARSET
     * 222 THAI_CHARSET
     * 238 EASTEUROPE_CHARSET
     * 255 OEM_CHARSET
     */
    /**
     * condense (Condense)
     * Macintosh compatibility setting. Represents special word/character rendering on Macintosh, when this flag is
     * set. The effect is to condense the text (squeeze it together). SpreadsheetML applications are not required to
     * render according to this flag.
     */
    /**
     * extend (Extend)
     * This element specifies a compatibility setting used for previous spreadsheet applications, resulting in special
     * word/character rendering on those legacy applications, when this flag is set. The effect extends or stretches out
     * the text. SpreadsheetML applications are not required to render according to this flag.
     */
    /**
     * shadow (Shadow)
     * Macintosh compatibility setting. Represents special word/character rendering on Macintosh, when this flag is
     * set. The effect is to render a shadow behind, beneath and to the right of the text. SpreadsheetML applications
     * are not required to render according to this flag.
     */
    /**
     * outline (Outline)
     * This element displays only the inner and outer borders of each character. This is very similar to Bold in behavior
     */
    /**
     * vertAlign (Vertical Alignment)
     * This element adjusts the vertical position of the text relative to the text's default appearance for this run. It is
     * used to get 'superscript' or 'subscript' texts, and shall reduce the font size (if a smaller size is available)
     * accordingly.
     *
     * val= An enumeration representing the vertical-alignment setting.
     * baseline, subscript, superscript
     * Setting this to either subscript or superscript shall make the font size smaller if a
     * smaller font size is available.
     */

    /**
     * given an rPr OOXML text run properties, create an OpenXLS font
     *
     * @param bk
     * @return
     */
    fun generateFont(bk: DocumentHandle): Font {
        // not using attributes:  charset, family, condense, extend, shadow, scheme, outline==bold
        val f = Font("Arial", 400, 200)
        var o: Any?

        o = this.getAttr("rFont")
        f.fontName = (o as String?)!!
        o = this.getAttr("sz")
        if (o != null) f.fontHeight = Font.PointsToFontHeight(java.lang.Double.parseDouble((o as String?)!!))

        // boolean attributes
        o = this.getAttr("b")
        if (o != null) f.bold = true
        o = this.getAttr("i")
        if (o != null) f.italic = true
        o = this.getAttr("u")
        if (o != null) f.underlined = true
        o = this.getAttr("strike")
        if (o != null) f.stricken = true
        o = this.getAttr("outline")
        if (o != null) f.bold = true
        o = this.getAttr("vertAlign")
        if (o != null) {
            val s = o as String?
            if (s == "baseline")
                f.script = 0
            else if (s == "superscript")
                f.script = 1
            else if (s == "subscript")
                f.script = 2
        }
        f.ooxmlColor = color
        return f
    }

    companion object {

        private val serialVersionUID = 8940630588129002652L

        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): OOXMLElement {
            val attrs = HashMap<String, String>()
            var c: Color? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "rFont") {
                            attrs["rFont"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "charset") {
                            attrs["charset"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "family") {
                            attrs["family"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "b") {
                            attrs["b"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "i") {
                            attrs["i"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "strike") {
                            attrs["strike"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "outline") {
                            attrs["outline"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "shadow") {
                            attrs["shadow"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "condense") {
                            attrs["condense"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "extend") {
                            attrs["extend"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "sz") {
                            attrs["sz"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "u") {
                            attrs["u"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "vertAlign") {
                            attrs["vertAlign"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "scheme") {
                            attrs["scheme"] = if (xpp.attributeCount > 0) xpp.getAttributeValue(0) else ""    // val
                        } else if (tnm == "color") {
                            c = Color.parseOOXML(xpp, FormatHandle.colorFONT, bk) as Color
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "rPr") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("rPr.parseOOXML: $e")
            }

            return Ss_rPr(attrs, c)
        }

        /**
         * create a new OOXML ss_rPr shared string table text run properties object using attributes from Font f
         *
         * @param f
         * @return
         */
        fun createFromFont(f: Font): Ss_rPr {
            // not using attributes:  charset, family, condense, extend, shadow, scheme, outline==bold
            val rp = Ss_rPr()
            rp.setAttr("rFont", f.fontName)
            rp.setAttr("sz", f.fontHeightInPoints.toString())

            // boolean attributes
            if (f.bold)
                rp.setAttr("b", "")
            if (f.italic)
                rp.setAttr("i", "")
            if (f.underlined)
                rp.setAttr("u", "")
            if (f.stricken)
                rp.setAttr("strike", "")
            val s = f.script
            if (s == 1)
                rp.setAttr("vertAlign", "superscript")
            else if (s == 2)
                rp.setAttr("vertAlign", "subscript")
            rp.color = f.ooxmlColor
            return rp
        }
    }
}

