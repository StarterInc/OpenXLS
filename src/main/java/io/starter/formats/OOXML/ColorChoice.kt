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
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.HashMap
import java.util.Stack

/**
 * ColorChoice: choice of: hslClr (Hue, Saturation, Luminance Color Model)
 * prstClr (Preset Color) §5.1.2.2.22 schemeClr (Scheme Color) §5.1.2.2.29
 * scrgbClr (RGB Color Model - Percentage Variant) srgbClr (RGB Color Model -
 * Hex Variant) sysClr (System Color)
 */
// TODO: FINISH: child elements governing color transformations
// finish hslClr
class ColorChoice : OOXMLElement {
    private var s: SchemeClr? = null
    private var srgb: SrgbClr? = null
    private var sys: SysClr? = null
    private var scrgb: ScrgbClr? = null
    private var p: PrstClr? = null
    var theme: Theme? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (s != null)
                ooxml.append(s!!.ooxml)
            else if (sys != null)
                ooxml.append(sys!!.ooxml)
            else if (srgb != null)
                ooxml.append(srgb!!.ooxml)
            else if (scrgb != null)
                ooxml.append(scrgb!!.ooxml)
            else if (p != null)
                ooxml.append(p!!.ooxml)
            return ooxml.toString()
        }

    val color: Int
        get() {
            if (s != null)
                return s!!.color
            else if (sys != null)
                sys!!.color
            else if (srgb != null)
                srgb!!.color
            else if (scrgb != null)
                scrgb!!.color
            else if (p != null)
                p!!.color
            return -1

        }

    constructor(s: SchemeClr, srgb: SrgbClr, sys: SysClr, scrgb: ScrgbClr,
                p: PrstClr) {
        this.s = s
        this.srgb = srgb
        this.sys = sys
        this.scrgb = scrgb
        this.p = p
    }

    private fun setTheme(t: Theme?) {
        this.theme = t
    }

    constructor(c: ColorChoice) {
        this.s = c.s
        this.srgb = c.srgb
        this.sys = c.sys
        this.scrgb = c.scrgb
        this.p = c.p
    }

    override fun cloneElement(): OOXMLElement {
        return ColorChoice(this)
    }

    companion object {

        private val serialVersionUID = -4117811305941771643L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var s: SchemeClr? = null
            var srgb: SrgbClr? = null
            var sys: SysClr? = null
            var scrgb: ScrgbClr? = null
            var p: PrstClr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "schemeClr") {
                            lastTag.push(tnm)
                            s = SchemeClr.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        } else if (tnm == "srgbClr") {
                            lastTag.push(tnm)
                            srgb = SrgbClr.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        } else if (tnm == "sysClr") {
                            lastTag.push(tnm)
                            sys = SysClr.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        } else if (tnm == "scrgbClr") {
                            lastTag.push(tnm)
                            scrgb = ScrgbClr.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        } else if (tnm == "prstClr") {
                            lastTag.push(tnm)
                            p = PrstClr.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        }
                        /* tnm.equals("hslClr") */// TODO: finish

                    } else if (eventType == XmlPullParser.END_TAG) { // shouldn't
                        // get here
                        break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("ColorChoice.parseOOXML: $e")
            }

            val c = ColorChoice(s, srgb, sys, scrgb, p)
            c.setTheme(bk.workBook!!.theme)
            return c
        }
    }

}

/**
 * schemeClr (Scheme Color) This element specifies a color bound to a user's
 * theme. As with all elements which define a color, it is possible to apply a
 * list of color transforms to the base color defined.
 *
 *
 * accent1 (Accent Color 1) Extra scheme color 1 accent2 (Accent Color 2) Extra
 * scheme color 2 accent3 (Accent Color 3) Extra scheme color 3 accent4 (Accent
 * Color 4) Extra scheme color 4 accent5 (Accent Color 5) Extra scheme color 5
 * accent6 (Accent Color 6) Extra scheme color 6 bg1 (Background Color 1)
 * Semantic background color bg2 (Background Color 2) Semantic additional
 * background color dk1 (Dark Color 1) Main dark color 1 dk2 (Dark Color 2) Main
 * dark color 2 folHlink (Followed Hyperlink Color) Followed Hyperlink Color
 * hlink (Hyperlink Color) Regular Hyperlink Color lt1 (Light Color 1) Main
 * Light Color 1 lt2 (Light Color 2) Main Light Color 2 phClr (Style Color) A
 * color used in theme definitions which means to use the color of the style.
 * tx1 (Text Color 1) Semantic text color tx2 (Text Color 2) Semantic additional
 * text color
 *
 *
 * parent: many children: many - TODO: handle color transformation children
 * (alpha ...)
 */
internal class SchemeClr : OOXMLElement {
    private var `val`: String? = null
    private var clrTransform: ColorTransform? = null
    private var theme: Theme? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:schemeClr val=\"$`val`\">")
            if (clrTransform != null)
                ooxml.append(clrTransform!!.ooxml)
            ooxml.append("</a:schemeClr>")
            return ooxml.toString()
        }

    val color: Int
        get() {
            val tint = (if (clrTransform == null) 0 else clrTransform!!.tint).toDouble()
            val o = Color.parseThemeColor(`val`, tint, 0.toShort(), this.theme)
            return o[0] as Int
        }

    constructor(`val`: String, clrTransform: ColorTransform, t: Theme) {
        this.`val` = `val`
        this.clrTransform = clrTransform
        this.theme = t
    }

    constructor(sc: SchemeClr) {
        this.`val` = sc.`val`
        this.clrTransform = sc.clrTransform
        this.theme = sc.theme
    }

    override fun cloneElement(): OOXMLElement {
        return SchemeClr(this)
    }

    companion object {

        private val serialVersionUID = 2127868578801669266L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): SchemeClr {
            var `val`: String? = null
            var clrTransform: ColorTransform? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "schemeClr") {
                            `val` = xpp.getAttributeValue(0)
                        } else {
                            clrTransform = ColorTransform.parseOOXML(xpp, lastTag)
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "schemeClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("schemeClr.parseOOXML: $e")
            }

            return SchemeClr(`val`, clrTransform, bk.workBook!!.theme)
        }
    }
}

/**
 * sysClr (System Color) This element specifies a color bound to predefined
 * operating system elements. // TODO: appropriate to hard-code???
 *
 *
 * parent: many children: COLORSTRANSFORM
 */
internal class SysClr : OOXMLElement {
    private var `val`: String? = null // This simple type specifies a system color
    // value. This color is based upon the value
    // that this color currently has within the
    // system on which the document is being viewed.
    private var lastClr: String? = null // Specifies the color value that was last
    // computed by the generating application.
    // Applications shall use the lastClr
    // attribute to determine the absolute value
    // of the last color used if system colors
    // are not supported.
    private var clrTransform: ColorTransform? = null

    override// TODO: Handle child elements
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:sysClr val=\"$`val`\"")
            if (lastClr != null)
                ooxml.append(" lastClr=\"$lastClr\">")
            if (clrTransform != null)
                ooxml.append(clrTransform!!.ooxml)
            ooxml.append("</a:sysClr>")
            return ooxml.toString()
        }

    /**
     * return the color int that represents this system color
     *
     * @return
     */
    val color: Int
        get() {
            for (i in OOXMLConstants.systemColors.indices) {
                if (OOXMLConstants.systemColors[i][0] == `val`) {
                    return FormatHandle.HexStringToColorInt(
                            OOXMLConstants.systemColors[i][1], 0.toShort())
                }
            }
            return -1
        }

    constructor(`val`: String, lastClr: String, clrTransform: ColorTransform) {
        this.`val` = `val`
        this.lastClr = lastClr
        this.clrTransform = clrTransform
    }

    constructor(sc: SysClr) {
        this.`val` = sc.`val`
        this.lastClr = sc.lastClr
        this.clrTransform = sc.clrTransform
    }

    override fun cloneElement(): OOXMLElement {
        return SysClr(this)
    }

    companion object {

        private val serialVersionUID = 8307422721346337409L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): SysClr {
            var `val`: String? = null
            var lastClr: String? = null
            var clrTransform: ColorTransform? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "sysClr") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "val") { //
                                    `val` = xpp.getAttributeValue(i)
                                } else if (nm == "lastClr") {
                                    lastClr = xpp.getAttributeValue(i)
                                }
                            }
                        } else {
                            clrTransform = ColorTransform.parseOOXML(xpp, lastTag)
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "sysClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("sysClr.parseOOXML: $e")
            }

            return SysClr(`val`, lastClr, clrTransform)
        }
    }
}

/**
 * srgbClr (RGB Color Model - Hex Variant)
 *
 *
 * This element specifies a color using the red, green, blue RGB color model.
 * Red, green, and blue is expressed as sequence of hex digits, RRGGBB. A
 * perceptual gamma of 2.2 is used. Specifies the level of red as expressed by a
 * percentage offset increase or decrease relative to the input color.
 *
 *
 * parent: many children: COLORSTRANSFORM
 */
internal class SrgbClr : OOXMLElement {
    private var `val`: String? = null
    private var clrTransform: ColorTransform? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:srgbClr val=\"$`val`\">")
            if (clrTransform != null)
                ooxml.append(clrTransform!!.ooxml)
            ooxml.append("</a:srgbClr>")
            return ooxml.toString()
        }

    /**
     * interpret val and return color int from color table
     *
     * @return
     */
    val color: Int
        get() = FormatHandle.HexStringToColorInt(`val`!!, 0.toShort())

    constructor(`val`: String, clrTransform: ColorTransform) {
        this.`val` = `val`
        this.clrTransform = clrTransform
    }

    constructor(sc: SrgbClr) {
        this.`val` = sc.`val`
        this.clrTransform = sc.clrTransform
    }

    override fun cloneElement(): OOXMLElement {
        return SrgbClr(this)
    }

    companion object {

        private val serialVersionUID = -999813417659560045L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): SrgbClr {
            var `val`: String? = null
            var clrTransform: ColorTransform? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "srgbClr") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "val") {
                                    `val` = xpp.getAttributeValue(i)
                                }
                            }
                        } else {
                            clrTransform = ColorTransform.parseOOXML(xpp, lastTag)
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "srgbClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("srgbClr.parseOOXML: $e")
            }

            return SrgbClr(`val`, clrTransform)
        }
    }
}

/**
 * scrgbClr (Scheme Color)
 *
 *
 * This element specifies a color using the red, green, blue RGB color model.
 * Each component, red, green, and blue is expressed as a percentage from 0% to
 * 100%. A linear gamma of 1.0 is assumed.
 *
 *
 * parent: many children: COLORSTRANSFORM
 */
internal class ScrgbClr : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var clrTransform: ColorTransform? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:scrgbClr")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (clrTransform != null)
                ooxml.append(clrTransform!!.ooxml)
            ooxml.append("</a:scrgbClr>")
            return ooxml.toString()
        }

    /**
     * interpret the rbg value into a color int/index into color table
     *
     * @return
     */
    // r, g, b in percentages (in 1000th of a percentage)
    // perecentage
    val color: Int
        get() {
            var rval = (Integer.valueOf(attrs!!["r"]) / 100000).toDouble()
            rval *= 255.0
            var gval = (Integer.valueOf(attrs!!["g"]) / 100000).toDouble()
            gval *= 255.0
            var bval = (Integer.valueOf(attrs!!["b"]) / 100000).toDouble()
            bval *= 255.0
            if (clrTransform != null)
                Logger.logWarn("Scheme Color must process color transforms")
            val c = java.awt.Color(rval.toInt(), gval.toInt(),
                    bval.toInt())
            return FormatHandle.getColorInt(c)
        }

    constructor(attrs: HashMap<String, String>, clrTransform: ColorTransform) {
        this.attrs = attrs
        this.clrTransform = clrTransform
    }

    constructor(sc: ScrgbClr) {
        this.attrs = sc.attrs
        this.clrTransform = sc.clrTransform
    }

    override fun cloneElement(): OOXMLElement {
        return ScrgbClr(this)
    }

    companion object {

        private val serialVersionUID = -8782954669829478560L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): ScrgbClr {
            val attrs = HashMap<String, String>()
            var clrTransform: ColorTransform? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "scrgbClr") {
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i) // r, g, b
                        } else {
                            clrTransform = ColorTransform.parseOOXML(xpp, lastTag)
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "scrgbClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("scrgbClr.parseOOXML: $e")
            }

            return ScrgbClr(attrs, clrTransform)
        }
    }
}

/**
 * prstClr (Preset Color) This element specifies a color which is bound to one
 * of a predefined collection of colors.
 *
 *
 * parent: many children: many
 */
internal class PrstClr : OOXMLElement {
    private var `val`: String? = null
    private var clrTransform: ColorTransform? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:prstClr val=\"$`val`\">")
            if (clrTransform != null)
                ooxml.append(clrTransform!!.ooxml)
            ooxml.append("</a:prstClr>")
            return ooxml.toString()
        }

    val color: Int
        get() {
            if (clrTransform != null)
                Logger.logWarn("Preset Color must process color transforms")
            for (i in FormatHandle.COLORNAMES.indices) {
                if (FormatHandle.COLORNAMES[i].equals(`val`!!, ignoreCase = true))
                    return i
            }
            return -1
        }

    constructor(`val`: String, clrTransform: ColorTransform) {
        this.`val` = `val`
        this.clrTransform = clrTransform
    }

    constructor(sc: PrstClr) {
        this.`val` = sc.`val`
        this.clrTransform = sc.clrTransform
    }

    override fun cloneElement(): OOXMLElement {
        return PrstClr(this)
    }

    companion object {

        private val serialVersionUID = -5773022185972396279L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): PrstClr {
            var `val`: String? = null
            var clrTransform: ColorTransform? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "prstClr") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "val") {
                                    `val` = xpp.getAttributeValue(i)
                                }
                            }
                        } else {
                            clrTransform = ColorTransform.parseOOXML(xpp, lastTag)
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "prstClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("prstClr.parseOOXML: $e")
            }

            return PrstClr(`val`, clrTransform)
        }
    }
}

/**
 * common Color Transformes used by parent elements: schemeColor, systemColor,
 * hslColor, presetColor, sRgbColor, scRgbColor color adjustments are in
 * percentage units <br></br>
 * // TODO: Finish // comp // inv // gamma // invGamma // gray, red, green, blue
 */
internal class ColorTransform(private val lum: IntArray?, private val hue: IntArray?, private val sat: IntArray?, private val alpha: IntArray?,
        // TODO: Finish
        // comp
        // inv
        // gamma
        // gray, red, green, blue

                              val tint: Int, private val shade: Int) {

    /**
     * returns the OOXML associated with color transforms of a parent color
     * element <br></br>
     * note that these color transforms must be part of either <br></br>
     * schemeColor, systemColor, hslColor, presetColor, sRgbColor, scRgbColor
     *
     * @return
     */
    /**
     * a:tint Tint a:shade Shade a:comp Complement a:inv Inverse a:gray Gray
     * a:alpha Alpha a:alphaOff Alpha Offset a:alphaMod Alpha Modulation
     * a:hue Hue a:hueOff Hue Offset a:hueMod Hue Modulate a:sat Saturation
     * a:satOff Saturation Offset a:satMod Saturation Modulation a:lum
     * Luminance a:lumOff Luminance Offset a:lumMod Luminance Modulation
     * a:red Red a:redOff Red Offset a:redMod Red Modulation a:green Green
     * a:greenOff Green Offset a:greenMod Green Modification a:blue Blue
     * a:blueOff Blue Offset a:blueMod Blue Modification a:gamma Gamma
     * a:invGamma Inverse Gamma
     */// Complement
    // Inverse
    // Gray
    // red,
    // green
    // blue
    // gamma
    // invGamma
    val ooxml: StringBuffer
        get() {
            val ooxml = StringBuffer()
            if (tint != 0)
                ooxml.append("<a:tint val=\"$tint\"/>")
            if (shade != 0)
                ooxml.append("<a:shade val=\"$shade\"/>")
            if (alpha != null) {
                if (alpha[0] != 0)
                    ooxml.append("<a:alpha val=\"" + alpha[0] + "\"/>")
                if (alpha[2] != 0)
                    ooxml.append("<a:alphaOff val=\"" + alpha[2] + "\"/>")
                if (alpha[1] != 0)
                    ooxml.append("<a:alphaMod val=\"" + alpha[1] + "\"/>")
            }
            if (hue != null) {
                if (hue[0] != 0)
                    ooxml.append("<a:hue val=\"" + hue[0] + "\"/>")
                if (hue[2] != 0)
                    ooxml.append("<a:hueOff val=\"" + hue[2] + "\"/>")
                if (hue[1] != 0)
                    ooxml.append("<a:hueMod val=\"" + hue[1] + "\"/>")
            }
            if (sat != null) {
                if (sat[0] != 0)
                    ooxml.append("<a:sat val=\"" + sat[0] + "\"/>")
                if (sat[2] != 0)
                    ooxml.append("<a:satOff val=\"" + sat[2] + "\"/>")
                if (sat[1] != 0)
                    ooxml.append("<a:satMod val=\"" + sat[1] + "\"/>")
            }
            if (lum != null) {
                if (lum[0] != 0)
                    ooxml.append("<a:lum val=\"" + lum[0] + "\"/>")
                if (lum[2] != 0)
                    ooxml.append("<a:lumOff val=\"" + lum[2] + "\"/>")
                if (lum[1] != 0)
                    ooxml.append("<a:lumMod val=\"" + lum[1] + "\"/>")
            }
            return ooxml
        }

    companion object {

        /**
         * parse color transform elements, common children of color-type elements: <br></br>
         * schemeColor, systemColor, hslColor, presetColor, sRgbColor, scRgbColor
         *
         * @param xpp
         * @param lastTag
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): ColorTransform {
            var lum: IntArray? = null
            var hue: IntArray? = null
            var sat: IntArray? = null
            val alpha: IntArray? = null
            var tint = 0
            var shade = 0
            try {
                val parentEl = lastTag.peek()

                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "lum") { // This element specifies the input
                            // color with its luminance
                            // modulated by the given
                            // percentage.
                            if (lum == null)
                                lum = IntArray(3)
                            lum[0] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "lumMod") { // This element specifies
                            // the input color with
                            // its luminance
                            // modulated by the
                            // given percentage.
                            if (lum == null)
                                lum = IntArray(3)
                            lum[1] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "lumOff") { // This element specifies
                            // the input color with
                            // its luminance
                            // shifted, but with its
                            // hue and saturation
                            // unchanged.
                            if (lum == null)
                                lum = IntArray(3)
                            lum[2] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "hue") { // This element specifies
                            // the input color with the
                            // specified hue, but with
                            // its saturation and
                            // luminance unchanged.
                            if (hue == null)
                                hue = IntArray(3)
                            hue[0] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "hueMod") {
                            if (hue == null)
                                hue = IntArray(3)
                            hue[1] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "hueOff") {
                            if (hue == null)
                                hue = IntArray(3)
                            hue[2] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "sat") { // This element specifies
                            // the input color with the
                            // specified saturation, but
                            // with its hue and
                            // luminance unchanged.
                            if (sat == null)
                                sat = IntArray(3)
                            sat[0] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "satMod") {
                            if (sat == null)
                                sat = IntArray(3)
                            sat[1] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "satOff") {
                            if (sat == null)
                                sat = IntArray(3)
                            sat[2] = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "shade") { // This element specifies
                            // a darker version of
                            // its input color
                            shade = Integer.valueOf(xpp.getAttributeValue(0))
                        } else if (tnm == "tint") {
                            tint = Integer.valueOf(xpp.getAttributeValue(0))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == parentEl) {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("ColorTransform.parseOOXML: $e")
            }

            return ColorTransform(lum, hue, sat, alpha, tint, shade)
        }
    }
}