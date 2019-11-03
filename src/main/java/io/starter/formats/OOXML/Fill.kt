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
import io.starter.formats.XLS.Xf
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList
import java.util.HashMap

/**
 * fill OOXML element
 *
 *
 * fill (Fill) This element specifies fill formatting
 *
 *
 * parent: styleSheet/fills element in styles.xml, dxf->fills children: REQ
 * CHOICE OF: patternFill, gradientFill
 */
class Fill : OOXMLElement {
    private var patternFill: PatternFill? = null
    private var gradientFill: GradientFill? = null
    private var theme: Theme? = null

    override val ooxml: String
        get() = getOOXML(false)

    val fillPattern: String?
        get() = if (patternFill != null) this.patternFill!!.fillPattern else null

    /**
     * return the fill pattern in 2003 int representation
     *
     * @return
     */
    val fillPatternInt: Int
        get() = if (patternFill != null) {
            this.patternFill!!.fillPatternInt
        } else -1

    /**
     * returns true if the background pattern is solid
     *
     * @return
     */
    val isBackgroundSolid: Boolean
        get() = if (patternFill != null) this.patternFill!!.fillPattern!!.equals("solid", ignoreCase = true) else false

    constructor(p: PatternFill, g: GradientFill, t: Theme) {
        this.patternFill = p
        this.gradientFill = g
        this.theme = t
    }

    constructor(f: Fill) {
        if (f.patternFill != null)
            this.patternFill = f.patternFill!!.cloneElement() as PatternFill
        if (f.gradientFill != null)
            this.gradientFill = f.gradientFill!!.cloneElement() as GradientFill
        this.theme = f.theme
    }

    /**
     * create a new Fill from external vals
     *
     * @param fs String pattern type
     * @param fg int color index
     * @param bg int color index
     */
    constructor(fs: String, fg: Int, bg: Int, t: Theme) {
        this.patternFill = PatternFill(fs, fg, bg)
        this.theme = t
    }

    /**
     * create a new Fill from external vals
     *
     * @param i             XLS indexed pattern
     * @param fg            int color index
     * @param fgColorCustom
     * @param bg            int color index
     */
    constructor(pattern: Int, fg: Int, fgColorCustom: String, bg: Int,
                bgColorCustom: String, t: Theme) {

        this.patternFill = PatternFill(PatternFill.translateIndexedFillPattern(pattern),
                fg, fgColorCustom, bg, bgColorCustom)
        this.theme = t
    }

    /**
     * dxfs apparently have different pattern fill syntax -- UNDOCUMENTED ****
     *
     * @param isDxf if this is an Dxf-generated fill, solid fills are handled
     * differently than regular fills
     * @return
     */
    fun getOOXML(isDxf: Boolean): String {
        val ooxml = StringBuffer()
        ooxml.append("<fill>")
        if (patternFill != null)
            ooxml.append(patternFill!!.getOOXML(isDxf))
        if (gradientFill != null)
            ooxml.append(gradientFill!!.ooxml)
        ooxml.append("</fill>")
        return ooxml.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return Fill(this)
    }

    /**
     * return the foreground color of this fill, if any
     *
     * @return
     */
    fun getFgColorAsRGB(t: Theme): String? {
        return if (patternFill != null) this.patternFill!!.getFgColorAsRGB(t) else null
    }

    /**
     * return the fg color in indexed (int) representation
     *
     * @return
     */
    fun getFgColorAsInt(t: Theme): Int {
        return if (patternFill != null) this.patternFill!!.getFgColorAsInt(t) else 0
// default= black
    }

    /**
     * return the bg color of this fill, if any
     *
     * @return
     */
    fun getBgColorAsRGB(t: Theme): String? {
        return if (patternFill != null) this.patternFill!!.getBgColorAsRGB(t) else null
    }


    /**
     * return the bg color in indexed (int) representation
     *
     * @return
     */
    fun getBgColorAsInt(t: Theme): Int {
        return if (patternFill != null) this.patternFill!!.getBgColorAsInt(t) else -1
    }

    /**
     * sets the foreground fill color via color int
     *
     * @param t
     */
    fun setFgColor(t: Int) {
        if (patternFill != null) {
            patternFill!!.setFgColor(t)
            return
        }
    }

    /**
     * sets the foreground fill color to a color string and Excel-2003-mapped
     * color int
     *
     * @param t           Excel-2003-mapped color int for the hex color string
     * @param colorString hex color string
     */
    fun setFgColor(t: Int, colorString: String) {
        if (patternFill != null)
            patternFill!!.setFgColor(t, colorString)
        else
            this.patternFill = PatternFill("none", t, colorString, -1, null)
    }

    /**
     * sets the bg fill color via color int
     *
     * @param t
     */
    fun setBgColor(t: Int) {
        if (patternFill != null) {
            patternFill!!.setBgColor(t)
            return
        }
    }

    /**
     * sets the foreground fill color to a color string and Excel-2003-mapped
     * color int
     *
     * @param t           Excel-2003-mapped color int for the hex color string
     * @param colorString hex color string
     */
    fun setBgColor(t: Int, colorString: String) {
        if (patternFill != null)
            patternFill!!.setBgColor(t, colorString)
        else
            this.patternFill = PatternFill("none", -1, null, t, colorString)
    }

    /**
     * sets the fill pattern
     */
    fun setFillPattern(t: Int) {
        if (patternFill != null)
            this.patternFill!!.setFillPattern(t)
    }

    companion object {

        private val serialVersionUID = -4510508531435037641L

        fun parseOOXML(xpp: XmlPullParser, isDxf: Boolean, bk: WorkBookHandle): OOXMLElement {
            var p: PatternFill? = null
            var g: GradientFill? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "patternFill") {
                            p = PatternFill.parseOOXML(xpp, isDxf, bk)

                        } else if (tnm == "gradientFill") {
                            g = GradientFill.parseOOXML(xpp, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fill") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("fill.parseOOXML: $e")
            }

            return Fill(p, g, bk.workBook!!.theme)
        }

        /**
         * OOXML values are stored with an intiall FF in their value,
         *
         *
         * this method assures that values returned are 6 digits + # for web usage
         *
         * @param rgbcolor
         */
        fun transformToWebRGBColor(rgbcolor: String): String {
            if (rgbcolor.indexOf("#") == 0) return rgbcolor
            return if (rgbcolor.indexOf("FF") == 0 && rgbcolor.length == 8) {
                "#" + rgbcolor.substring(2)
            } else "#$rgbcolor"
        }

        /**
         * OOXML values are stored with an intiall FF in their value,
         *
         *
         * this method assures that values returned are 8 digits with
         * ff appended to rgb value for ooxml usage
         *
         * @param rgbcolor
         */
        fun transformToOOXMLRGBColor(rgbcolor: String): String {
            if (rgbcolor.indexOf("#") == 0) {
                return "FF" + rgbcolor.substring(1)
            }
            return if (rgbcolor.indexOf("FF") == 0 && rgbcolor.length == 8) {
                rgbcolor
            } else "FF$rgbcolor"
        }

        /**
         * returns the OOXML specifying the fill based on this FormatHandle object
         */
        fun getOOXML(xf: Xf): String {
            if (xf.getFill() != null)
                return xf.getFill()!!.ooxml

            // otherwise, create fill from 2003-style xf
            val ooxml = StringBuffer()
            ooxml.append("<fill>")
            try {
                ooxml.append("<patternFill patternType=\""
                        + OOXMLConstants.patternFill[xf.fillPattern] + "\">")
            } catch (e: IndexOutOfBoundsException) {
                ooxml.append("<patternFill>") // apparently there are less patterns
                // in xlsx? some other way of
                // storage
            }

            val fg = xf.foregroundColor.toInt()
            if (fg > -1 && fg != 64)
                ooxml.append("<fgColor rgb=\""
                        + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[fg])
                        .substring(1) + "\"/>")
            val bg = xf.backgroundColor.toInt()
            if (bg > -1 && bg != 64)
                ooxml.append("<bgColor rgb=\""
                        + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[bg])
                        .substring(1) + "\"/>")
            ooxml.append("</patternFill>")
            ooxml.append("\r\n")
            ooxml.append("</fill>")
            ooxml.append("\r\n")
            return ooxml.toString()
        }

        /**
         * returns the OOXML specifying the fill based on fill pattern fs,
         * foreground color fg, background color bg
         */
        fun getOOXML(fs: Int, fg: Int, bg: Int): String {
            val ooxml = StringBuffer()
            ooxml.append("<fill>")
            try {
                ooxml.append("<patternFill patternType=\""
                        + OOXMLConstants.patternFill[fs] + "\">")
            } catch (e: ArrayIndexOutOfBoundsException) {
                ooxml.append("<patternFill>")
            }

            if (fg > -1 && fg != 64) {
                ooxml.append("<fgColor rgb=\""
                        + FormatHandle
                        .colorToHexString(FormatHandle.COLORTABLE[fg])
                        .substring(1) + "\"/>")
                ooxml.append("\r\n")
            }
            if (bg > -1 && bg != 64) {
                ooxml.append("<bgColor rgb=\""
                        + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[bg])
                        .substring(1) + "\"/>")
                ooxml.append("\r\n")
            }
            ooxml.append("</patternFill>")
            ooxml.append("\r\n")
            ooxml.append("</fill>")
            ooxml.append("\r\n")
            return ooxml.toString()
        }
    }
}

/**
 * patternFill (Pattern) This element is used to specify cell fill information
 * for pattern and solid color cell fills. For solid cell fills (no pattern),
 * fgColor is used. For cell fills with patterns specified, then the cell fill
 * color is specified by the bgColor element.
 *
 *
 * parent: fill children: SEQ: fgColor, bgColor
 */
internal class PatternFill : OOXMLElement {
    /**
     * return the pattern type of this Fill in String representation
     *
     * @return
     */
    /**
     * sets the pattern type of this Fill <br></br>
     * One of:  * "none",  * "solid",  * "mediumGray",  * "darkGray",  *
     * "lightGray",  * "darkHorizontal",  * "darkVertical",  * "darkDown",
     *  * "darkUp",  * "darkGrid",  * "darkTrellis",  * "lightHorizontal",
     *  * "lightVertical",  * "lightDown",  * "lightUp",  * "lightGrid",  *
     * "lightTrellis",  * "gray125",  * "gray0625",
     *
     * @param s
     */
    var fillPattern: String? = null
    private var fgColor: FgColor? = null
    private var bgColor: BgColor? = null
    private var theme: Theme? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<patternFill")
            ooxml.append(" patternType=\"$fillPattern\">")
            if (fgColor != null)
                ooxml.append(fgColor!!.ooxml)
            if (bgColor != null)
                ooxml.append(bgColor!!.ooxml)
            ooxml.append("</patternFill>")
            return ooxml.toString()
        }

    /**
     * return the pattern type of this fill pattern in 2003 int representation <br></br>
     * BIG NOTE: Apparently doc is wrong in that if patternType is missing it
     * does NOT == none; instead, it is a SOLID FILL with BG==fill color <br></br>
     * THIS METHOD returns -1 in those cases
     *
     * @return int pattern fill integer
     * @see PatternFill.setFillPattern
     */
    // a missing entry *should*==none or 0, but
    // apparently is a distinct value in and of
    // itself
    // none --SHOULD BE
    // NONE BUT IS
    // 1==SOLID????
    // none
    val fillPatternInt: Int
        get() {
            if (fillPattern == null)
                return OOXMLConstants.patternFill.size + 1
            for (i in OOXMLConstants.patternFill.indices) {
                if (OOXMLConstants.patternFill[i] == fillPattern)
                    return i
            }
            return -1
        }

    constructor(patternType: String, fg: FgColor, bg: BgColor, t: Theme) {
        this.fillPattern = patternType
        this.fgColor = fg
        this.bgColor = bg
        this.theme = t
    }

    constructor(p: PatternFill) {
        this.fillPattern = p.fillPattern
        if (p.fgColor != null)
            this.fgColor = p.fgColor!!.cloneElement() as FgColor
        if (p.bgColor != null)
            this.bgColor = p.bgColor!!.cloneElement() as BgColor
        this.theme = p.theme
    }

    /**
     * create a new pattern fill from external vals
     *
     * @param patternType String OOXML pattern type
     * @param fg          int color index
     * @param bg          int color index
     */
    constructor(patternType: String, fg: Int, bg: Int) {
        this.fillPattern = patternType
        if (fg > -1 && fg != 64) {
            val attrs = HashMap<String, String>()
            attrs["rgb"] = "FF" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[fg]).substring(1)
            this.fgColor = FgColor(attrs)
        }
        if (bg > -1 && bg != 65) {
            val attrs = HashMap<String, String>()
            attrs["rgb"] = "FF" + FormatHandle.colorToHexString(
                    FormatHandle.COLORTABLE[bg]).substring(1)
            this.bgColor = BgColor(attrs)
        }
    }

    constructor(patternType: String, fg: Int, fgCustom: String?, bg: Int,
                bgCustom: String?) {
        this.fillPattern = patternType
        if (fg > 0 || fgCustom != null) { // 64= default fg color
            val attrs = HashMap<String, String>()
            if (fgCustom == null)
                attrs["indexed"] = fg.toString()
            else
                attrs["rgb"] = Fill.transformToOOXMLRGBColor(fgCustom)
            this.fgColor = FgColor(attrs)
        }
        if (bg > -1 || bgCustom != null) { // 65= default bg color
            val attrs = HashMap<String, String>()
            if (bgCustom == null)
                attrs["indexed"] = bg.toString()
            else
                attrs["rgb"] = Fill.transformToOOXMLRGBColor(bgCustom)
            this.bgColor = BgColor(attrs)
        }
    }

    /**
     * apparently Fill OOXML from Dxf has differnt syntax - UNDOCUMENTED
     *
     * @param isDxf if this is an Dxf-generated fill, solid fills are handled
     * differently than regular fills
     * @return
     */
    fun getOOXML(isDxf: Boolean): String {
        val ooxml = StringBuffer()
        ooxml.append("<patternFill")
        if (!isDxf) {
            ooxml.append(" patternType=\"$fillPattern\">")
            if (fgColor != null)
                ooxml.append(fgColor!!.ooxml)
            if (bgColor != null)
                ooxml.append(bgColor!!.ooxml)
        } else {
            if (fillPattern == "solid") { // dxf needs "none" or "soild" to
                // have bg color set, not fg
                // color as is normal
                ooxml.append(">")
                if (fgColor != null) { // shoudln't!
                    val tempbg = BgColor(fgColor!!.attrs)
                    ooxml.append(tempbg.ooxml)
                }
            } else {
                ooxml.append(" patternType=\"$fillPattern\">")
                if (fgColor != null)
                    ooxml.append(fgColor!!.ooxml)
                if (bgColor != null)
                    ooxml.append(bgColor!!.ooxml)
            }
        }
        ooxml.append("</patternFill>")
        return ooxml.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return PatternFill(this)
    }

    override fun toString(): String {
        return ((if (fillPattern != null) fillPattern else "<none>") + " fg:"
                + (if (fgColor != null) fgColor!!.toString() else "<none>") + " bg:"
                + if (bgColor != null) bgColor!!.toString() else "<none>")
    }

    /**
     * sets the pattern type of this fill via a 2003-style pattern int <br></br>
     * One of:  * FLSNULL 0x00 No fill pattern  * FLSSOLID 0x01 Solid  *
     * FLSMEDGRAY 0x02 50% gray  * FLSDKGRAY 0x03 75% gray  * FLSLTGRAY 0x04
     * 25% gray  * FLSDKHOR 0x05 Horizontal stripe  * FLSDKVER 0x06 Vertical
     * stripe  * FLSDKDOWN 0x07 Reverse diagonal stripe  * FLSDKUP 0x08
     * Diagonal stripe  * FLSDKGRID 0x09 Diagonal crosshatch  * FLSDKTRELLIS
     * 0x0A Thick Diagonal crosshatch  * FLSLTHOR 0x0B Thin horizontal stripe
     *  * FLSLTVER 0x0C Thin vertical stripe  * FLSLTDOWN 0x0D Thin reverse
     * diagonal stripe  * FLSLTUP 0x0E Thin diagonal stripe  * FLSLTGRID 0x0F
     * Thin horizontal crosshatch  * FLSLTTRELLIS 0x10 Thin diagonal crosshatch
     *  * FLSGRAY125 0x11 12.5% gray  * FLSGRAY0625 0x12 6.25% gray <br></br>
     * NOTE: There is a "Special Code" that indicates a missing patternType,
     * which has a significant meaning in OOXML
     *
     * @param t
     */
    fun setFillPattern(t: Int) {
        this.fillPattern = translateIndexedFillPattern(t)
    }

    /**
     * return the foreground color of this fill as an RGB string
     *
     * @return
     */
    fun getFgColorAsRGB(t: Theme): String? {
        return if (fgColor != null) Fill.transformToWebRGBColor(fgColor!!.getColorAsRGB(t)!!) else null
    }

    /**
     * return the foreground color of this fill as indexed color int
     *
     * @return
     */
    fun getFgColorAsInt(t: Theme): Int {
        if (fgColor != null)
            return fgColor!!.getColorAsInt(t)
        return if ("solid" == fillPattern) 0 else -1
    }

    /**
     * sets the foreground color of this pattern fill via color int
     *
     * @param t
     */
    fun setFgColor(t: Int) {
        if (fgColor != null)
            fgColor!!.setColor(t)
    }

    /**
     * sets the foreground fill color to a color string and Excel-2003-mapped
     * color int
     *
     * @param t           Excel-2003-mapped color int for the hex color string
     * @param colorString hex color string
     */
    fun setFgColor(t: Int, colorString: String?) {
        if (t > 0 || colorString != null) { // 64= default fg color
            val attrs = HashMap<String, String>()
            if (colorString == null)
                attrs["indexed"] = t.toString()
            else
                attrs["rgb"] = colorString
            this.fgColor = FgColor(attrs)
        }
    }

    /**
     * return the background color of this fill as an RGB string
     *
     * @return
     */
    fun getBgColorAsRGB(t: Theme): String? {
        return if (bgColor != null) Fill.transformToWebRGBColor(bgColor!!.getColorAsRGB(t)!!) else null
    }


    /**
     * return the background color of this fill as an indexed color int
     *
     * @return
     */
    fun getBgColorAsInt(t: Theme): Int {
        return if (bgColor != null) bgColor!!.getColorAsInt(t) else -1
// default=white
    }

    /**
     * sets the background color of this pattern fill via color int
     *
     * @param t
     */
    fun setBgColor(t: Int) {
        if (bgColor != null)
            bgColor!!.setColor(t)
    }

    /**
     * sets the background fill color to a color string and Excel-2003-mapped
     * color int
     *
     * @param t           Excel-2003-mapped color int for the hex color string
     * @param colorString hex color string
     */
    fun setBgColor(t: Int, colorString: String?) {
        if (t > 0 || colorString != null) {    // 65= default bg color
            val attrs = HashMap<String, String>()
            if (colorString == null)
                attrs["indexed"] = t.toString()
            else
                attrs["rgb"] = Fill.transformToOOXMLRGBColor(colorString)
            this.bgColor = BgColor(attrs)
        }
    }

    companion object {

        private val serialVersionUID = -4399355217499895956L

        fun parseOOXML(xpp: XmlPullParser, isDxf: Boolean, bk: WorkBookHandle): PatternFill {
            var patternType: String? = null // "none"; // default when missing -- so sez
            // the doc but doesn't appear what Excel
            // does
            /**
             * APPARENLTY patternFills in dxfs are DIFFERENT AND NOT FOLLOWING THE
             * DOCUMENTATION on regular patternFills
             *
             * APPARENTLY patternType="none" and missing patternType ARE NOT THE
             * SAME: if Missing, APPARENTLY means to fill with BG color ("solid"
             * means to fill with FG color)
             */
            var fg: FgColor? = null
            var bg: BgColor? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "patternFill") { // get attributes
                            if (xpp.attributeCount > 0)
                                patternType = xpp.getAttributeValue(0)
                        } else if (tnm == "fgColor") {
                            fg = FgColor.parseOOXML(xpp)
                        } else if (tnm == "bgColor") {
                            bg = BgColor.parseOOXML(xpp)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "patternFill") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("patternFill.parseOOXML: $e")
            }

            if (isDxf) {
                if (patternType == null)
                // null apparently does NOT mean none for
                // Dxf's
                    patternType = "solid" // see Dxf and Cf handling: null means
                // solid fill with bg color as cell
                // background
                if (patternType == "solid") {
                    if (bg != null) {// shouldn't!
                        fg = FgColor(bg.attrs) // so 2003-v can properly
                        // set solid cell
                        // pattern color
                    }
                    bg = BgColor(64)
                }
            }
            return PatternFill(patternType, fg, bg, bk.workBook!!.theme)
        }

        fun translateIndexedFillPattern(pattern: Int): String? {
            var newPattern: String? = null
            if (pattern == OOXMLConstants.patternFill.size + 1)
            // special code
                newPattern = null
            else if (pattern >= 0 && pattern < OOXMLConstants.patternFill.size)
                newPattern = OOXMLConstants.patternFill[pattern]

            return newPattern
        }
    }

}

/**
 * gradientFill (Gradient) This element defines a gradient-style cell fill.
 * Gradient cell fills can use one or two colors as the end points of color
 * interpolation.
 *
 *
 * parent: fill children: stop (0 or more) attributes: many
 */
internal class GradientFill : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var stops: ArrayList<Stop>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<gradientFill")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (stops != null) {
                for (j in stops!!.indices)
                    ooxml.append(stops!![j].ooxml)
            }
            ooxml.append("</gradientFill>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>, stops: ArrayList<Stop>) {
        this.attrs = attrs
        this.stops = stops
    }

    constructor(g: GradientFill) {
        this.attrs = g.attrs
        this.stops = g.stops
    }

    override fun cloneElement(): OOXMLElement {
        return GradientFill(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3633230059631047503L

        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): GradientFill {
            val attrs = HashMap<String, String>()
            var stops: ArrayList<Stop>? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "gradientFill") { // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "stop") {
                            if (stops == null)
                                stops = ArrayList()
                            stops.add(Stop.parseOOXML(xpp, bk))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "gradientFill") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("gradientFill.parseOOXML: $e")
            }

            return GradientFill(attrs, stops)
        }
    }
}

/**
 * fgColor (Foreground Color) Foreground color of the cell fill pattern. Cell
 * fill patterns operate with two colors: a background color and a foreground
 * color. These combine together to make a patterned cell fill
 *
 *
 * parent: patternFill chilren: none
 */
internal class FgColor : OOXMLElement {
    var attrs: HashMap<String, String>? = null
        private set

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<fgColor")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    protected constructor(f: FgColor) {
        this.attrs = f.attrs!!.clone() as HashMap<String, String>
    }

    protected constructor(c: Int) {
        attrs = HashMap()
        attrs!!["indexed"] = c.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return FgColor(this)
    }

    /**
     * return the Html string representation of this foreground color
     *
     *
     * Note that this will return an HTML correct value, such as #000000.
     *
     *
     * Values set in ooxml need to look like FF000000
     *
     * @return
     */
    fun getColorAsRGB(t: Theme): String? {
        var `val`: String? = attrs!!["rgb"]
        if (`val` != null)
            return `val`
        `val` = attrs!!["indexed"]
        if (`val` != null) {
            return if (Integer.parseInt(`val`) == 64) null else Color.parseColor(`val`, Color.COLORTYPEINDEXED, FormatHandle.colorFOREGROUND, t) // return "#000000"; //null;
        }
        `val` = attrs!!["theme"]
        if (`val` != null)
            return Color.parseColor(`val`, Color.COLORTYPETHEME, FormatHandle.colorFOREGROUND, t)
        `val` = attrs!!["auto"]
        return if (`val` != null) "#000000" else null
    }

    /**
     * sets the foreground color to the indexed color integer
     *
     * @param c
     */
    fun setColor(c: Int) {
        attrs!!.clear()
        attrs!!["indexed"] = c.toString()
    }

    /**
     * returns the fg color as an indexed color int
     */
    fun getColorAsInt(t: Theme): Int {
        var `val`: String? = attrs!!["auto"]
        if (`val` != null)
            return 0
        `val` = attrs!!["rgb"]
        if (`val` != null)
            return Color.parseColorInt(`val`, Color.COLORTYPERGB, FormatHandle.colorFOREGROUND, t)
        `val` = attrs!!["indexed"]
        if (`val` != null)
            return Integer.valueOf(`val`)
        `val` = attrs!!["theme"]
        return if (`val` != null) Color.parseColorInt(`val`, Color.COLORTYPETHEME, FormatHandle.colorFOREGROUND, t) else -1
    }

    override fun toString(): String {
        if (attrs != null) {
            val s = attrs!!.toString()
            return s.substring(1, s.length - 1)
        }
        return "none"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1274598491373019241L

        fun parseOOXML(xpp: XmlPullParser): FgColor {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "fgColor") { // get attributes
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fgColor") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("fgColor.parseOOXML: $e")
            }

            return FgColor(attrs)
        }
    }
}

/**
 * bgColor (Background Color) Background color of the cell fill pattern. Cell
 * fill patterns operate with two colors: a background color and a foreground
 * color. These combine together to make a patterned cell fill.
 *
 *
 * parent: patternFill children: none
 */
internal class BgColor : OOXMLElement {
    var attrs: HashMap<String, String>? = null
        private set

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<bgColor")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    protected constructor(f: BgColor) {
        this.attrs = f.attrs!!.clone() as HashMap<String, String>
    }

    constructor(c: Int) {
        attrs = HashMap()
        attrs!!["indexed"] = c.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return BgColor(this)
    }

    /**
     * return the Html string representation of this background color
     *
     * @return
     */
    fun getColorAsRGB(t: Theme): String? {
        var `val`: String? = attrs!!["rgb"]
        if (`val` != null)
            return `val`
        `val` = attrs!!["indexed"]
        if (`val` != null) {
            return if (Integer.parseInt(`val`) == 65) null else Color.parseColor(`val`, Color.COLORTYPEINDEXED, FormatHandle.colorBACKGROUND, t) // return "#FFFFFF"; //return null;
        }
        `val` = attrs!!["theme"]
        return if (`val` != null) Color.parseColor(`val`, Color.COLORTYPETHEME, FormatHandle.colorBACKGROUND, t) else null
    }


    /**
     * sets the background color to the indexed color integer
     */
    fun setColor(c: Int) {
        attrs!!.clear()
        attrs!!["indexed"] = c.toString()
    }

    /**
     * returns the bg color as an indexed color int
     */
    fun getColorAsInt(t: Theme): Int {
        var `val`: String? = attrs!!["rgb"]
        if (`val` != null)
            return Color.parseColorInt(`val`, Color.COLORTYPERGB, FormatHandle.colorBACKGROUND, t)
        `val` = attrs!!["indexed"]
        if (`val` != null)
            return Integer.valueOf(`val`)
        `val` = attrs!!["theme"]
        return if (`val` != null) Color.parseColorInt(`val`, Color.COLORTYPETHEME, FormatHandle.colorBACKGROUND, t) else 64
// default
    }

    override fun toString(): String {
        if (attrs != null) {
            val s = attrs!!.toString()
            return s.substring(1, s.length - 1)
        }
        return "none"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 43028503491956217L

        fun parseOOXML(xpp: XmlPullParser): BgColor {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "bgColor") { // get attributes
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "bgColor") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("bgColor.parseOOXML: $e")
            }

            return BgColor(attrs)
        }
    }
}

/**
 * stop (Gradient Stop) One of a sequence of two or more gradient stops,
 * constituting this gradient fill.
 *
 *
 * parent: gradientFill children: color REQ
 */
internal class Stop : OOXMLElement {
    private var position: String? = null
    private var c: Color? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<stop")
            ooxml.append(" position=\"" + this.position + "\"")
            ooxml.append(">")
            if (c != null)
                ooxml.append(c!!.ooxml)
            ooxml.append("</stop>")
            return ooxml.toString()
        }

    constructor(position: String, c: Color) {
        this.position = position
        this.c = c
    }

    constructor(s: Stop) {
        this.position = s.position
        this.c = s.c
    }

    override fun cloneElement(): OOXMLElement {
        return Stop(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -9215564484103992694L

        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): Stop {
            var position: String? = null
            var c: Color? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "stop") {
                            position = xpp.getAttributeValue(0) // position=
                            // REQUIRED
                        } else if (tnm == "color") {
                            c = Color.parseOOXML(xpp, (-1).toShort(), bk)
                                    .cloneElement() as Color
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "stop") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("stop.parseOOXML: $e")
            }

            return Stop(position, c)
        }
    }


}

