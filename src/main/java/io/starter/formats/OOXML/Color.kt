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

/**
 * color (Data Bar Color)
 * One of the colors associated with the data bar or color scale.  Also font ...
 * Note: the auto attribute is not used in the context of data bars
 *
 *
 * NOTE: both color, fgColor and bgColor use same CT_COLOR schema:
 * <complexType name="CT_Color">
 * <attribute name="auto" type="xsd:boolean" use="optional"></attribute>
 * <attribute name="indexed" type="xsd:unsignedInt" use="optional"></attribute>
 * <attribute name="rgb" type="ST_UnsignedIntHex" use="optional"></attribute>
 * <attribute name="theme" type="xsd:unsignedInt" use="optional"></attribute>
 * <attribute name="tint" type="xsd:double" use="optional" default="0.0"></attribute>
</complexType> *
 *
 *
 * fgColor (Foreground Color)
 * Foreground color of the cell fill pattern. Cell fill patterns operate with two colors: a background color and a
 * foreground color. These combine together to make a patterned cell fill.
 *
 *
 *
 *
 * bgColor (Background Color)
 * Background color of the cell fill pattern. Cell fill patterns operate with two colors: a background color and a
 * foreground color. These combine together to make a patterned cell fill.
 */
class Color : OOXMLElement {
    private var auto = false
    /**
     * return the color type for this OOXML color
     * (indexed= 0, rgb= 1, theme= 2)
     *
     * @return
     */
    var colorType = -1
        private set
    private var colorval: String? = null        // value of colortype
    private var tint = 0.0
    private var element: String? = null
    private var colorint = -1        // parsed color (tint + value) translated to OpenXLS color int value
    /**
     * return the translated HTML color string for this OOXML color
     *
     * @return
     */
    var colorAsOOXMLRBG: String? = null
        private set    // parsed color (tint + value) translated to HTML color string
    private var theme: Theme? = null

    override// rgb
    // theme
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<" + this.element!!)
            if (this.auto)
                ooxml.append(" auto=\"1\"")
            else if (this.colorType == COLORTYPERGB)
                ooxml.append(" rgb=\"$colorAsOOXMLRBG\"")
            else if (this.colorType == COLORTYPEINDEXED)
                ooxml.append(" indexed=\"$colorval\"")
            else if (this.colorType == COLORTYPETHEME)
                ooxml.append(" theme=\"$colorval\"")
            if (tint != 0.0)
                ooxml.append(" tint=\"$tint\"")
            ooxml.append("/>")
            return ooxml.toString()
        }

    /**
     * return the translated color int for this OOXML color
     *
     * @return
     */
    /**
     * manually set color int
     * automatically defaults to indexed color, and looks up color in COLORTABLE
     *
     * @param clr
     */
    // indexed
    // reset other vars as well
    var colorInt: Int
        get() = this.colorint
        set(clr) {
            this.colorint = clr
            this.colorType = 0
            if (clr > -1)
                this.colorAsOOXMLRBG = FormatHandle.colorToHexString(FormatHandle.COLORTABLE[clr]).substring(1)
            else
                this.colorAsOOXMLRBG = null
            this.tint = 0.0
            this.auto = false
        }

    val colorAsColor: java.awt.Color
        get() = FormatHandle.HexStringToColor(this.colorAsOOXMLRBG!!)

    constructor() {}

    constructor(element: String, auto: Boolean, colortype: Int, colorval: String, tint: Double, type: Short, t: Theme) {
        this.element = element
        this.auto = auto
        this.colorType = colortype
        this.colorval = colorval
        this.tint = tint
        this.theme = t
        this.parseColor(type)
    }

    constructor(c: Color) {
        this.element = c.element
        this.auto = c.auto
        this.colorType = c.colorType
        this.colorval = c.colorval
        this.tint = c.tint
        this.colorint = c.colorint
        this.colorAsOOXMLRBG = c.colorAsOOXMLRBG
        this.theme = c.theme
    }


    /**
     * creates a new Color object based upon a java.awt.Color
     *
     * @param clr     Color objeect
     * @param element "color", "fgColor" or "bgColor"
     */
    constructor(c: java.awt.Color, element: String, t: Theme) {
        this.colorType = COLORTYPERGB
        this.colorval = "FF" + FormatHandle.colorToHexString(c).substring(1)
        this.element = element
        this.theme = t    // ok if it's null
        this.parseColor(0.toShort())
    }

    /**
     * creates a new Color object based upon a web-compliant Hex Color string
     *
     * @param clr
     * @param element "color", "fgColor" or "bgColor"
     */
    constructor(clr: String, element: String, t: Theme) {
        this.colorType = COLORTYPERGB
        if (clr.startsWith("#"))
            this.colorval = "FF" + clr.substring(1)
        else if (clr.length == 6)
            this.colorval = "FF$clr"
        else
            this.colorval = clr
        this.element = element
        this.theme = t    // ok if it's null
        this.parseColor(0.toShort())
    }

    override fun cloneElement(): OOXMLElement {
        return Color(this)
    }

    override fun toString(): String {
        var ret = ""
        if (colorType == COLORTYPERGB)
        // rgb
            ret = " rgb=" + colorAsOOXMLRBG!!
        else if (colorType == COLORTYPEINDEXED)
        // indexed
            ret = " indexed=" + colorval!!
        else if (colorType == COLORTYPETHEME)
        // theme
            ret = " theme=" + colorval!!
        if (tint != 0.0)
            ret = " tint=$tint"
        return ret
    }

    /**
     * sets the color via java.awt.Color
     *
     * @param c
     */
    fun setColor(c: java.awt.Color) {
        this.colorType = COLORTYPERGB
        this.colorval = "FF" + FormatHandle.colorToHexString(c).substring(1)
        this.parseColor(0.toShort())
        this.tint = 0.0
        this.auto = false
    }


    /**
     * sets the color via a web-compliant Hex Color String
     *
     * @param clr
     */
    fun setColor(clr: String) {
        this.colorType = COLORTYPERGB
        if (clr.startsWith("#"))
            this.colorval = "FF" + clr.substring(1)
        else if (clr.length == 6)
            this.colorval = "FF$clr"
        else
            this.colorval = clr
        this.parseColor(0.toShort())
    }

    /**
     * simple utility to parse correct colors (int + html string) from OOXML style color
     *
     * @param type 0= font, -1= fill
     */
    private fun parseColor(type: Short) {
        try {
            if (this.colorType == COLORTYPERGB) {        // rgb - color string
                this.colorint = FormatHandle.HexStringToColorInt(this.colorval!!, type) // find best match
                this.colorAsOOXMLRBG = this.colorval
            } else if (this.colorType == COLORTYPEINDEXED) {    // indexed (corresponds to either our color int or a custom set of indexed colors
                this.colorint = Integer.valueOf(this.colorval!!).toInt()
                if (this.colorint == 64 && type == FormatHandle.colorFONT)
                // means system foreground: Default foreground color. This is the window text color in the sheet display.
                    this.colorAsOOXMLRBG = FormatHandle.colorToHexString(FormatHandle.getColor(0))
                else
                    this.colorAsOOXMLRBG = FormatHandle.colorToHexString(FormatHandle.getColor(this.colorint))//					this.colorint= 0; // black
            } else if (this.colorType == COLORTYPETHEME) {    // theme
                val o = Color.parseThemeColor(this.colorval, this.tint, type, this.theme)
                this.colorint = o[0] as Int
                this.colorAsOOXMLRBG = o[1] as String
            }
        } catch (e: Exception) {
            Logger.logWarn("color.parseColor: $colorType:$colorval: $e")
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2546003092245407502L
        var COLORTYPEINDEXED = 0
        var COLORTYPERGB = 1
        var COLORTYPETHEME = 2

        /**
         * parses a color element
         * root= color, fgColor or bgColor
         *
         * @param xpp
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser, type: Short, bk: WorkBookHandle): OOXMLElement {
            var element: String? = null
            var auto = false
            var colortype = -1
            var colorval: String? = null
            var tint = 0.0
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "color" ||
                                tnm == "fgColor" ||
                                tnm == "bgColor") {        // get attributes
                            element = tnm    // save element name
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                val `val` = xpp.getAttributeValue(i)
                                if (n == "auto") {
                                    auto = true
                                } else if (n == "indexed") {
                                    colortype = Color.COLORTYPEINDEXED
                                    colorval = `val`
                                } else if (n == "rgb") {
                                    colortype = COLORTYPERGB
                                    colorval = `val`
                                } else if (n == "theme") {
                                    colortype = COLORTYPETHEME
                                    colorval = `val`
                                } else if (n == "tint") {
                                    tint = Double(`val`)
                                }
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == element) {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("color.parseOOXML: $e")
            }

            return Color(element, auto, colortype, colorval, tint, type, bk.workBook!!.theme)
        }

        fun getOOXML(element: String, colortype: Int, colorval: Int, colorstr: String, tint: Double): String {
            val ooxml = StringBuffer()
            ooxml.append("<$element")
            if (colortype == COLORTYPERGB)
            // rgb
                ooxml.append(" rgb=\"$colorstr\"")
            else if (colortype == COLORTYPEINDEXED)
            // indexed
                ooxml.append(" indexed=\"$colorval\"")
            else if (colortype == COLORTYPETHEME)
            // theme
                ooxml.append(" theme=\"$colorval\"")
            if (tint != 0.0)
                ooxml.append(" tint=\"$tint\"")
            ooxml.append("/>")
            return ooxml.toString()
        }

        /**
         * static version of parseColor --takes a color value of OOXML type type
         * <br></br>COLORTYPERBG, COLORTYPEINDEXED or COLORTYPETHEME
         * <br></br>and returns the
         *
         * @param val       String OOXML color value, value depends upon colortype
         * @param colortype int val's colortype, one of:  Color.COLORTYPERGB, Color.COLORTYPEINDEXED, Color.COLORTYPETHEME
         * @param type      0= font, 1= fill
         * @return String    HEX-style color string
         */
        fun parseColor(`val`: String, colortype: Int, type: Short, t: Theme): String? {
            val c = Color("", false, colortype, `val`, 0.0, type, t)
            return c.colorAsOOXMLRBG
        }

        /**
         * static version of parseColor --takes a color value of OOXML type type
         * <br></br>COLORTYPERBG, COLORTYPEINDEXED or COLORTYPETHEME
         * <br></br>and returns the indexed color int it represents
         *
         * @param val       String OOXML color value, value depends upon colortype
         * @param colortype int val's colortype, one of:  Color.COLORTYPERGB, Color.COLORTYPEINDEXED, Color.COLORTYPETHEME
         * @param type      0= font, 1= fill
         * @return int
         */
        fun parseColorInt(`val`: String, colortype: Int, type: Short, t: Theme): Int {
            val c = Color("", false, colortype, `val`, 0.0, type, t)
            //		c.parseColor(type);
            return c.colorint
        }

        /**
         * interprets theme colorval and tint
         * TODO: read in theme colors from theme1.xml
         *
         * @param colorval String theme colorval (see SchemeClr and Theme for details)
         * EITHER an Index into the <clrScheme> collection, referencing a particular <sysClr> or 	<srgbClr> value expressed in the Theme part.
         * OR the actual clrScheme string
         * @param tint     double tint to apply to the base color (0 for none)
         * @param type     short 0= font, 1= fill
         * @return
        </srgbClr></sysClr></clrScheme> */
        fun parseThemeColor(colorval: String?, tint: Double, type: Short, t: Theme?): Array<Any> {
            /**
             * TODO: read this in from THEME!!!
             */
            val themeColors = arrayOf(arrayOf("dk1", "000000"), arrayOf("lt1", "FFFFFF"), arrayOf("dk2", "1F497D"), arrayOf("lt2", "EEECE1"), arrayOf("accent1", "4F81BD"), arrayOf("accent1", "C0504D"), arrayOf("accent3", "9BBB59"), arrayOf("accent4", "8064A2"), arrayOf("accent5", "4BACC6"), arrayOf("accent6", "F79646"), arrayOf("hlink", "0000FF"), arrayOf("folHlink", "800080"))

            // TODO: read theme colors from file
            val o = arrayOfNulls<Any>(2)
            var i = 0
            var clr = ""
            try {
                i = Integer.valueOf(colorval!!).toInt()
                clr = t!!.genericThemeClrs[i]
            } catch (e: Exception) {        // OpenXLS-created xlsx files will not have theme entries
                i = 0
                while (i < themeColors.size) {
                    if (themeColors[i][0] == colorval) {
                        clr = themeColors[i][1]
                        break
                    }
                    i++
                }
            }

            if (Math.abs(tint) > .005) {
                val c = applyTint(tint, clr)
                o[0] = FormatHandle.getColorInt(c)
                o[1] = FormatHandle.colorToHexString(c).substring(1)    // avoid #
            } else {
                o[0] = FormatHandle.HexStringToColorInt("#$clr", type)
                o[1] = "#$clr"
            }
            return o
        }


        /**
         * If tint is supplied, then it is applied to the RGB value of the color to determine the final
         * color applied.
         * The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and
         * 1.0 means 100% lighten. Also, 0.0 means no change.
         * In loading the RGB value, it is converted to HLS where HLS values are (0..HLSMAX), where
         * HLSMAX is currently 255.
         */
        private fun applyTint(tint: Double, clr: String): java.awt.Color {
            val HLSMAX = 255
            val r = Integer.parseInt(clr.substring(0, 2), 16)
            val g = Integer.parseInt(clr.substring(2, 4), 16)
            val b = Integer.parseInt(clr.substring(4, 6), 16)

            val hsl = HSLColor()
            hsl.initHSLbyRGB(r, g, b)
            var l = hsl.luminence.toDouble()
            if (tint < 0)
            //darken
                l = l * (1 + tint)
            else
            // lighten
                l = l * (1 - tint) + (HLSMAX - HLSMAX * (1.0 - tint))
            l = Math.round(l).toDouble()
            hsl.initRGBbyHSL(hsl.hue, hsl.saturation, l.toInt())
            return java.awt.Color(hsl.red, hsl.green, hsl.blue)

        }
    }
}


/* HSL stands for hue, saturation, lightness
 * Both HSL and HSV describe colors as points in a cylinder
 * whose central axis ranges from black at the bottom to white at the top
 * with neutral colors between them,
 * where angle around the axis corresponds to hue,
 * distance from the axis corresponds to saturation,
 * and distance along the axis corresponds to lightness, value, or brightness.*/

/*
 * Conversion from RGB to HSL

    Let r, g, b E [0,1] be the red, green, and blue coordinates, respectively, of a color in RGB space.
    Let max be the greatest of r, g, and b, and min the least.

    hue angle h:

    h= 0 													if max=min (for grays)
     = 60 deg x (g-b)/(max-min) + 360 deg) mod 360 deg		if max=r
     = 60 deg x (b-r)/(max-min) + 120 deg)					if max=g
     = 60 deg x (r-g)/(max-min) + 240 deg) 					if max=b

    lightness l= (max-min)/2

    Saturation s

    s= 0													if max=min
     = (max-min)/2l											if l <= 1/2
     = (max-min)/(2-2l)										if l > 1/2

    The value of h is generally normalized to lie between 0 and 360, and h = 0 is used when max = min (that is, for grays)
    though the hue has no geometric meaning there, where the saturation s is zero.
    Similarly, the choice of 0 as the value for s when l is equal to 0 or 1 is arbitrary.

    Conversion from HSL to RGB

    Given a color defined by (h, s, l) values in HSL space, with h in the range [0, 360), indicating the angle,
    in degrees of the hue, and with s and l in the range [0, 1], representing the saturation and lightness, respectively,
    a corresponding (r, g, b) triplet in RGB space, with r, g, and b also in range [0, 1], and corresponding to red, green, and blue,
    respectively, can be computed as follows:

    First, if s = 0, then the resulting color is achromatic, or gray. In this special case, r, g, and b all equal l.
    Note that the value of h is ignored, and may be undefined in this situation.

    The following procedure can be used, even when s is zero:

    q= l x (1+s) 					if l < 1/2
     = l + s - (l x s) 				if l >= 1/2

    p= 2 x l - q

    hk= h / 360 (h normalized in the range 0-1)

    tr= hk + 1/3

    tg= hk

    tb= hk - 1/3

    if tc

 */
internal class HSLColor {

    private var pHue: Int = 0
    private var pSat: Int = 0
    private var pLum: Int = 0
    private var pRed: Int = 0
    private var pGreen: Int = 0
    private var pBlue: Int = 0


    // --

    var hue: Int
        get() = pHue
        set(iToValue) {
            var iToValue = iToValue
            while (iToValue < 0) {
                iToValue = HSLMAX + iToValue
            }
            while (iToValue > HSLMAX) {
                iToValue = iToValue - HSLMAX
            }

            initRGBbyHSL(iToValue, pSat, pLum)
        }

    // --

    var saturation: Int
        get() = pSat
        set(iToValue) {
            var iToValue = iToValue
            if (iToValue < 0) {
                iToValue = 0
            } else if (iToValue > HSLMAX) {
                iToValue = HSLMAX
            }

            initRGBbyHSL(pHue, iToValue, pLum)
        }

    // --

    var luminence: Int
        get() = pLum
        set(iToValue) {
            var iToValue = iToValue
            if (iToValue < 0) {
                iToValue = 0
            } else if (iToValue > HSLMAX) {
                iToValue = HSLMAX
            }

            initRGBbyHSL(pHue, pSat, iToValue)
        }

    // --

    var red: Int
        get() = pRed
        set(iNewValue) = initHSLbyRGB(iNewValue, pGreen, pBlue)

    // --

    var green: Int
        get() = pGreen
        set(iNewValue) = initHSLbyRGB(pRed, iNewValue, pBlue)

    // --

    var blue: Int
        get() = pBlue
        set(iNewValue) = initHSLbyRGB(pRed, pGreen, iNewValue)

    fun initHSLbyRGB(R: Int, G: Int, B: Int) {
        // sets Hue, Sat, Lum
        val cMax: Int
        val cMin: Int
        val RDelta: Int
        val GDelta: Int
        val BDelta: Int
        val cMinus: Int
        val cPlus: Int

        pRed = R
        pGreen = G
        pBlue = B

        //Set Max & MinColor Values
        cMax = iMax(iMax(R, G), B)
        cMin = iMin(iMin(R, G), B)

        cMinus = cMax - cMin
        cPlus = cMax + cMin

        // Calculate luminescence (lightness)
        pLum = (cPlus * HSLMAX + RGBMAX) / (2 * RGBMAX)

        if (cMax == cMin) {
            // greyscale
            pSat = 0
            pHue = UNDEFINED
        } else {
            // Calculate color saturation
            if (pLum <= HSLMAX / 2) {
                pSat = ((cMinus * HSLMAX + 0.5) / cPlus).toInt()
            } else {
                pSat = ((cMinus * HSLMAX + 0.5) / (2 * RGBMAX - cPlus)).toInt()
            }

            //Calculate hue
            RDelta = (((cMax - R) * (HSLMAX / 6) + 0.5) / cMinus).toInt()
            GDelta = (((cMax - G) * (HSLMAX / 6) + 0.5) / cMinus).toInt()
            BDelta = (((cMax - B) * (HSLMAX / 6) + 0.5) / cMinus).toInt()

            if (cMax == R) {
                pHue = BDelta - GDelta
            } else if (cMax == G) {
                pHue = HSLMAX / 3 + RDelta - BDelta
            } else if (cMax == B) {
                pHue = 2 * HSLMAX / 3 + GDelta - RDelta
            }

            if (pHue < 0) {
                pHue = pHue + HSLMAX
            }
        }
    }

    fun initRGBbyHSL(H: Int, S: Int, L: Int) {
        val Magic1: Int
        val Magic2: Int

        pHue = H
        pLum = L
        pSat = S

        if (S == 0) { //Greyscale
            pRed = L * RGBMAX / HSLMAX //luminescence: set to range
            pGreen = pRed
            pBlue = pRed
        } else {
            if (L <= HSLMAX / 2) {
                Magic2 = (L * (HSLMAX + S) + HSLMAX / 2) / HSLMAX
            } else {
                Magic2 = L + S - (L * S + HSLMAX / 2) / HSLMAX
            }
            Magic1 = 2 * L - Magic2

            //get R, G, B; change units from HSLMAX range to RGBMAX range
            pRed = (hueToRGB(Magic1, Magic2, H + HSLMAX / 3) * RGBMAX + HSLMAX / 2) / HSLMAX
            if (pRed > RGBMAX) {
                pRed = RGBMAX
            }

            pGreen = (hueToRGB(Magic1, Magic2, H) * RGBMAX + HSLMAX / 2) / HSLMAX
            if (pGreen > RGBMAX) {
                pGreen = RGBMAX
            }

            pBlue = (hueToRGB(Magic1, Magic2, H - HSLMAX / 3) * RGBMAX + HSLMAX / 2) / HSLMAX
            if (pBlue > RGBMAX) {
                pBlue = RGBMAX
            }
        }
    }

    private fun hueToRGB(mag1: Int, mag2: Int, Hue: Int): Int {
        var Hue = Hue
        // check the range
        if (Hue < 0) {
            Hue = Hue + HSLMAX
        } else if (Hue > HSLMAX) {
            Hue = Hue - HSLMAX
        }

        if (Hue < HSLMAX / 6)
            return mag1 + ((mag2 - mag1) * Hue + HSLMAX / 12) / (HSLMAX / 6)

        if (Hue < HSLMAX / 2)
            return mag2

        return if (Hue < HSLMAX * 2 / 3) mag1 + ((mag2 - mag1) * (HSLMAX * 2 / 3 - Hue) + HSLMAX / 12) / (HSLMAX / 6) else mag1

    }

    private fun iMax(a: Int, b: Int): Int {
        return if (a > b)
            a
        else
            b
    }

    private fun iMin(a: Int, b: Int): Int {
        return if (a < b)
            a
        else
            b
    }


    fun greyscale() {
        initRGBbyHSL(UNDEFINED, 0, pLum)
    }

    // --

    fun reverseColor() {
        hue = pHue + HSLMAX / 2
    }

    // --

    fun reverseLight() {
        luminence = HSLMAX - pLum
    }

    // --

    fun brighten(fPercent: Float) {
        var L: Int

        if (fPercent == 0f) {
            return
        }

        L = (pLum * fPercent).toInt()
        if (L < 0) L = 0
        if (L > HSLMAX) L = HSLMAX

        luminence = L
    }

    // --
    // --

    fun blend(R: Int, G: Int, B: Int, fPercent: Float) {
        if (fPercent >= 1) {
            initHSLbyRGB(R, G, B)
            return
        }
        if (fPercent <= 0)
            return

        val newR = (R * fPercent + pRed * (1.0 - fPercent)).toInt()
        val newG = (G * fPercent + pGreen * (1.0 - fPercent)).toInt()
        val newB = (B * fPercent + pBlue * (1.0 - fPercent)).toInt()

        initHSLbyRGB(newR, newG, newB)
    }

    companion object {
        private val HSLMAX = 255
        private val RGBMAX = 255
        private val UNDEFINED = 170
    }
}

/*
	Color.getHSBColor(hue, saturation, brightness
	 * The hue parameter is a decimal number between 0.0 and 1.0 which indicates the hue of the color. You'll have to experiment with the hue number to find out what color it represents.
	 * The saturation is a decimal number between 0.0 and 1.0 which indicates how deep the color should be. Supplying a "1" will make the color as deep as possible, and to the other extreme,
	 * supplying a "0," will take all the color out of the mixture and make it a shade of gray.
	 * The brightness is also a decimal number between 0.0 and 1.0 which obviously indicates how bright the color should be.
	 * A 1 will make the color as light as possible and a 0 will make it very dark.
	 */

/*
			int maxC = Math.max(Math.max(r,g),b);
			int minC = Math.min(Math.min(r,g),b);
			int l = (((maxC + minC)*HLSMAX) + RGBMAX)/(2*RGBMAX);
			int delta= maxC - minC;
			int sum= maxC + minC;
			if (delta != 0) {
					if (l <= (HLSMAX/2))
						s = ( (delta*HLSMAX) + (sum/2) ) / sum;
					else
						s = ( (delta*HLSMAX) + ((2*RGBMAX - sum)/2) ) / (2*RGBMAX - sum);
			
					if (r == maxC)
						h =                ((g - b) * HLSMAX) / (6 * delta);
					else if (g == maxC)
						h = (  HLSMAX/3) + ((b - r) * HLSMAX) / (6 * delta);
					else if (b == maxC)
						h = (2*HLSMAX/3) + ((r - g) * HLSMAX) / (6 * delta);

					if (h < 0)
						h += HLSMAX;
					else if (h >= HLSMAX)
						h -= HLSMAX;
			} else {
					h = 0;
					s = 0;
			}
			

	*/
	
	