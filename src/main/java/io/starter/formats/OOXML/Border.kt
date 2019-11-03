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

import java.util.HashMap

/**
 * border OOXML element
 *
 *
 * parent:  	styleSheet/borders element in styles.xml
 * children: 	SEQ: left, right, top, bottom, diagonal, vertical, horizontal
 */
class Border : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var borderElements: HashMap<String, BorderElement>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<border")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append(">")
            if (borderElements!!["left"] != null) ooxml.append(borderElements!!["left"].ooxml)
            if (borderElements!!["right"] != null) ooxml.append(borderElements!!["right"].ooxml)
            if (borderElements!!["top"] != null) ooxml.append(borderElements!!["top"].ooxml)
            if (borderElements!!["bottom"] != null) ooxml.append(borderElements!!["bottom"].ooxml)
            if (borderElements!!["diagonal"] != null) ooxml.append(borderElements!!["diagonal"].ooxml)
            if (borderElements!!["vertical"] != null) ooxml.append(borderElements!!["vertical"].ooxml)
            if (borderElements!!["horizontal"] != null) ooxml.append(borderElements!!["horizontal"].ooxml)

            ooxml.append("</border>")
            return ooxml.toString()
        }

    /**
     * returns an array representing the border sizes
     * <br></br>top, left, bottom, right, diag
     *
     * @return int[5] representing border sizes
     */
    val borderSizes: IntArray
        get() {
            val sizes = IntArray(5)
            if (borderElements!!["top"] != null) sizes[0] = borderElements!!["top"].borderSize
            if (borderElements!!["left"] != null) sizes[1] = borderElements!!["left"].borderSize
            if (borderElements!!["bottom"] != null) sizes[2] = borderElements!!["bottom"].borderSize
            if (borderElements!!["right"] != null) sizes[3] = borderElements!!["right"].borderSize
            if (borderElements!!["diagonal"] != null) sizes[4] = borderElements!!["diagonal"].borderSize
            return sizes
        }

    /**
     * returns an array representing the border styles
     * translated from OOXML String value to 2003-int value
     * <br></br>top, left, bottom, right, diag
     *
     * @return int[5]
     */
    val borderStyles: IntArray
        get() {
            val styles = IntArray(5)
            if (borderElements!!["top"] != null) styles[0] = borderElements!!["top"].borderStyle
            if (borderElements!!["left"] != null) styles[1] = borderElements!!["left"].borderStyle
            if (borderElements!!["bottom"] != null) styles[2] = borderElements!!["bottom"].borderStyle
            if (borderElements!!["right"] != null) styles[3] = borderElements!!["right"].borderStyle
            if (borderElements!!["diagonal"] != null) styles[4] = borderElements!!["diagonal"].borderStyle
            return styles
        }

    /**
     * returns an array representing the border colors as rgb string
     * <br></br>top, left, bottom, right, diag
     *
     * @return String[6]
     */
    val borderColors: Array<String>
        get() {
            try {
                val clrs = arrayOfNulls<String>(5)
                if (borderElements!!["top"] != null) clrs[0] = borderElements!!["top"].borderColor
                if (borderElements!!["left"] != null) clrs[1] = borderElements!!["left"].borderColor
                if (borderElements!!["bottom"] != null) clrs[2] = borderElements!!["bottom"].borderColor
                if (borderElements!!["right"] != null) clrs[3] = borderElements!!["right"].borderColor
                if (borderElements!!["diagonal"] != null) clrs[4] = borderElements!!["diagonal"].borderColor
                return clrs
            } catch (e: NullPointerException) {
                return arrayOfNulls(5)
            }

        }

    /**
     * returns an array representing the border colors as rgb string
     * <br></br>top, left, bottom, right, diag
     *
     * @return String[6]
     */
    val borderColorInts: IntArray
        get() {
            try {
                val clrs = IntArray(5)
                if (borderElements!!["top"] != null) clrs[0] = borderElements!!["top"].borderColorInt
                if (borderElements!!["left"] != null) clrs[1] = borderElements!!["left"].borderColorInt
                if (borderElements!!["bottom"] != null) clrs[2] = borderElements!!["bottom"].borderColorInt
                if (borderElements!!["right"] != null) clrs[3] = borderElements!!["right"].borderColorInt
                if (borderElements!!["diagonal"] != null) clrs[4] = borderElements!!["diagonal"].borderColorInt
                return clrs
            } catch (e: NullPointerException) {
                return IntArray(5)
            }

        }

    constructor() {}

    /**
     * @param styles int array {top, left, top, bottom, right, [diagonal]}
     * @param colors int array {top, left, top, bottom, right, [diagonal]}
     */
    constructor(attrs: HashMap<String, String>, borderElements: HashMap<String, BorderElement>) {
        this.attrs = attrs
        this.borderElements = borderElements
    }

    constructor(b: Border) {
        this.attrs = b.attrs
        this.borderElements = b.borderElements
    }

    /**
     * set borders
     *
     * @param bk
     * @param styles t, l, b, r
     * @param colors
     */
    constructor(bk: WorkBookHandle, styles: IntArray, colors: IntArray) {
        this.borderElements = HashMap()
        val borderElements = arrayOf("top", "left", "bottom", "right")
        for (i in 0..3) {
            if (styles[i] > 0) {
                val style = OOXMLConstants.borderStyle[styles[i]]
                this.borderElements!![borderElements[i]] = BorderElement(style, colors[i], borderElements[i], bk)
            }
        }
        // diagonal? vertical?  horizontal?
    }

    override fun cloneElement(): OOXMLElement {
        return Border(this)
    }

    override fun toString(): String {
        return if (borderElements != null) borderElements!!.toString() else "<none>"
    }

    companion object {

        private val serialVersionUID = 4340789910636828223L

        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): OOXMLElement {
            val attrs = HashMap<String, String>()
            val borderElements = HashMap<String, BorderElement>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "border") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "left") {
                            borderElements["left"] = BorderElement.parseOOXML(xpp, bk)
                        } else if (tnm == "right") {
                            borderElements["right"] = BorderElement.parseOOXML(xpp, bk)
                        } else if (tnm == "top") {
                            borderElements["top"] = BorderElement.parseOOXML(xpp, bk)
                        } else if (tnm == "bottom") {
                            borderElements["bottom"] = BorderElement.parseOOXML(xpp, bk)
                        } else if (tnm == "diagonal") {
                            borderElements["diagonal"] = BorderElement.parseOOXML(xpp, bk)
                        } else if (tnm == "vertical") {
                            borderElements["vertical"] = BorderElement.parseOOXML(xpp, bk)
                        } else if (tnm == "horizontal") {
                            borderElements["horizontal"] = BorderElement.parseOOXML(xpp, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "border") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("border.parseOOXML: $e")
            }

            return Border(attrs, borderElements)
        }

        /**
         * return an OOXML representation of this border based on this FormatHandle object
         */
        fun getOOXML(xf: Xf): String {
            val ooxml = StringBuffer()
            ooxml.append("<border>")
            val lineStyles = IntArray(5)
            lineStyles[0] = xf.leftBorderLineStyle.toInt()
            lineStyles[1] = xf.rightBorderLineStyle.toInt()
            lineStyles[2] = xf.topBorderLineStyle.toInt()
            lineStyles[3] = xf.bottomBorderLineStyle.toInt()
            lineStyles[4] = xf.diagBorderLineStyle.toInt()

            val colors = IntArray(5)
            colors[0] = xf.leftBorderColor
            colors[1] = xf.rightBorderColor.toInt()
            colors[2] = xf.topBorderColor
            colors[3] = xf.bottomBorderColor
            colors[4] = xf.diagBorderColor.toInt()

            if (lineStyles[0] > 0) {
                ooxml.append("<left")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[lineStyles[0]] + "\"")
                if (colors[0] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[0]]).substring(1) + "\"/></left>")
                else
                    ooxml.append("/>")
            }

            if (lineStyles[1] > 0) {
                ooxml.append("<right")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[lineStyles[1]] + "\"")
                if (colors[1] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[1]]).substring(1) + "\"/></right>")
                else
                    ooxml.append("/>")
            }

            if (lineStyles[2] > 0) {
                ooxml.append("<top")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[lineStyles[2]] + "\"")
                if (colors[2] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[2]]).substring(1) + "\"/></top>")
                else
                    ooxml.append("/>")
            }

            if (lineStyles[3] > 0) {
                ooxml.append("<bottom")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[lineStyles[3]] + "\"")
                if (colors[3] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[3]]).substring(1) + "\"/></bottom>")
                else
                    ooxml.append("/>")
            }

            if (lineStyles[4] > 0) {
                ooxml.append("<diagonal")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[lineStyles[4]] + "\"")
                if (colors[4] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[4]]).substring(1) + "\"/></diagonal>")
                else
                    ooxml.append("/>")
            }
            ooxml.append("</border>")
            return ooxml.toString()
        }

        /**
         * return the OOMXL to define a border based on the below specifications:
         *
         * @param styles int array {top, left, top, bottom, right, [diagonal]}
         * @param colors int array {top, left, top, bottom, right, [diagonal]}
         */
        fun getOOXML(styles: IntArray, colors: IntArray): String {
            val ooxml = StringBuffer()
            ooxml.append("<border>")
            // left
            if (styles[0] > 0) {
                ooxml.append("<left")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[styles[1]] + "\"")
                if (colors[0] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[0]]).substring(1) + "\"/></left>")
                else
                    ooxml.append("/>")
            }
            // right
            if (styles[1] > 0) {
                ooxml.append("<right")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[styles[3]] + "\"")
                if (colors[1] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[1]]).substring(1) + "\"/></right>")
                else
                    ooxml.append("/>")
            }
            // top
            if (styles[2] > 0) {
                ooxml.append("<top")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[styles[0]] + "\"")
                if (colors[2] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[2]]).substring(1) + "\"/></top>")
                else
                    ooxml.append("/>")
            }
            // bottom
            if (styles[3] > 0) {
                ooxml.append("<bottom")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[styles[2]] + "\"")
                if (colors[3] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[3]]).substring(1) + "\"/></top>")
                else
                    ooxml.append("/>")
            }
            // diagonal
            if (styles[4] > 0) {
                ooxml.append("<diagonal")
                ooxml.append(" style=\"" + OOXMLConstants.borderStyle[styles[4]] + "\"")
                if (colors[4] > 0)
                    ooxml.append("><color rgb=\"" + FormatHandle.colorToHexString(FormatHandle.COLORTABLE[colors[4]]).substring(1) + "\"/></diagonal>")
                else
                    ooxml.append("/>")
            }
            ooxml.append("</border>")
            return ooxml.toString()
        }
    }
}

/**
 * one of:
 * left, right, top, bottom, diagonal, vertical, horizontal
 *
 *
 * parent: 	border
 * children:  color
 */
internal class BorderElement : OOXMLElement {
    private var style: String? = null
    private var color: Color? = null
    private var borderElement: String? = null

    /**
     * @return the border size for this border element, translated from OOXML string to 2003-style int
     */
    // otherwise, interpret style --> size???
    // hair
    val borderSize: Int
        get() {
            val st = borderStyle
            if (st <= 4)
                return st + 1
            if (st == 7)
                return 1
            return if (st == 6 || st == 8 || st == 0xC) 3 else 2
        }

    /**
     * return the border style for this border element, translated from OOXML string to 2003-style int
     *
     * @return
     */
    val borderStyle: Int
        get() {
            if (style == null || style == "none")
                return -1
            else if (style == "thin")
                return 1
            else if (style == "medium")
                return 2
            else if (style == "dashed")
                return 3
            else if (style == "dotted")
                return 4
            else if (style == "thick")
                return 5
            else if (style == "double")
                return 6
            else if (style == "hair")
                return 7
            else if (style == "mediumDashed")
                return 8
            else if (style == "dashDot")
                return 9
            else if (style == "mediumDashDot")
                return 0xA
            else if (style == "dashDotDot")
                return 0xB
            else if (style == "mediumDashDotDot")
                return 0xC
            else if (style == "slantDashDot")
                return 0xD
            return -1
        }

    /**
     * return the rgb color string for this border element
     *
     * @return
     */
    val borderColor: String?
        get() = if (color != null) this.color!!.colorAsOOXMLRBG else null

    val borderColorInt: Int
        get() = if (color != null) this.color!!.colorInt else 0

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<" + borderElement!!)
            if (style != null) {
                ooxml.append(" style=\"$style\">")
                if (color != null) ooxml.append(color!!.ooxml)
                ooxml.append("</$borderElement>")
            } else
                ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(style: String, c: Color, borderElement: String) {
        this.style = style
        this.color = c
        this.borderElement = borderElement
    }

    constructor(b: BorderElement) {
        this.style = b.style
        this.color = b.color
        this.borderElement = b.borderElement
    }

    /**
     * constructor from string representation of a border element
     *
     * @param style         "thin", "thick" ...
     * @param val           rgb color value
     * @param borderElement "left", "right", "top", "bottom", "diagonal"
     */
    constructor(style: String?, `val`: String?, borderElement: String, bk: WorkBookHandle) {
        this.style = style
        if (style != null && `val` != null)
            this.color = Color("color", false, Color.COLORTYPERGB, `val`, 0.0, 0.toShort(), bk.workBook!!.theme)
        this.borderElement = borderElement
    }

    /**
     * create a new border element
     *
     * @param style         "thin", "thick" ...
     * @param val           color int
     * @param borderElement "left", "right", "top", "bottom", "diagonal"
     */
    constructor(style: String?, `val`: Int, borderElement: String, bk: WorkBookHandle) {
        this.style = style
        if (style != null && `val` != -1)
            this.color = Color("color", false, Color.COLORTYPEINDEXED, `val`.toString(), 0.0, 0.toShort(), bk.workBook!!.theme)
        this.borderElement = borderElement
    }

    override fun cloneElement(): OOXMLElement {
        return BorderElement(this)
    }

    override fun toString(): String {
        return (if (style != null) style else "<none>") +
                " c:" + if (color != null) color!!.toString() else "<none>"

    }

    companion object {

        private val serialVersionUID = -8040551653089261574L

        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): BorderElement {
            var style: String? = null
            var c: Color? = null
            var borderElement: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "color") {
                            c = Color.parseOOXML(xpp, 0.toShort(), bk) as Color
                        } else {    // one of the border elements
                            if (xpp.attributeCount > 0)
                            // style
                                style = xpp.getAttributeValue(0)
                            borderElement = tnm
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == borderElement) {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("borderElement.parseOOXML: $e")
            }

            return BorderElement(style, c, borderElement)
        }
    }
}

