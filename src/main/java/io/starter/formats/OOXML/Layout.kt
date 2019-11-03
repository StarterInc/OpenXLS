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

import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * chart layout element
 *
 *
 *
 *
 * parent:  plotarea, title ...
 * children:  manualLayout
 */
class Layout : OOXMLElement {
    private var ml: ManualLayout? = null

    /**
     * return manual layout coords
     * TODO: interpret xMode, yMode, hMode, wMode
     *
     * @return
     */
    val coords: FloatArray
        get() {
            val coords = FloatArray(4)
            for (i in 0..3)
                if (ml!!.offs[i] != null)
                    coords[i] = Float(ml!!.offs[i])
            return coords
        }

    /**
     * generate ooxml to define a layout
     *
     * @return
     */
    override val ooxml: String
        get() {
            val looxml = StringBuffer()
            if (ml != null) {
                looxml.append("<c:layout>")
                looxml.append(ml!!.ooxml)
                looxml.append("</c:layout>")
            } else
                looxml.append("<c:layout/>")
            return looxml.toString()
        }


    private constructor(ml: ManualLayout) {
        this.ml = ml
    }

    private constructor(l: Layout) {
        this.ml = l.ml
    }

    /**
     * create a new plot area Layout/manual layout element
     * <br></br>note that the layout is calculated via "edge" i.e the w and h are the bottom and right edges
     *
     * @param target "inner" or "outer" <br></br>inner specifies the plot area size does not include tick marks and axis labels<br></br>outer does
     * @param offs   x, y, w, h as a fraction of the width or height of the actual chart
     */
    constructor(target: String, offs: DoubleArray) {
        val modes = arrayOf<String>("edge", "edge", null, null)
        val soffs = arrayOfNulls<String>(4)
        for (i in 0..3) {
            if (offs[i] > 0)
                soffs[i] = offs[i].toString()
        }
        this.ml = ManualLayout(target, modes, soffs)
    }

    override fun cloneElement(): OOXMLElement {
        return Layout(this)
    }

    companion object {

        private val serialVersionUID = -6547994902298821138L

        /**
         * parse title OOXML element title
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spPr object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            var ml: ManualLayout? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "manualLayout") {
                            lastTag.push(tnm)
                            ml = ManualLayout.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "layout") {
                            lastTag.pop()    // pop layout tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("layout.parseOOXML: $e")
            }

            return Layout(ml)
        }

        /**
         * static version for 2003-style charts, takes offsets and creates a layout element
         *
         * @param offs int offsets of chart
         * @return
         */
        fun getOOXML(offs: DoubleArray): String {
            val looxml = StringBuffer()
            looxml.append("<c:layout>")
            looxml.append("<c:manualLayout>")
            looxml.append("<c:layoutTarget val=\"inner\"/>")
            looxml.append("<c:xMode val=\"edge\"/>")
            looxml.append("<c:yMode val=\"edge\"/>")
            looxml.append("<c:x val=\"" + offs[0] + "\"/>")
            looxml.append("<c:y val=\"" + offs[1] + "\"/>")
            looxml.append("<c:w val=\"" + offs[2] + "\"/>")
            looxml.append("<c:h val=\"" + offs[3] + "\"/>")
            looxml.append("</c:manualLayout>")
            looxml.append("</c:layout>")
            return looxml.toString()
        }
    }
}

/**
 * manualLayout
 * specifies exact position of chart
 *
 *
 * parent:  layout
 * children:  h, hMode, layoutTarget, w, wMode, x, xMode, y, yMode
 */
internal class ManualLayout : OOXMLElement {
    var modes: Array<String> // xMode, yMode, wMode, hMode
    var offs: Array<String>    // x, y, w, h;
    var target: String? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<c:manualLayout>")
            if (target != null) ooxml.append("<c:layoutTarget val=\"" + this.target + "\"/>")
            if (modes[0] != null) ooxml.append("<c:xMode val=\"" + this.modes[0] + "\"/>")
            if (this.modes[1] != null) ooxml.append("<c:yMode val=\"" + this.modes[1] + "\"/>")
            if (this.modes[2] != null) ooxml.append("<c:wMode val=\"" + this.modes[2] + "\"/>")
            if (this.modes[3] != null) ooxml.append("<c:hMode val=\"" + this.modes[3] + "\"/>")
            if (this.offs[0] != null) ooxml.append("<c:x val=\"" + this.offs[0] + "\"/>")
            if (this.offs[1] != null) ooxml.append("<c:y val=\"" + this.offs[1] + "\"/>")
            if (this.offs[2] != null) ooxml.append("<c:w val=\"" + this.offs[2] + "\"/>")
            if (this.offs[3] != null) ooxml.append("<c:h val=\"" + this.offs[3] + "\"/>")
            ooxml.append("</c:manualLayout>")
            return ooxml.toString()
        }

    constructor(target: String, modes: Array<String>, offs: Array<String>) {
        this.modes = modes
        this.target = target
        this.offs = offs
    }

    constructor(ml: ManualLayout) {
        this.modes = ml.modes
        this.target = ml.target
        this.offs = ml.offs
    }

    override fun cloneElement(): OOXMLElement {
        return ManualLayout(this)
    }

    companion object {

        private val serialVersionUID = 6460833211809500902L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): ManualLayout {
            val modes = arrayOfNulls<String>(4) // xMode, yMode, wMode, hMode
            val offs = arrayOfNulls<String>(4)    // x, y, w, h;
            var target: String? = null

            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "manualLayout") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                if (n == "ATTR") {
                                    //sz= Integer.valueOf(xpp.getAttributeValue(i)).intValue();
                                }
                            }
                        } else if (tnm == "layoutTarget") {
                            target = xpp.getAttributeValue(0)
                        } else if (tnm == "xMode") {
                            modes[0] = xpp.getAttributeValue(0)
                        } else if (tnm == "yMode") {
                            modes[1] = xpp.getAttributeValue(0)
                        } else if (tnm == "wMode") {
                            modes[2] = xpp.getAttributeValue(0)
                        } else if (tnm == "hMode") {
                            modes[3] = xpp.getAttributeValue(0)
                        } else if (tnm == "x") {
                            offs[0] = xpp.getAttributeValue(0)
                        } else if (tnm == "y") {
                            offs[1] = xpp.getAttributeValue(0)
                        } else if (tnm == "w") {
                            offs[2] = xpp.getAttributeValue(0)
                        } else if (tnm == "h") {
                            offs[3] = xpp.getAttributeValue(0)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "manualLayout") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("manualLayout.parseOOXML: $e")
            }

            return ManualLayout(target, modes, offs)
        }
    }
}

