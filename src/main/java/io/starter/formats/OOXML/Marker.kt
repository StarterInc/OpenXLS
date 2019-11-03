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
 * class holds OOXML to define markers for a chart (radar, scatter or line)
 *
 *
 * parent:  dPt, Ser ...
 * children: symbol, size, spPr
 * NOTE: child elements symbol and size have only 1 attribute and no children, and so are treated as strings
 */
class Marker : OOXMLElement {
    private var sp: SpPr? = null
    private var size: String? = null
    private var symbol: String? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<c:marker>")
            if (this.symbol != null)
                ooxml.append("<c:symbol val=\"" + this.symbol + "\"/>")
            if (this.size != null)
                ooxml.append("<c:size val=\"" + this.size + "\"/>")
            if (this.sp != null) ooxml.append(sp!!.ooxml)
            ooxml.append("</c:marker>")
            return ooxml.toString()
        }

    constructor(s: String, sz: String, sp: SpPr) {
        this.symbol = s
        this.size = sz
        this.sp = sp
    }

    constructor(m: Marker) {
        this.symbol = m.symbol
        this.size = m.size
        this.sp = m.sp
    }

    override fun cloneElement(): OOXMLElement {
        return Marker(this)
    }

    companion object {

        private val serialVersionUID = -5070227633357072878L

        /**
         * parse marker OOXML element
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return marker object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var sp: SpPr? = null
            var size: String? = null        // size element:  val is only attribute + no children
            var symbol: String? = null  // symbol element:  val is only attribute + no children
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "symbol") {
                            symbol = xpp.getAttributeValue(0)
                        } else if (tnm == "size") {
                            size = xpp.getAttributeValue(0)
                        } else if (tnm == "spPr") {
                            lastTag.push(tnm)
                            sp = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //sp.setNS("c");
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "marker") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("marker.parseOOXML: $e")
            }

            return Marker(symbol, size, sp)
        }
    }
}
