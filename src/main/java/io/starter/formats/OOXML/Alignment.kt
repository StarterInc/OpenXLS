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

import java.util.HashMap


/**
 * alignment (Alignment) OOMXL element
 *
 *
 * Formatting information pertaining to text alignment in cells. There are a variety of choices for how text is
 * aligned both horizontally and vertically, as well as indentation settings, and so on.
 *
 *
 * parent:		(styles.xml) xf, dxf
 * children: none
 */
class Alignment : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<alignment")
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

    constructor(a: Alignment) {
        this.attrs = a.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Alignment(this)
    }

    /**
     * @param type horizontal vertical
     * @return
     */
    fun getAlignment(type: String): String? {
        // attributes:
        // horizontal: center, centerContinuous, fill, general, justify, left, right, distributed
        // indent: int value
        // justifyLastLine: bool
        // readingOrder:	0=Context Dependent, 1=Left-to-Right, 2=Right-to-Left
        // relativeIndent: #
        // shrinkToFit: bool
        // textRotation:	degrees from 0-180
        // vertical: bottom, centered, distributed, justify, top
        // wrapText	- true/false
        return if (attrs != null) {
            attrs!![type]
        } else null
    }

    companion object {

        private val serialVersionUID = 995367747930839216L

        fun parseOOXML(xpp: XmlPullParser): OOXMLElement {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "alignment") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "alignment") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("alignment.parseOOXML: $e")
            }

            return Alignment(attrs)
        }
    }
}

