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
import java.util.Stack

/**
 * gd (Shape Guide)
 *
 *
 * This element specifies the presence of a shape guide that will be used to govern the geometry of the specified
 * shape. A shape guide consists of a formula and a name that the result of the formula is assigned to. Recognized
 * formulas are listed with the fmla attribute documentation for this element.
 *
 *
 * parents:  avLst, gdLst
 * children: none
 */
class Gd : OOXMLElement {
    private var attrs: HashMap<*, *>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:gd")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next() as String
                val `val` = attrs!!.get(key) as String
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<*, *>) {
        this.attrs = attrs
    }

    constructor(g: Gd) {
        this.attrs = g.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Gd(this)
    }

    companion object {

        private val serialVersionUID = -633176234309521998L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "gd") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "gd") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("gd.parseOOXML: $e")
            }

            return Gd(attrs)
        }
    }
}

