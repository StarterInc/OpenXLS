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

internal class Xfrm : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var o: Off? = null
    private var e: Ext? = null
    private var ns: String? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<" + this.ns + ":xfrm")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append(">")
            if (o != null) ooxml.append(o!!.ooxml)
            if (e != null) ooxml.append(e!!.ooxml)
            ooxml.append("</$ns:xfrm>")
            return ooxml.toString()
        }

    constructor() {
        this.ns = "xdr"    // set default
    }

    constructor(attrs: HashMap<String, String>, o: Off, e: Ext, ns: String) {
        this.attrs = attrs
        this.o = o
        this.e = e
        this.ns = ns
    }

    constructor(x: Xfrm) {
        this.attrs = x.attrs
        this.o = x.o
        this.e = x.e
        this.ns = x.ns
    }

    /**
     * set the namespace for xfrm element
     * xdr (graphicFrame) or a(spPr)
     *
     * @param ns
     */
    fun setNS(ns: String) {
        this.ns = ns
    }

    override fun cloneElement(): OOXMLElement {
        return Xfrm(this)
    }

    companion object {

        private val serialVersionUID = 5383438744617393878L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            val attrs = HashMap<String, String>()
            var o: Off? = null
            var e: Ext? = null
            var ns: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "xfrm") {        // get attributes
                            ns = xpp.prefix
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "off") {
                            lastTag.push(tnm)
                            o = Off.parseOOXML(xpp, lastTag)
                            //o.setNS("a");
                        } else if (tnm == "ext") {
                            lastTag.push(tnm)
                            e = Ext.parseOOXML(xpp, lastTag) as Ext
                            //e.setNS("a");
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "xfrm") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (ex: Exception) {
                Logger.logErr("xfrm.parseOOXML: $ex")
            }

            return Xfrm(attrs, o, e, ns)
        }
    }

}

internal class Off : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var ns = ""

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<" + this.ns + ":off")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor() {
        attrs = HashMap()
        attrs!!["x"] = "0"
        attrs!!["y"] = "0"
    }

    constructor(attrs: HashMap<String, String>, ns: String) {
        this.attrs = attrs
        this.ns = ns
    }

    constructor(o: Off) {
        this.attrs = o.attrs
        this.ns = o.ns
    }

    fun setNS(ns: String) {
        this.ns = ns
    }

    override fun cloneElement(): OOXMLElement {
        return Off(this)
    }

    companion object {

        private val serialVersionUID = -7624630398053353694L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): Off {
            val attrs = HashMap<String, String>()
            var ns: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "off") {        // get attributes
                            ns = xpp.prefix
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "off") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("off.parseOOXML: $e")
            }

            return Off(attrs, ns)
        }
    }
}
	
