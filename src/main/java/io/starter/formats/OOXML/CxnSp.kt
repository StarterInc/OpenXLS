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

import java.util.HashMap
import java.util.Stack

/**
 * parents:  absoluteAnchor, oneCellAnchor, twoCellAnchor
 * children: nvCxnSpPr, spPr, style
 */
//TODO: finish nvCxnSpPr.cNvCxnSpPr element
class CxnSp : OOXMLElement {
    private var attrs = HashMap<String, String>()
    private var nvc: NvCxnSpPr? = null
    private var spPr: SpPr? = null
    private var style: Style? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:cxnSp")
            val i = attrs.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs[key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (nvc != null) ooxml.append(nvc!!.ooxml)
            if (spPr != null) ooxml.append(spPr!!.ooxml)
            if (style != null) ooxml.append(style!!.ooxml)
            ooxml.append("</xdr:cxnSp>")
            return ooxml.toString()
        }

    /**
     * get cNvPr name attribute
     *
     * @return
     */
    /**
     * set cNvPr name attribute
     *
     * @param name
     */
    var name: String?
        get() = if (nvc != null) nvc!!.name else null
        set(name) {
            if (nvc != null)
                nvc!!.name = name
        }

    /**
     * get cNvPr descr attribute
     *
     * @return
     */
    /**
     * set cNvPr description attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var descr: String?
        get() = if (nvc != null) nvc!!.descr else null
        set(descr) {
            if (nvc != null)
                nvc!!.descr = descr
        }

    /**
     * get Macro attribute
     */
    /**
     * set Macro attribute
     *
     * @param macro
     */
    var macro: String?
        get() = if (attrs["macro"] != null) attrs["macro"] else null
        set(macro) {
            attrs["macro"] = macro
        }

    /**
     * return the cNvPr id for this element
     *
     * @return
     */
    /**
     * set the cNvPr id for this element
     *
     * @param id
     */
    var id: Int
        get() = if (nvc != null) nvc!!.id else -1
        set(id) {
            if (nvc != null)
                nvc!!.id = id
        }

    /**
     * return the id for the embedded picture (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed attribute for this blip (the id for the embedded picture)
     *
     * @param embed
     */
    var embed: String?
        get() = if (spPr != null) spPr!!.embed else null
        set(embed) {
            if (spPr != null) spPr!!.embed = embed
        }

    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (spPr != null) spPr!!.link else null
        set(link) {
            if (spPr != null) spPr!!.link = link
        }

    constructor(attrs: HashMap<String, String>, nvc: NvCxnSpPr, sp: SpPr, s: Style) {
        this.attrs = attrs
        this.nvc = nvc
        this.spPr = sp
        this.style = s
    }

    constructor(c: CxnSp) {
        this.attrs = c.attrs
        this.nvc = c.nvc
        this.spPr = c.spPr
        this.style = c.style
    }

    override fun cloneElement(): OOXMLElement {
        return CxnSp(this)
    }

    companion object {

        private val serialVersionUID = -8492664135843926551L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            val attrs = HashMap<String, String>()
            var nvc: NvCxnSpPr? = null
            var sp: SpPr? = null
            var s: Style? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cxnSp") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "nvCxnSpPr") {
                            lastTag.push(tnm)
                            nvc = NvCxnSpPr.parseOOXML(xpp, lastTag)
                        } else if (tnm == "spPr") {
                            lastTag.push(tnm)
                            sp = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //	        			 sp.setNS("xdr");
                        } else if (tnm == "style") {
                            lastTag.push(tnm)
                            s = Style.parseOOXML(xpp, lastTag, bk) as Style
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cxnSp") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cxnSp.parseOOXML: $e")
            }

            return CxnSp(attrs, nvc, sp, s)
        }
    }
}

/**
 * nvCxnSpPr (Non-Visual Properties for a Connection Shape)
 * This element specifies all non-visual properties for a connection shape. This element is a container for the non15
 * visual identification properties, shape properties and application properties that are to be associated with a
 * DrawingML Reference Material - DrawingML - SpreadsheetML Drawing connection shape.
 * This allows for additional information that does not affect 1 the appearance of the connection
 * shape to be stored.
 *
 *
 * parent: cxnSp
 * children:  cNvPr, cNvCxnSpPr
 */
// TODO: finish cNvCxnSpPr
internal class NvCxnSpPr : OOXMLElement {
    private var cpr: CNvPr? = null

    override// TODO: finihs cNvCxnSpPr
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:nvCxnSpPr>")
            if (cpr != null) ooxml.append(cpr!!.ooxml)
            ooxml.append("<xdr:cNvCxnSpPr/>")
            ooxml.append("</xdr:nvCxnSpPr>")
            return ooxml.toString()
        }

    /**
     * get cNvPr name attribute
     *
     * @return
     */
    /**
     * set cNvPr name attribute
     *
     * @param name
     */
    var name: String?
        get() = if (cpr != null) cpr!!.name else null
        set(name) {
            if (cpr != null)
                cpr!!.name = name
        }

    /**
     * get cNvPr descr attribute
     *
     * @return
     */
    /**
     * set cNvPr description attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var descr: String?
        get() = if (cpr != null) cpr!!.descr else null
        set(descr) {
            if (cpr != null)
                cpr!!.descr = descr
        }

    /**
     * return the cNvPr id for this element
     *
     * @return
     */
    /**
     * set the cNvPr id for this element
     *
     * @param id
     */
    var id: Int
        get() = if (cpr != null) cpr!!.id else -1
        set(id) {
            if (cpr != null)
                cpr!!.id = id
        }
    //	private cNvCxnSpPr sppr= null;

    constructor(cpr: CNvPr/*, cNvCxnSpPr sppr*/) {
        this.cpr = cpr
        //this.sppr= sppr;
    }

    constructor(n: NvCxnSpPr) {
        this.cpr = n.cpr
        //this.sppr= n.sppr;
    }

    override fun cloneElement(): OOXMLElement {
        return NvCxnSpPr(this)
    }

    companion object {

        private val serialVersionUID = -4808617992996239153L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): NvCxnSpPr {
            var cpr: CNvPr? = null
            //    	cNvCxnSpPr sppr= null;
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvPr") {
                            lastTag.push(tnm)
                            cpr = CNvPr.parseOOXML(xpp, lastTag) as CNvPr
                            /*		            } else if (tnm.equals("cNvCxnSpPr")) {
	        			 lastTag.push(tnm);
	        			 sppr= (cNvCxnSpPr) cNvCxnSpPr.parseOOXML(xpp, lastTag).clone();
*/
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "nvCxnSpPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("nvCxnSpPr.parseOOXML: $e")
            }

            return NvCxnSpPr(cpr/*, sppr*/)
        }
    }
}


