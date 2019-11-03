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
 * pic (Picture)
 * This element specifies the existence of a picture object within the spreadsheet.
 *
 *
 * parent:  absoluteAnchor, grpSp, oneCellAnchor, twoCellAnchor
 * children: nvPicPr, blipFill, spPr, style
 */
//TODO: handle nvPicPr.cNvPicPr.picLocks element
class Pic : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var nvPicPr: NvPicPr? = null
    private var blipFill: BlipFill? = null
    /**
     * utility to return the shape properties element for this picture
     * should be depreciated when OOXML is completely distinct from BIFF8
     *
     * @return
     */
    /**
     * utility to return the shape properties element for this picture
     * should be depreciated when OOXML is completely distinct from BIFF8
     *
     * @return
     */
    var sppr: SpPr? = null
    private var style: Style? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:pic")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append(">")
            if (nvPicPr != null) ooxml.append(nvPicPr!!.ooxml)
            if (blipFill != null) ooxml.append(blipFill!!.ooxml)
            if (sppr != null) ooxml.append(sppr!!.ooxml)
            if (style != null) ooxml.append(style!!.ooxml)
            ooxml.append("</xdr:pic>")
            return ooxml.toString()
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
        get() = if (blipFill != null) blipFill!!.embed else null
        set(embed) {
            if (blipFill != null)
                blipFill!!.embed = embed
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
        get() = if (blipFill != null) blipFill!!.link else null
        set(link) {
            if (blipFill != null)
                blipFill!!.link = link
        }

    /**
     * return the name of this shape, if any
     *
     * @return
     */
    /**
     * set the name of this shape, if any
     */
    var name: String?
        get() = if (nvPicPr != null) nvPicPr!!.name else null
        set(name) {
            if (nvPicPr != null)
                nvPicPr!!.name = name
        }

    /**
     * get cNvPr descr attribute
     *
     * @return
     */
    /**
     * set cNvPr desc attribute
     *
     * @param name
     */
    var descr: String?
        get() = if (nvPicPr != null) nvPicPr!!.descr else null
        set(descr) {
            if (nvPicPr != null)
                nvPicPr!!.descr = descr
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
        get() {
            if (nvPicPr == null)
                nvPicPr = NvPicPr()
            return nvPicPr!!.id
        }
        set(id) {
            if (nvPicPr != null)
                nvPicPr!!.id = id
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
        get() = if (attrs!!["macro"] != null) attrs!!["macro"] else null
        set(macro) {
            attrs!!["macro"] = macro
        }

    constructor() {        // set common defaults
        nvPicPr = NvPicPr()
        blipFill = BlipFill()
        sppr = SpPr("xdr")
        sppr!!.setNS("xdr")
        attrs = null
    }

    constructor(attrs: HashMap<String, String>, nv: NvPicPr, bf: BlipFill, sp: SpPr, s: Style) {
        this.attrs = attrs
        this.nvPicPr = nv
        this.blipFill = bf
        this.sppr = sp
        this.style = s
    }

    constructor(p: Pic) {
        this.attrs = p.attrs
        this.nvPicPr = p.nvPicPr
        this.blipFill = p.blipFill
        this.sppr = p.sppr
        this.style = p.style
    }

    override fun cloneElement(): OOXMLElement {
        return Pic(this)
    }

    /**
     * add a line to this image
     *
     * @param w   int line width in
     * @param clr String color html string
     * @return
     */
    fun setLine(w: Int, clr: String) {
        if (sppr != null) {
            sppr!!.setLine(w, clr)
        }
    }

    companion object {

        private val serialVersionUID = -4929177274389163606L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            val attrs = HashMap<String, String>()
            var nv: NvPicPr? = null
            var bf: BlipFill? = null
            var sp: SpPr? = null
            var s: Style? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pic") {        // get attributes
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                        } else if (tnm == "nvPicPr") {
                            lastTag.push(tnm)
                            nv = NvPicPr.parseOOXML(xpp, lastTag)
                        } else if (tnm == "blipFill") {
                            lastTag.push(tnm)
                            bf = BlipFill.parseOOXML(xpp, lastTag)
                        } else if (tnm == "spPr") {
                            lastTag.push(tnm)
                            sp = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //sp.setNS("xdr");
                        } else if (tnm == "style") {
                            lastTag.push(tnm)
                            s = Style.parseOOXML(xpp, lastTag, bk) as Style
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pic") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("pic.parseOOXML: $e")
            }

            return Pic(attrs, nv, bf, sp, s)
        }
    }
}

/**
 * nvPicPr (Non-Visual Properties for a Picture)
 * This element specifies all non-visual properties for a picture. This element is a container for the non-visual
 * identification properties, shape properties and application properties that are to be associated with a picture.
 * This allows for additional information that does not affect the appearance of the picture to be stored.
 *
 *
 * parent: pic
 * children: cNvPr, cNvPicPr
 */
internal class NvPicPr : OOXMLElement {
    private var cpr: CNvPr? = null
    private var ppr: CNvPicPr? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:nvPicPr>")
            ooxml.append(cpr!!.ooxml)
            ooxml.append(ppr!!.ooxml)
            ooxml.append("</xdr:nvPicPr>")
            return ooxml.toString()
        }

    /**
     * return the name of this shape, if any
     *
     * @return
     */
    /**
     * set the name of this shape, if any
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
     * set cNvPr desc attribute
     *
     * @param name
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

    constructor() {    // set common defaults
        this.cpr = CNvPr()
        this.ppr = CNvPicPr()
    }

    constructor(cpr: CNvPr, ppr: CNvPicPr) {
        this.cpr = cpr
        this.ppr = ppr
    }

    constructor(n: NvPicPr) {
        this.cpr = n.cpr
        this.ppr = n.ppr
    }

    override fun cloneElement(): OOXMLElement {
        return NvPicPr(this)
    }

    companion object {

        private val serialVersionUID = -3722424348721713313L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): NvPicPr {
            var cpr: CNvPr? = null
            var ppr: CNvPicPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvPr") {
                            lastTag.push(tnm)
                            cpr = CNvPr.parseOOXML(xpp, lastTag) as CNvPr
                        } else if (tnm == "cNvPicPr") {
                            lastTag.push(tnm)
                            ppr = CNvPicPr.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "nvPicPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("nvPicPr.parseOOXML: $e")
            }

            return NvPicPr(cpr, ppr)
        }
    }

}

/**
 * cNvPicPr (Non-Visual Picture Drawing Properties)
 * This element describes the non-visual properties of a picture within a spreadsheet. These are the set of
 * properties of a picture which do not affect its display within a spreadsheet.
 */
// TODO: handle child picLocks
internal class CNvPicPr : OOXMLElement {
    private var preferRelativeResize: String? = null

    override// TODO: finish picLocks
    //if (p!=null) ooxml.append(p.getOOXML());
    //ooxml.append("</>");
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>")
            return ooxml.toString()
        }

    constructor() {

    }

    constructor(preferRelativeResize: String) {
        this.preferRelativeResize = preferRelativeResize
    }

    constructor(c: CNvPicPr) {
        this.preferRelativeResize = c.preferRelativeResize
    }

    override fun cloneElement(): OOXMLElement {
        return CNvPicPr(this)
    }

    companion object {

        private val serialVersionUID = 3690228348761065940L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): CNvPicPr {
            var preferRelativeResize: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvPicPr") {        // get attributes
                            if (xpp.attributeCount > 0)
                                preferRelativeResize = xpp.getAttributeValue(0)
                            /* TODO: Finish		} else if (tnm.equals("picLocks")) {
	        			 lastTag.push(tnm);
	        			 //p = (picLocks) picLocks.parseOOXML(xpp, lastTag).clone();
*/
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cNvPicPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cNvPicPr.parseOOXML: $e")
            }

            return CNvPicPr(preferRelativeResize)
        }
    }
}

