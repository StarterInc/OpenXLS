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
 * sp (Shape)
 *
 *
 * OOXML/DrawingML element representing a single shape
 * A shape can either be a preset or a custom geometry,
 * defined using the DrawingML framework. In addition to a geometry each shape can have both visual and non
 * visual properties attached. Text and corresponding styling information can also be attached to a shape.
 *
 *
 * parent = twoCellAnchor, oneCellAnchor, absoluteAnchor, grpSp (group shape)
 * children:  nvSpPr, spPr, style, txBody
 */
class Sp : OOXMLElement {
    private var nvsp: NvSpPr? = null
    private var sppr: SpPr? = null
    private var sty: Style? = null
    private var txb: TxBody? = null
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:sp")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            ooxml.append(nvsp!!.ooxml)
            ooxml.append(sppr!!.ooxml)
            if (sty != null) ooxml.append(sty!!.ooxml)
            if (txb != null) ooxml.append(txb!!.ooxml)
            ooxml.append("</xdr:sp>")
            return ooxml.toString()
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
        get() = if (sppr != null) sppr!!.embed else null
        set(embed) {
            if (sppr != null) sppr!!.embed = embed
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
        get() = if (sppr != null) sppr!!.link else null
        set(link) {
            if (sppr != null) sppr!!.link = link
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
        get() = if (nvsp != null) nvsp!!.name else null
        set(name) {
            if (nvsp != null)
                nvsp!!.name = name
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
        get() = if (nvsp != null) nvsp!!.descr else null
        set(descr) {
            if (nvsp != null)
                nvsp!!.descr = descr
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
        get() = if (nvsp != null) nvsp!!.id else -1
        set(id) {
            if (nvsp != null)
                nvsp!!.id = id
        }

    constructor(nvsp: NvSpPr, sppr: SpPr, sty: Style, txb: TxBody, attrs: HashMap<String, String>) {
        this.nvsp = nvsp
        this.sppr = sppr
        this.sty = sty
        this.txb = txb
        this.attrs = attrs
    }

    constructor(shp: Sp) {
        this.nvsp = shp.nvsp
        this.sppr = shp.sppr
        this.sty = shp.sty
        this.txb = shp.txb
        this.attrs = shp.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Sp(this)
    }

    companion object {

        private val serialVersionUID = 7454285931503575078L


        /**
         * return the OOXML specific for this object
         *
         * @param xpp
         * @param lastTag
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var nvsp: NvSpPr? = null
            var sppr: SpPr? = null
            var sty: Style? = null
            var txb: TxBody? = null
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "sp") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "nvSpPr") {        // non-visual shape props
                            lastTag.push(tnm)
                            nvsp = NvSpPr.parseOOXML(xpp, lastTag)
                        } else if (tnm == "spPr") {        // shape properties
                            lastTag.push(tnm)
                            sppr = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //sppr.setNS("xdr");
                        } else if (tnm == "style") {        // shape style
                            lastTag.push(tnm)
                            sty = Style.parseOOXML(xpp, lastTag, bk) as Style
                        } else if (tnm == "txBody") {        // shape text body
                            lastTag.push(tnm)
                            txb = TxBody.parseOOXML(xpp, lastTag, bk) as TxBody
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "sp") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("sp.parseOOXML: $e")
            }

            return Sp(nvsp, sppr, sty, txb, attrs)
        }
    }

}

/**
 * This element specifies all non-visual properties for a shape. This element is a container for the non-visual
 * identification properties, shape properties and application properties that are to be associated with a shape.
 * This allows for additional information that does not affect the appearance of the shape to be stored
 *
 *
 * parent:  sp
 * children: cNvPr  REQ cNvSpPr REQ
 */
internal class NvSpPr : OOXMLElement {
    private var cnv: CNvPr? = null
    private var cnvsp: CNvSpPr? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:nvSpPr>")
            ooxml.append(cnv!!.ooxml)
            ooxml.append(cnvsp!!.ooxml)
            ooxml.append("</xdr:nvSpPr>")
            return ooxml.toString()
        }

    /**
     * get name attribute
     *
     * @return
     */
    /**
     * set name attribute
     *
     * @param name
     */
    var name: String?
        get() = if (cnv != null) cnv!!.name else null
        set(name) {
            if (cnv != null)
                cnv!!.name = name
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
        get() = if (cnv != null) cnv!!.descr else null
        set(descr) {
            if (cnv != null)
                cnv!!.descr = descr
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
        get() = if (cnv != null) cnv!!.id else -1
        set(id) {
            if (cnv != null)
                cnv!!.id = id
        }

    constructor(cnv: CNvPr, cnvsp: CNvSpPr) {
        this.cnv = cnv
        this.cnvsp = cnvsp
    }

    constructor(nvsp: NvSpPr) {
        this.cnv = nvsp.cnv
        this.cnvsp = nvsp.cnvsp
    }

    override fun cloneElement(): OOXMLElement {
        return NvSpPr(this)
    }

    companion object {

        private val serialVersionUID = 9121235009516398367L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): NvSpPr {
            var cnv: CNvPr? = null
            var cnvsp: CNvSpPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvPr") {
                            lastTag.push(tnm)
                            cnv = CNvPr.parseOOXML(xpp, lastTag) as CNvPr
                        } else if (tnm == "cNvSpPr") {
                            lastTag.push(tnm)
                            cnvsp = CNvSpPr.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "nvSpPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("txBody.parseOOXML: $e")
            }

            return NvSpPr(cnv, cnvsp)
        }
    }
}

/**
 * cNvSpPr (Non-visual Drawing Properties for a shape
 *
 *
 * This OOXML/DrawingML element specifies the non-visual drawing properties for a shape.
 * These properties are to be used by the generating application to determine how the shape should be dealt with
 *
 *
 * attributes:  txBox (optional boolean)
 * parent:    nvSpPr
 * children:  spLocks (shapeLocks)
 */
internal class CNvSpPr : OOXMLElement {
    private var txBox: String? = null
    private var sp: SpLocks? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:cNvSpPr")
            if (this.txBox != null) ooxml.append(" txBox=\"$txBox\"")
            ooxml.append(">")
            if (sp != null) ooxml.append(sp!!.ooxml)
            ooxml.append("</xdr:cNvSpPr>")
            return ooxml.toString()
        }

    constructor(t: String, sp: SpLocks) {
        this.txBox = t
        this.sp = sp
    }

    constructor(c: CNvSpPr) {
        this.txBox = c.txBox
        this.sp = c.sp
    }

    override fun cloneElement(): OOXMLElement {
        return CNvSpPr(this)
    }

    companion object {

        private val serialVersionUID = 7895953516797436713L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): CNvSpPr {
            var txBox: String? = null
            var sp: SpLocks? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvSpPr") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                if (xpp.getAttributeName(i) == "txBox")
                                    txBox = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "spLocks") {
                            lastTag.push(tnm)
                            sp = SpLocks.parseOOXML(xpp, lastTag)

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cNvSpPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cNvSpPr.parseOOXML: $e")
            }

            return CNvSpPr(txBox, sp)
        }
    }
}

/**
 * spLocks (Shape Locks)
 *
 *
 * OOXML/DrawingML element specifies all locking properties for a shape.
 * These properties inform the generating application
 * about specific properties that have been previously locked and thus should not be changed.
 *
 *
 * parents:    cNvSpPr
 */
internal class SpLocks : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:spLocks")
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

    constructor(sp: SpLocks) {
        this.attrs = sp.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return SpLocks(this)
    }

    companion object {

        private val serialVersionUID = -3805557220039550941L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): SpLocks {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "spLocks") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "spLocks") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("spLocks.parseOOXML: $e")
            }

            return SpLocks(attrs)
        }
    }
}


