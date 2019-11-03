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
 * oneCellAnchor (One Cell Anchor Shape Size)
 *
 *
 * This element specifies a one cell anchor placeholder for a group, a shape, or a drawing element. It moves with
 * the cell and its extents is in EMU units.
 *
 *
 * parent: wsDr
 * children: from, ext, OBJECTCHOICES (sp, grpSp, graphicFrame, cxnSp, pic), clientData
 */
//TODO: finish grpSp Group Shape
// TODO: finish clientData element
class OneCellAnchor : OOXMLElement {
    private var from: From? = null
    private var ext: Ext? = null
    private var objectChoice: ObjectChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:oneCellAnchor>")
            if (from != null) ooxml.append(from!!.ooxml)
            ooxml.append(ext!!.ooxml)
            ooxml.append(objectChoice!!.ooxml)
            ooxml.append("<xdr:clientData/>")
            ooxml.append("</xdr:oneCellAnchor>")
            return ooxml.toString()
        }

    // access methods ******

    /**
     * return the bounds of this object
     * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST
     *
     * @return bounds short[4]
     */
    // from bounds
    val bounds: ShortArray
        get() {
            val bounds = ShortArray(8)
            System.arraycopy(from!!.bounds, 0, bounds, 0, 4)
            return bounds
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
        get() = if (objectChoice != null) objectChoice!!.name else null
        set(name) {
            if (objectChoice != null)
                objectChoice!!.name = name
        }

    /**
     * get cNvPr descr attribute
     *
     * @return
     */
    /**
     * set cNvPr descr attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var descr: String?
        get() = if (objectChoice != null) objectChoice!!.descr else null
        set(descr) {
            if (objectChoice != null)
                objectChoice!!.descr = descr
        }

    /**
     * get macro attribute
     *
     * @return
     */
    /**
     * set Macro attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var macro: String?
        get() = if (objectChoice != null) objectChoice!!.macro else null
        set(macro) {
            if (objectChoice != null)
                objectChoice!!.macro = macro
        }

    /**
     * get the URI associated with this graphic Data
     */
    /**
     * set the URI associated with this graphic data
     *
     * @param uri
     */
    var uri: String?
        get() = if (objectChoice != null) objectChoice!!.uri else null
        set(uri) {
            if (objectChoice != null)
                objectChoice!!.uri = uri
        }

    /**
     * return the id for the embedded picture, shape or chart (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed or rId attribute for the embedded picture, shape or chart (i.e. resides within the file)
     *
     * @param embed
     */
    var embed: String?
        get() = if (objectChoice != null) objectChoice!!.embed else null
        set(embed) {
            if (objectChoice != null) objectChoice!!.embed = embed
        }

    /**
     * return the id for the linked object (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (objectChoice != null) objectChoice!!.link else null
        set(link) {
            if (objectChoice != null) objectChoice!!.link = link
        }

    /**
     * utility to return the shape properties element (picture element only)
     * should be depreciated when OOXML is moved into ImageHandle
     *
     * @return
     */
    val sppr: SpPr?
        get() = if (objectChoice != null) objectChoice!!.sppr else null

    constructor(f: From, e: Ext, o: ObjectChoice) {
        this.from = f
        this.ext = e
        this.objectChoice = o
    }

    constructor(oca: OneCellAnchor) {
        this.from = oca.from
        this.ext = oca.ext
        this.objectChoice = oca.objectChoice
    }

    override fun cloneElement(): OOXMLElement {
        return OneCellAnchor(this)
    }

    /**
     * set the bounds of this object
     * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST
     *
     * @param bounds short[4]
     */
    fun setBounds(bounds: IntArray) {
        val b = IntArray(4)
        System.arraycopy(bounds, 0, b, 0, 4)
        if (from == null)
            from = From(b)
        else
            from!!.bounds = b
    }

    /**
     * return if this oneCellAnchor element refers to an image rather than a chart or shape
     *
     * @return
     */
    fun hasImage(): Boolean {
        return if (objectChoice != null) objectChoice!!.`object` is Pic && objectChoice!!.embed != null else false
    }

    /**
     * return if this oneCellAnchor element refers to a chart as opposed to a shape or image
     *
     * @return
     */
    fun hasChart(): Boolean {
        return if (objectChoice != null) objectChoice!!.`object` is GraphicFrame && objectChoice!!.chartRId != null else false
    }

    /**
     * return if this oneCellAnchor element refers to a shape, as opposed a chart or an image
     *
     * @return
     */
    fun hasShape(): Boolean {
        return if (objectChoice != null) objectChoice!!.`object` is CxnSp || objectChoice!!.`object` is Sp else false
    }

    /**
     * set this oneCellAnchor as a chart element
     *
     * @param rid
     * @param name
     * @param bounds
     */
    fun setAsChart(rid: Int, name: String, bounds: IntArray) {
        objectChoice = ObjectChoice()
        objectChoice!!.`object` = GraphicFrame()
        objectChoice!!.name = name
        objectChoice!!.embed = "rId" + Integer.valueOf(rid)!!.toString()
        objectChoice!!.id = rid
        this.setBounds(bounds)
        // id???
    }

    /**
     * set this oneCellAnchor as an image
     *
     * @param rid
     * @param name
     * @param id
     */
    fun setAsImage(rid: String, name: String, id: String) {
        val o = ObjectChoice()
        o.`object` = Pic()
        o.name = name
        o.embed = rid
    }

    companion object {

        private val serialVersionUID = -8498556079325357165L
        val EMU: Short = 1270

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var f: From? = null
            var e: Ext? = null
            var o: ObjectChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "from") {
                            lastTag.push(tnm)
                            f = From.parseOOXML(xpp, lastTag)
                        } else if (tnm == "ext") {
                            lastTag.push(tnm)
                            e = Ext.parseOOXML(xpp, lastTag).cloneElement() as Ext
                            //e.setNS("xdr");
                        } else if (tnm == "cxnSp" ||    // connection shape

                                tnm == "graphicFrame" ||
                                tnm == "grpSp" ||    // group shape

                                tnm == "pic" ||    // picture

                                tnm == "sp") {        // shape
                            lastTag.push(tnm)
                            o = ObjectChoice.parseOOXML(xpp, lastTag, bk) as ObjectChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "oneCellAnchor") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (ex: Exception) {
                Logger.logErr("oneCellAnchor.parseOOXML: $ex")
            }

            return OneCellAnchor(f, e, o)
        }
    }
}

