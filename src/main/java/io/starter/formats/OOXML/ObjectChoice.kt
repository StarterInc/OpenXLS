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
 * One of:  cxnSp, graphicFrame, grpSp, pic, sp
 */
class ObjectChoice : OOXMLElement {
    private var cxnSp: CxnSp? = null
    private var graphicFrame: GraphicFrame? = null
    private var grpSp: GrpSp? = null
    private var pic: Pic? = null
    private var sp: Sp? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (cxnSp != null) ooxml.append(cxnSp!!.ooxml)
            if (graphicFrame != null) ooxml.append(graphicFrame!!.ooxml)
            if (grpSp != null) ooxml.append(grpSp!!.ooxml)
            if (pic != null) ooxml.append(pic!!.ooxml)
            if (sp != null) ooxml.append(sp!!.ooxml)
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
        get() {
            if (cxnSp != null)
                return cxnSp!!.name
            else if (sp != null)
                return sp!!.name
            else if (pic != null)
                return pic!!.name
            else if (graphicFrame != null)
                return graphicFrame!!.name
            else if (grpSp != null)
                return grpSp!!.name
            return null
        }
        set(name) {
            if (cxnSp != null)
                cxnSp!!.name = name
            else if (sp != null)
                sp!!.name = name
            else if (pic != null)
                pic!!.name = name
            else if (graphicFrame != null)
                graphicFrame!!.name = name
            else if (grpSp != null)
                grpSp!!.name = name
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
        get() {
            if (cxnSp != null)
                return cxnSp!!.descr
            else if (sp != null)
                return sp!!.descr
            else if (pic != null)
                return pic!!.descr
            else if (graphicFrame != null)
                return graphicFrame!!.descr
            else if (grpSp != null)
                return grpSp!!.descr
            return null
        }
        set(descr) {
            if (cxnSp != null)
                cxnSp!!.descr = descr
            else if (sp != null)
                sp!!.descr = descr
            else if (pic != null)
                pic!!.descr = descr
            else if (graphicFrame != null)
                graphicFrame!!.descr = descr
            else if (grpSp != null)
                grpSp!!.descr = descr
        }

    /**
     * get macro attribute (valid for cnxSp, sp and graphicFrame)
     *
     * @return
     */
    /**
     * set Macro attribute (valid for cnxSp, sp and graphicFrame)
     * sometimes associated with shape name
     *
     * @param descr
     */
    var macro: String?
        get() {
            if (cxnSp != null)
                return cxnSp!!.macro
            else if (graphicFrame != null)
                return graphicFrame!!.macro
            else if (sp != null)
                return sp!!.macro
            return null
        }
        set(macro) {
            if (cxnSp != null)
                cxnSp!!.macro = macro
            else if (graphicFrame != null)
                graphicFrame!!.macro = macro
            else if (sp != null)
                sp!!.macro = macro
            else if (grpSp != null)
                grpSp!!.macro = macro
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
        get() = if (graphicFrame != null) graphicFrame!!.uri else null
        set(uri) {
            if (graphicFrame != null)
                graphicFrame!!.uri = uri
            else if (grpSp != null)
                grpSp!!.setURI(uri)
        }

    /**
     * return the rid for the embedded object (picture or picture shape) (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the rid for this object (picture or picture shape) (resides within the file)
     *
     * @param embed
     */
    // embedded blip/pict
    // embedded image
    // group shape embedded image
    var embed: String?
        get() {
            if (sp != null)
                return sp!!.embed
            else if (pic != null)
                return pic!!.embed
            else if (grpSp != null)
                return grpSp!!.embed
            return null
        }
        set(embed) {
            if (sp != null)
                sp!!.embed = embed
            else if (pic != null)
                pic!!.embed = embed
            else if (grpSp != null) grpSp!!.embed = embed
        }


    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param link
     */
    var link: String?
        get() {
            if (sp != null)
                return sp!!.link
            else if (pic != null)
                return pic!!.link
            else if (grpSp != null)
                return grpSp!!.link
            return null
        }
        set(link) {
            if (sp != null)
                sp!!.link = link
            else if (pic != null)
                pic!!.link = link
            else if (grpSp != null) grpSp!!.link = link
        }

    /**
     * return the rid of the chart element, if exists
     *
     * @return
     */
    /**
     * set the rid for this chart (resides within the file)
     *
     * @param rId
     */
    var chartRId: String?
        get() {
            if (graphicFrame != null)
                return graphicFrame!!.chartRId
            else if (grpSp != null)
                return grpSp!!.chartRId
            return null
        }
        set(rId) {
            if (graphicFrame != null) graphicFrame!!.chartRId = rId
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
    // embedded blip/pict
    // embedded image
    // chart
    // embedded image
    // embedded blip/pict
    // embedded image
    // chart
    // embedded image
    var id: Int
        get() {
            if (sp != null)
                return sp!!.id
            else if (pic != null)
                return pic!!.id
            else if (graphicFrame != null)
                return graphicFrame!!.id
            else if (cxnSp != null)
                return cxnSp!!.id
            else if (grpSp != null)
                return grpSp!!.id
            return -1
        }
        set(id) {
            if (sp != null)
                sp!!.id = id
            else if (pic != null)
                pic!!.id = id
            else if (graphicFrame != null)
                graphicFrame!!.id = id
            else if (cxnSp != null)
                cxnSp!!.id = id
            else if (grpSp != null)
                grpSp!!.id = id
        }

    /**
     * utility to return the shape properties element for this picture
     * should be depreciated when OOXML is completely distinct from BIFF8
     *
     * @return
     */
    val sppr: SpPr?
        get() {
            if (pic != null)
                return pic!!.sppr
            else if (grpSp != null)
                return grpSp!!.sppr
            return null
        }

    /**
     * return the actual object associated with this ObjectChoice
     *
     * @return
     */
    /**
     * set the object associated with this ObjectChoice
     *
     * @param o
     */
    var `object`: Any?
        get() {
            if (cxnSp != null)
                return cxnSp
            else if (graphicFrame != null)
                return graphicFrame
            else if (pic != null)
                return pic
            else if (sp != null)
                return sp
            else if (grpSp != null)
                return grpSp
            return null
        }
        set(o) {
            if (o is GraphicFrame) {
                graphicFrame = o
            } else if (o is Pic) {
                pic = o
            } else if (o is Sp) {
                sp = o
            } else if (o is CxnSp) {
                cxnSp = o
            } else if (o is GrpSp) {
                grpSp = o
            }
        }

    constructor() {

    }

    constructor(c: CxnSp, g: GraphicFrame, grp: GrpSp, p: Pic, s: Sp) {
        this.cxnSp = c
        this.graphicFrame = g
        this.grpSp = grp
        this.pic = p
        this.sp = s
    }

    constructor(oc: ObjectChoice) {
        this.cxnSp = oc.cxnSp
        this.graphicFrame = oc.graphicFrame
        this.grpSp = oc.grpSp
        this.pic = oc.pic
        this.sp = oc.sp
    }

    override fun cloneElement(): OOXMLElement {
        return ObjectChoice(this)
    }

    /**
     * return if this Object Choice refers to an image rather than a chart or shape
     *
     * @return
     */
    fun hasImage(): Boolean {
        // o will be a pic element or a group shape containing a pic element, it's blipFill.blip child references the rId of the embedded file
        return this.embed != null
    }

    /**
     * return if this Object Choice refers to a shape, as opposed a chart or an image
     *
     * @return
     */
    fun hasShape(): Boolean {
        return cxnSp != null || sp != null || grpSp != null && grpSp!!.hasShape()
    }

    /**
     * return if this Object Choice element refers to a chart as opposed to a shape or image
     *
     * @return
     */
    fun hasChart(): Boolean {
        return this.chartRId != null
    }

    companion object {

        private val serialVersionUID = 3548474869557092714L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var c: CxnSp? = null
            var g: GraphicFrame? = null
            var grp: GrpSp? = null
            var p: Pic? = null
            var s: Sp? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cxnSp") {    // connection shape
                            lastTag.push(tnm)
                            c = CxnSp.parseOOXML(xpp, lastTag, bk) as CxnSp
                            break
                        } else if (tnm == "graphicFrame") {    // graphic data usually chart
                            lastTag.push(tnm)
                            g = GraphicFrame.parseOOXML(xpp, lastTag) as GraphicFrame
                            break
                        } else if (tnm == "grpSp") {    // group shape - combines one or more of sp/pic/graphicFrame/cxnSp
                            lastTag.push(tnm)
                            grp = GrpSp.parseOOXML(xpp, lastTag, bk)
                            break
                        } else if (tnm == "sp") {        // shape
                            lastTag.push(tnm)
                            s = Sp.parseOOXML(xpp, lastTag, bk) as Sp
                            break
                        } else if (tnm == "pic") {        // picture/image
                            lastTag.push(tnm)
                            p = Pic.parseOOXML(xpp, lastTag, bk) as Pic
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("ObjectChoice.parseOOXML: $e")
            }

            return ObjectChoice(c, g, grp, p, s)
        }
    }
}