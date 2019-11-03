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
 * graphicFrame (Graphic Frame)
 * This element describes a single graphical object frame for a spreadsheet which contains a graphical object
 *
 *
 * parent: oneCellAnchor, twoCellAnchor, absoluteAnchor, grpSp
 * children: graphic, nvGraphicFramePr, xfrm (all required and in sequence)
 */
//TODO: finish cNvGraphicFramePr.graphicFrameLocks element
class GraphicFrame : OOXMLElement {
    private var attrs = HashMap<String, String>()
    private var graphic: Graphic? = Graphic()

    private var graphicFramePr: NvGraphicFramePr? = NvGraphicFramePr()
    private var xfrm = Xfrm()

    override// attributes
    // all are required so no null checks - must ensure x.ns is set
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:graphicFrame")
            val i = attrs.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs[key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            ooxml.append(graphicFramePr!!.ooxml)
            ooxml.append(xfrm.ooxml)
            ooxml.append(graphic!!.ooxml)
            ooxml.append("</xdr:graphicFrame>")
            return ooxml.toString()
        }

    /**
     * get graphicFrame Macro attribute
     */
    /**
     * set graphicFrame Macro attribute
     *
     * @param macro
     */
    var macro: String?
        get() = if (attrs["macro"] != null) attrs["macro"] else null
        set(macro) {
            attrs["macro"] = macro
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
        get() = if (graphic != null) graphic!!.uri else null
        set(uri) {
            if (graphic != null)
                graphic!!.uri = uri
        }

    /**
     * return the rid of the chart element, if exists
     *
     * @return
     */
    /**
     * set the rid of the chart element
     */
    var chartRId: String?
        get() = if (graphic != null) graphic!!.chartRId else null
        set(rid) {
            if (graphic != null) graphic!!.chartRId = rid
        }

    /**
     * get the cNvPr name attribute
     *
     * @return
     */
    /**
     * set cNvPr name attribute
     *
     * @param name
     */
    var name: String?
        get() = if (graphicFramePr != null) graphicFramePr!!.name else null
        set(name) {
            if (graphicFramePr != null)
                graphicFramePr!!.name = name
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
        get() = if (graphicFramePr != null) graphicFramePr!!.descr else null
        set(descr) {
            if (graphicFramePr != null)
                graphicFramePr!!.descr = descr
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
        get() = if (graphicFramePr != null) graphicFramePr!!.id else -1
        set(id) {
            if (graphicFramePr != null)
                graphicFramePr!!.id = id
        }

    constructor() {
        attrs["macro"] = ""
    }

    constructor(attrs: HashMap<String, String>, g: Graphic, gfp: NvGraphicFramePr, x: Xfrm) {
        this.attrs = attrs
        this.graphic = g
        this.graphicFramePr = gfp
        this.xfrm = x
    }

    constructor(gf: GraphicFrame) {
        this.attrs = gf.attrs
        this.graphic = gf.graphic
        this.graphicFramePr = gf.graphicFramePr
        this.xfrm = gf.xfrm
    }

    override fun cloneElement(): OOXMLElement {
        return GraphicFrame(this)
    }

    companion object {

        private val serialVersionUID = 2494490998000511917L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            val attrs = HashMap<String, String>()
            var g: Graphic? = null
            var gfp: NvGraphicFramePr? = null
            var x: Xfrm? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "graphicFrame") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "xfrm") {
                            lastTag.push(tnm)
                            x = Xfrm.parseOOXML(xpp, lastTag) as Xfrm
                            x.setNS("xdr")
                        } else if (tnm == "graphic") {
                            lastTag.push(tnm)
                            g = Graphic.parseOOXML(xpp, lastTag) as Graphic
                        } else if (tnm == "nvGraphicFramePr") {
                            lastTag.push(tnm)
                            gfp = NvGraphicFramePr.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "graphicFrame") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("graphicFrame.parseOOXML: $e")
            }

            return GraphicFrame(attrs, g, gfp, x)
        }
    }
}

/**
 * nvGraphicFramePr (Non-Visual Properties for a Graphic Frame)
 * This element specifies all non-visual properties for a graphic frame. This element is a container for the non-visual
 * identification properties, shape properties and application properties that are to be associated with a graphic
 * frame. This allows for additional information that does not affect the appearance of the graphic frame to be
 * stored.
 *
 *
 * parent: graphicFrame
 * children: cNvPr REQ cNvGraphicFramePr REQ
 */
internal class NvGraphicFramePr : OOXMLElement {
    private var cp: CNvPr? = CNvPr()
    private var nvpr: CNvGraphicFramePr? = CNvGraphicFramePr()

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:nvGraphicFramePr>")
            if (cp != null) ooxml.append(cp!!.ooxml)
            if (nvpr != null) ooxml.append(nvpr!!.ooxml)
            ooxml.append("</xdr:nvGraphicFramePr>")
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
        get() = if (cp != null) cp!!.name else null
        set(name) {
            if (cp != null)
                cp!!.name = name
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
        get() = if (cp != null) cp!!.descr else null
        set(descr) {
            if (cp != null)
                cp!!.descr = descr
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
        get() = if (cp != null) cp!!.id else -1
        set(id) {
            if (cp != null)
                cp!!.id = id
        }

    constructor() {}

    constructor(cp: CNvPr, nvpr: CNvGraphicFramePr) {
        this.cp = cp
        this.nvpr = nvpr
    }

    constructor(nvg: NvGraphicFramePr) {
        this.cp = nvg.cp
        this.nvpr = nvg.nvpr
    }

    override fun cloneElement(): OOXMLElement {
        return NvGraphicFramePr(this)
    }

    companion object {
        private val serialVersionUID = -47476384268955296L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): NvGraphicFramePr {
            var cp: CNvPr? = null
            var nvpr: CNvGraphicFramePr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvGraphicFramePr") {
                            lastTag.push(tnm)
                            nvpr = CNvGraphicFramePr.parseOOXML(xpp, lastTag)
                        } else if (tnm == "cNvPr") {
                            lastTag.push(tnm)
                            cp = CNvPr.parseOOXML(xpp, lastTag) as CNvPr
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "nvGraphicFramePr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("nvGraphicFramePr.parseOOXML: $e")
            }

            return NvGraphicFramePr(cp, nvpr)
        }
    }


}

/**
 * cNvGraphicFramePr (Non-Visual Graphic Frame Drawing Properties)
 *
 *
 * This element specifies the non-visual properties for a single graphical object frame within a spreadsheet. These
 * are the set of properties of a frame which do not affect its display within a spreadsheet
 *
 *
 * parent: nvGraphicFramePr
 * children: graphicFrameLocks
 */
// TODO: finish graphicFrameLocks
internal class CNvGraphicFramePr : OOXMLElement {

    override// TODO: Finish child graphicFrameLocks
    //if (gf!=null) ooxml.append(gf.getOOXML());
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:cNvGraphicFramePr>")
            ooxml.append("<a:graphicFrameLocks/>")
            ooxml.append("</xdr:cNvGraphicFramePr>")
            return ooxml.toString()
        }

    // private graphicFrameLocks gf= null;
    constructor() {}

    constructor(g: CNvGraphicFramePr) {}

    override fun cloneElement(): OOXMLElement {
        return CNvGraphicFramePr(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 769474804434194488L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): CNvGraphicFramePr {
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        /*		            String tnm = xpp.getName();
		            if (tnm.equals("graphicFrameLocks")) {
	        			 lastTag.push(tnm);
	        			 //gf= (graphicFrameLocks) graphicFrameLocks.parseOOXML(xpp, lastTag).clone();
		            }
*/
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cNvGraphicFramePr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cNvGraphicFramePr.parseOOXML: $e")
            }

            return CNvGraphicFramePr()
        }
    }

}

