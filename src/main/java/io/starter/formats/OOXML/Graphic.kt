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

import java.util.Stack

/**
 * graphic (Graphic Object)
 * This element specifies the existence of a single graphic object. Document authors should refer to this element
 * when they wish to persist a graphical object of some kind. The specification for this graphical object will be
 * provided entirely by the document author and referenced within the graphicData child element
 *
 *
 * parent:  anchor, graphicFrame, inline
 * children: graphicData
 */
class Graphic : OOXMLElement {
    private var graphicData: GraphicData? = GraphicData()

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:graphic>")
            if (graphicData != null) ooxml.append(graphicData!!.ooxml)
            ooxml.append("</a:graphic>")
            return ooxml.toString()
        }

    /**
     * get the URI associated with this graphic Data
     */
    var uri: String?
        get() = if (graphicData != null) graphicData!!.uri else null
        set(uri) {
            if (graphicData != null)
                graphicData!!.uri = uri
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
        get() = if (graphicData != null) graphicData!!.chartRId else null
        set(rid) {
            if (graphicData != null) graphicData!!.chartRId = rid
        }

    constructor() {}

    constructor(g: GraphicData) {
        this.graphicData = g
    }

    constructor(gr: Graphic) {
        this.graphicData = gr.graphicData
    }

    override fun cloneElement(): OOXMLElement {
        return Graphic(this)
    }

    companion object {

        private val serialVersionUID = -7027946026352255398L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            var g: GraphicData? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "graphicData") {
                            lastTag.push(tnm)
                            g = GraphicData.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "graphic") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("graphic.parseOOXML: $e")
            }

            return Graphic(g)
        }
    }
}

/**
 * graphicData (Graphic Object Data)
 * This element specifies the reference to a graphic object within the document. This graphic object is provided
 * entirely by the document authors who choose to persist this data within the document.
 *
 *
 * parent: graphic
 * children: chart ... anything else?
 */
internal class GraphicData : OOXMLElement {
    /**
     * return the URI attribute associated with this graphic Data
     *
     * @return
     */
    /**
     * set the URI attribute for this graphic data
     *
     * @param uri
     */
    var uri: String? = OOXMLConstants.chartns //xmlns:r=\"" + OOXMLConstants.relns + "\"";	// default
    /**
     * return the rid of the chart element, if exists
     *
     * @return
     */
    /**
     * set the rid of the chart element
     */
    var chartRId: String? = null

    override// we'll assume it's a chart for nwo
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:graphicData")
            if (uri != null)
                ooxml.append(" uri=\"$uri\"")
            if (chartRId != null) {
                ooxml.append("><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"  r:id=\"$chartRId\"/></a:graphicData>")
            } else
                ooxml.append("/>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(uri: String, rid: String) {
        this.uri = uri
        this.chartRId = rid
    }

    constructor(gd: GraphicData) {
        this.uri = gd.uri
        this.chartRId = gd.chartRId
    }

    override fun cloneElement(): OOXMLElement {
        return GraphicData(this)
    }

    companion object {

        private val serialVersionUID = 7395991759307532325L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): GraphicData {
            var uri: String? = null
            var rid: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "graphicData") {        // get attributes
                            if (xpp.attributeCount > 0)
                                uri = xpp.getAttributeValue(0)
                        } else if (tnm == "chart") {        // one of many possible children
                            if (xpp.attributeCount > 0)
                                rid = xpp.getAttributeValue(0)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "graphicData") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("graphicData.parseOOXML: $e")
            }

            return GraphicData(uri, rid)
        }
    }
}

