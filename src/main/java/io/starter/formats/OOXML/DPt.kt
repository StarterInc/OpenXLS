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
 * dPt (Data Point)
 * This element specifies a single data point.
 *
 *
 * parent: series
 * children: idx REQ, invertIfNegative, marker, bubble3D, explosion, spPr, pictureOptions
 */
// TODO: finish pictureOptions
class DPt : OOXMLElement {
    private var idx: Int = 0
    private var invertIfNegative: Boolean = false
    private var bubble3D: Boolean = false
    private var marker: Marker? = null
    private var spPr: SpPr? = null
    private var explosion: Int = 0

    override// default= true
    // default= true
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<c:dPt>")
            ooxml.append("<c:idx val=\"$idx\"/>")
            if (!invertIfNegative) ooxml.append("<c:invertIfNegative val=\"0\"/>")
            if (marker != null) ooxml.append(marker!!.ooxml)
            if (!bubble3D) ooxml.append("<c:bubble3D val=\"0\"/>")
            if (explosion != 0) ooxml.append("<c:explosion val=\"$explosion\"/>")
            if (spPr != null) ooxml.append(spPr!!.ooxml)
            ooxml.append("</c:dPt>")
            return ooxml.toString()
        }

    constructor(idx: Int, invertIfNegative: Boolean, bubble3D: Boolean, m: Marker, sp: SpPr, explosion: Int) {
        this.idx = idx
        this.invertIfNegative = invertIfNegative
        this.bubble3D = bubble3D
        this.marker = m
        this.spPr = sp
        this.explosion = explosion
    }

    constructor(d: DPt) {
        this.idx = d.idx
        this.invertIfNegative = d.invertIfNegative
        this.bubble3D = d.bubble3D
        this.marker = d.marker
        this.spPr = d.spPr
        this.explosion = d.explosion
    }

    override fun cloneElement(): OOXMLElement {
        return DPt(this)
    }

    companion object {

        private val serialVersionUID = 8354707071603571747L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var idx = -1
            var invertIfNegative = true
            var bubble3D = true
            var m: Marker? = null
            var sp: SpPr? = null
            var explosion = 0
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "idx") {    // child element only contains 1 element
                            if (xpp.attributeCount > 0)
                                idx = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "invertIfNegative") {    // child element only contains 1 element
                            if (xpp.attributeCount > 0)
                                invertIfNegative = xpp.getAttributeValue(0) == "1"
                        } else if (tnm == "bubble3D") {    // child element only contains 1 element
                            if (xpp.attributeCount > 0)
                                bubble3D = xpp.getAttributeValue(0) == "1"
                        } else if (tnm == "explosion") {    // child element only contains 1 element
                            if (xpp.attributeCount > 0)
                                explosion = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "spPr") {
                            lastTag.push(tnm)
                            sp = SpPr.parseOOXML(xpp, lastTag, bk) as SpPr
                            //		            	 sp.setNS("c");
                        } else if (tnm == "marker") {
                            lastTag.push(tnm)
                            m = Marker.parseOOXML(xpp, lastTag, bk) as Marker
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "dPt") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("dPt.parseOOXML: $e")
            }

            return DPt(idx, invertIfNegative, bubble3D, m, sp, explosion)
        }
    }
}


/*
 * generates the OOXML necessary to represent the data points of the chart 
 * includes fill and line color
 * @return
private String getDataPointOOXML(ChartSeriesHandle[] sh) {
	StringBuffer ooxml= new StringBuffer();
	for (int i= 0; i < sh.length; i++) {
		ooxml.append("<c:dPt>");				ooxml.append("\r\n");
		ooxml.append("<c:idx val=\"" + i + "\"/>");
		if (sh[i].getSpPr()!=null) ooxml.append(sh[i].getSpPr().getOOXML("c"));
		ooxml.append("</c:dPt>");				ooxml.append("\r\n");
	}
	return ooxml.toString();
}
*/