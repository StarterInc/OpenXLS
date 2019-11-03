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

import io.starter.formats.XLS.OOXMLAdapter
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * cNvPr  (Non-Visual Drawing Properties)
 *
 *
 * OOXML/DrawingML element specifies non-visual canvas properties. This allows for additional information that does not affect
 * the appearance of the picture to be stored.
 *
 *
 * attributes: descr, hidden, id REQ, name REQ,
 * parents:     nvSpPr, nvPicPr ...
 * children:   hlinkClick, hlinkHover
 */
// TODO: Handle Child elements hlinkClick, hlinkHover
class CNvPr : OOXMLElement {
    //private hlinkClick hc;
    //private hlinkHover hh;
    /**
     * get descr attribute
     *
     * @return
     */

    /**
     * set description attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var descr: String? = null
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
    var name: String? = null
    private var hidden = false
    /**
     * return the id for this element
     *
     * @return
     */
    /**
     * set the id for this element
     *
     * @param id
     */
    var id = -1

    override// TODO: HANDLE  if (hc!=null) tooxml.append(hc.getOOXML());
    // TODO: HANDLE  if (hh!=null) tooxml.append(hh.getOOXML());
    val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<xdr:cNvPr id=\"" + id + "\" name=\"" + OOXMLAdapter.stripNonAscii(name) + "\"")
            if (descr != null) tooxml.append(" descr=\"$descr\"")
            if (hidden) tooxml.append(" hidden=\"" + (if (hidden) "1" else "0") + "\"")
            tooxml.append(">")
            tooxml.append("</xdr:cNvPr>")
            return tooxml.toString()
        }

    constructor() {}

    constructor(/*hlinkClick hc, hlinkHover hh, */id: Int, name: String, descr: String, hidden: Boolean) {
        //this.hc= hc;
        //this.hh= hh;
        this.id = id
        this.name = name
        this.descr = descr
        this.hidden = hidden
    }

    constructor(cnv: CNvPr) {
        //this.hc= cnv.hc;
        //this.hh= cnv.hh;
        this.id = cnv.id
        this.name = cnv.name
        this.descr = cnv.descr
        this.hidden = cnv.hidden
    }

    override fun cloneElement(): OOXMLElement {
        return CNvPr(this)
    }

    companion object {

        private val serialVersionUID = -3382139449400844949L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            //hlinkClick hc= null;
            //hlinkHover hh= null;
            var descr: String? = null
            var name: String? = null
            var hidden = false
            var id = -1
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvPr") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                val `val` = xpp.getAttributeValue(i)
                                if (nm == "id") {
                                    id = Integer.valueOf(`val`).toInt()
                                } else if (nm == "name") {
                                    name = `val`
                                } else if (nm == "descr") {
                                    descr = `val`
                                } else if (nm == "hidden") {
                                    hidden = `val` == "1"
                                }
                            }
                        } else if (tnm == "hlinkClick") {
                            //            	hc= (hlinkClick) hlinkClick.parseOOXML(xpp).clone();
                        } else if (tnm == "hlinkHover") {
                            //				hh= (hlinkHover) hlinkHover.parseOOXML(xpp).clone();
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cNvPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cNvPr.parseOOXML: $e")
            }

            return CNvPr(/*hc, hh, */id, name, descr, hidden)
        }
    }
}

