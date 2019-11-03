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

import java.util.ArrayList
import java.util.Stack


/**
 * avLst (List of Shape Adjust Values)
 *
 *
 * This element specifies the adjust values that will be applied to the specified shape. An adjust value is simply a
 * guide that has a value based formula specified. That is, no calculation takes place for an adjust value guide.
 * Instead, this guide specifies a parameter value that is used for calculations within the shape guides.
 *
 *
 * parent:	prstGeom, prstTxWarp, custGeom
 * children: gd (shape guide) (0 or more)
 */
class AvLst : OOXMLElement {
    private var gds: ArrayList<*>? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:avLst>")
            if (gds != null) {
                for (i in gds!!.indices)
                    ooxml.append((gds!![i] as Gd).ooxml)
            }
            ooxml.append("</a:avLst>")
            return ooxml.toString()
        }

    constructor() {

    }

    constructor(gds: ArrayList<*>) {
        this.gds = gds
    }

    constructor(av: AvLst) {
        this.gds = av.gds
    }

    override fun cloneElement(): OOXMLElement {
        return AvLst(this)
    }

    companion object {

        private val serialVersionUID = 4823524943145191780L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            var gds: ArrayList<Gd>? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "gd") {
                            lastTag.push(tnm)
                            if (gds == null) gds = ArrayList()
                            gds.add(Gd.parseOOXML(xpp, lastTag) as Gd)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "avLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("avLst.parseOOXML: $e")
            }

            return AvLst(gds)
        }
    }
}

