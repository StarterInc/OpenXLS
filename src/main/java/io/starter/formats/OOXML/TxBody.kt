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
 * txBody (Shape Text Body)
 *
 *
 * OOXML/DrawingML element specifies the existence of text to be contained within the corresponding shape. All visible text and
 * visible text related properties are contained within this element. There can be multiple paragraphs and within
 * paragraphs multiple runs of text
 *
 *
 * parent:    sp (shape)
 * children:  bodyPr REQ, lstStyle, p REQ
 */
// TODO: handle lstStyle Text List Styles
class TxBody : OOXMLElement {
    private var bPr: BodyPr? = null
    private var para: P? = null

    override// TODO: HANDLE
    val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<xdr:txBody>")
            if (bPr != null) tooxml.append(bPr!!.ooxml)
            tooxml.append("<a:lstStyle/>")
            if (para != null) tooxml.append(para!!.ooxml)
            tooxml.append("</xdr:txBody>")
            return tooxml.toString()
        }

    constructor(b: BodyPr, para: P) {
        this.bPr = b
        this.para = para
    }

    constructor(tbd: TxBody) {
        this.bPr = tbd.bPr
        this.para = tbd.para
    }

    override fun cloneElement(): OOXMLElement {
        return TxBody(this)
    }

    companion object {

        private val serialVersionUID = 2407194628070113668L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var b: BodyPr? = null
            var para: P? = null
            try {        // need: endParaRPr?
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "bodyPr") {        // default text properties
                            lastTag.push(tnm)
                            b = BodyPr.parseOOXML(xpp, lastTag) as BodyPr
                        } else if (tnm == "p") {    // part of p element
                            lastTag.push(tnm)
                            para = P.parseOOXML(xpp, lastTag, bk) as P
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "txBody") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("txBody.parseOOXML: $e")
            }

            return TxBody(b, para)
        }
    }
}
