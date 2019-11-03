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
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.FormatConstants
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack


/**
 * txpr (Text Properties)
 *
 *
 * OOXML/DrawingML element specifies text formatting. The lstStyle element is not supported
 *
 *
 * parent:  axes, title, labels ..
 * children: bodyPr, lstStyle, p
 */
// TODO: Handle lstStyle
class TxPr : OOXMLElement {
    // TODO: handle lstStyle Text List Styles
    private var bPr: BodyPr? = null
    private var para: P? = null

    override// TODO: HANDLE
    val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<c:txPr>")
            if (bPr != null) tooxml.append(bPr!!.ooxml)
            tooxml.append("<a:lstStyle/>")
            if (para != null) tooxml.append(para!!.ooxml)
            tooxml.append("</c:txPr>")
            return tooxml.toString()
        }

    constructor(b: BodyPr, para: P) {
        this.bPr = b
        this.para = para
    }

    constructor(tpr: TxPr) {
        this.bPr = tpr.bPr
        this.para = tpr.para
    }

    override fun cloneElement(): OOXMLElement {
        return TxPr(this)
    }

    /**
     * specify text formatting properties
     *
     * @param fontFace font face
     * @param sz       size in 100 pts (e.g. font size 12.5 pts,, sz= 1250)
     * @param b        true if bold
     * @param i        true if italic
     * @param u        underline:  dash, dashHeavy, dashLong, dashLongHeavy, dbl, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, dotted
     * dottedHeavy, heavy, none, sng, wavy, wavyDbl, wavyHeavy, words (underline only words not spaces)
     * @param strike   one of: dblStrike, noStrike or sngStrike  or null if none
     * @param clr      fill color in hex form without the #
     */
    constructor(fontFace: String, sz: Int, b: Boolean, i: Boolean, u: String, strike: String, clr: String) {
        this.para = P(fontFace, sz, b, i, u, strike, clr)
    }

    constructor(fx: Font, hrot: Int, vrot: String) {
        /*
			 * Specifies the rotation that is being applied to the text within the bounding box. If it not
specified, the rotation of the accompanying shape is used. If it is specified, then this is
applied independently from the shape. That is the shape can have a rotation applied in
addition to the text itself having a rotation applied to it. If this attribute is omitted, then a
value of 0, is implied.

represents an angle in 60,000ths of a degree
5400000=90 degrees
-5400000=90 degrees counter-clockwise
			 */
        val u = fx.underlineStyle
        var usty = "none"
        when (u) {
            FormatConstants.STYLE_UNDERLINE_SINGLE -> usty = "sng"
            FormatConstants.STYLE_UNDERLINE_DOUBLE -> usty = "dbl"
            FormatConstants.STYLE_UNDERLINE_SINGLE_ACCTG -> usty = "sng"
            FormatConstants.STYLE_UNDERLINE_DOUBLE_ACCTG -> usty = "dbl"
        }
        val strike = if (fx.stricken) "sngStrike" else "noStrike"
        bPr = BodyPr(hrot, "horz")
        this.para = P(fx.fontName, fx.fontHeightInPoints.toInt() * 100, fx.bold, fx.italic, usty, strike,
                fx.colorAsOOXMLRBG)
        /*
         * <c:txPr><a:bodyPr rot="-5400000" vert="horz"/>
         * <a:lstStyle/>
         * <a:p><a:pPr><a:defRPr sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0">
         * <a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial Narrow"/><a:ea typeface="Arial Narrow"/><a:cs typeface="Arial Narrow"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
         */
    }

    companion object {

        private val serialVersionUID = -4293247897525807479L

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
                        if (endTag == "txPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("txPr.parseOOXML: $e")
            }

            return TxPr(b, para)
        }
    }
}
