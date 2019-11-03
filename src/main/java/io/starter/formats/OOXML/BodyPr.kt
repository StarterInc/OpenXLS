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
 * bodyPr (Body Properties)
 *
 *
 * This element defines the body properties for the text body within a shape
 *
 *
 * parents:  many, including txBody and txPr
 * attributes:  many
 * children: flatTx, noAutoFit, normAutoFit, prstTxWarp, scene3d, sp3d, spAutoFit
 */
// TODO: Handle CHILDREN ***********************************
class BodyPr : OOXMLElement {
    private var attrs = HashMap<String, String>()
    private var txwarp: PrstTxWarp? = null
    private var spAutoFit = false

    override// attributes
    // TODO: Should be a choice of autofit options
    // scene3d
    // text3d choice
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:bodyPr")
            val i = attrs.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs[key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (txwarp != null) ooxml.append(txwarp!!.ooxml)
            if (spAutoFit) ooxml.append("<a:spAutoFit/>")
            ooxml.append("</a:bodyPr>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(attrs: HashMap<String, String>, txwarp: PrstTxWarp, spAutoFit: Boolean) {
        this.attrs = attrs
        this.txwarp = txwarp
        this.spAutoFit = spAutoFit
    }

    constructor(tpr: BodyPr) {
        this.attrs = tpr.attrs
        this.txwarp = tpr.txwarp
        this.spAutoFit = tpr.spAutoFit
    }

    /**
     * defines the body properties for the text body within a shape
     *
     * @param hrot
     * @param vert Determines if the text within the given text body should be displayed vertically. If this attribute is omitted, then a value of horz, or no vertical text is implied.
     * vert Determines if all of the text is vertical orientation	(each line is 90 degrees rotated clockwise, so it goes from top to bottom; each next line is to the left from the previous one).
     * vert270 Determines if all of the text is vertical orientation (each line is 270 degrees rotated clockwise, so it goes from bottom to top; each next line is to the right from the previous one).
     * wordArtVert Determines if all of the text is vertical ("one letter on top of another").
     * wordArtVertRtl  Specifies that vertical WordArt should be shown from right to left rather than left to right.
     * eaVert  A special version of vertical text, where some fonts are displayed as if rotated by 90 degrees while some fonts (mostly East Asian) are displayed vertical.
     */
    constructor(hrot: Int, vert: String?) {
        attrs = HashMap()
        attrs["rot"] = hrot.toString()
        if (vert != null)
            attrs["vert"] = vert
    }

    override fun cloneElement(): OOXMLElement {
        return BodyPr(this)
    }

    companion object {

        private val serialVersionUID = 3693893834015788452L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            val attrs = HashMap<String, String>()
            var txwarp: PrstTxWarp? = null
            var spAutoFit = false
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "bodyPr") {        // body text properties
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                            // TODO: handle flatTx element
                            // TODO: handle scene3d element
                            // TODO: handle sp3d element
                        } else if (tnm == "spAutoFit") {    // TODO: should be a choice of autofit options
                            spAutoFit = true    // no attributes or children
                        } else if (tnm == "prstTxWarp") {
                            lastTag.push(tnm)
                            txwarp = PrstTxWarp.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "bodyPr") {
                            lastTag.pop()    // pop this tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("bodyPr.parseOOXML: $e")
            }

            return BodyPr(attrs, txwarp, spAutoFit)
        }
    }

}


/**
 * prstTxWarp (Preset Text Warp)
 *
 *
 * This element specifies when a preset geometric shape should be used to transform a piece of text. This
 * operation is known formally as a text warp. The generating application should be able to render all preset
 * geometries enumerated in the ST_TextShapeType list.
 *
 *
 * parent:  bodyPr
 * children: avLst
 */
internal class PrstTxWarp : OOXMLElement {
    private var prst: String? = null
    private var av: AvLst? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:prstTxWarp prst=\"$prst\">")
            if (av != null) ooxml.append(av!!.ooxml)
            ooxml.append("</a:prstTxWarp>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(prst: String, av: AvLst) {
        this.prst = prst
        this.av = av
    }

    constructor(p: PrstTxWarp) {
        this.prst = p.prst
        this.av = p.av
    }

    override fun cloneElement(): OOXMLElement {
        return PrstTxWarp(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2627323317407321668L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): PrstTxWarp {
            var prst: String? = null
            var av: AvLst? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "prstTxWarp") {        // prst is only attribute and is required
                            prst = xpp.getAttributeValue(0)
                        } else if (tnm == "avLst") {
                            lastTag.push(tnm)
                            av = AvLst.parseOOXML(xpp, lastTag) as AvLst

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "prstTxWarp") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("prstTxWarp.parseOOXML: $e")
            }

            return PrstTxWarp(prst, av)
        }
    }
}

