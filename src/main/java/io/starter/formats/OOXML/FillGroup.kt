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

import java.util.ArrayList
import java.util.HashMap
import java.util.Stack

/**
 * FillGroup -- holds ONE OF the Fill-type OOXML/DrawingML Elements: blipFill
 * gradFill grpFill pattFill solidFill noFill
 *
 *
 * parents: bgPr, spPr, fmtScheme, endParaRPr, rPr ...
 */
// TODO: Handle blip element children
// TODO: Handle gradFill element shade properties
class FillGroup : OOXMLElement {
    private var bf: BlipFill? = null
    private var gpf: GradFill? = null
    private var grpf: GrpFill? = null
    private var pf: PattFill? = null
    private var sf: SolidFill? = null

    override// CHOICE OF fill
    // no fill
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (sf != null)
                ooxml.append(sf!!.ooxml)
            else if (gpf != null)
                ooxml.append(gpf!!.ooxml)
            else if (grpf != null)
                ooxml.append(grpf!!.ooxml)
            else if (bf != null)
                ooxml.append(bf!!.ooxml)
            else if (pf != null)
                ooxml.append(pf!!.ooxml)
            else
                ooxml.append("<a:noFill/>")
            return ooxml.toString()
        }

    /**
     * return the id for the embedded picture (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed attribute for this blip (the id for the embedded picture)
     *
     * @param embed
     */
    var embed: String?
        get() = if (bf != null) bf!!.embed else null
        set(embed) {
            if (bf == null)
                bf = BlipFill()
            bf!!.embed = embed
        }

    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (bf != null) bf!!.link else null
        set(link) {
            if (bf == null)
                bf = BlipFill()
            bf!!.link = link
        }

    val color: Int
        get() = if (sf != null) sf!!.getColor() else -1

    constructor(bf: BlipFill, gpf: GradFill, grpf: GrpFill, pf: PattFill,
                sf: SolidFill) {
        this.bf = bf
        this.gpf = gpf
        this.grpf = grpf
        this.pf = pf
        this.sf = sf
    }

    constructor(f: FillGroup) {
        this.bf = f.bf
        this.gpf = f.gpf
        this.grpf = f.grpf
        this.pf = f.pf
        this.sf = f.sf
    }

    override fun cloneElement(): OOXMLElement {
        return FillGroup(this)
    }

    /**
     * set the color for this fill
     * NOTE: at this time only solid fills may be specified
     *
     * @param clr
     */
    fun setColor(clr: String) {
        if (bf != null)
            return
        gpf = null
        grpf = null
        pf = null
        sf = SolidFill()
        sf!!.setColor(clr)
    }

    companion object {
        private val serialVersionUID = 8320871291479597945L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var bf: BlipFill? = null
            var gpf: GradFill? = null
            var grpf: GrpFill? = null
            var pf: PattFill? = null
            var sf: SolidFill? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "solidFill") {
                            lastTag.push(tnm)
                            sf = SolidFill.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        } else if (tnm == "noFill") {
                            // do nothing
                        } else if (tnm == "gradFill") {
                            lastTag.push(tnm)
                            gpf = GradFill.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        } else if (tnm == "grpFill") {
                            lastTag.push(tnm)
                            grpf = GrpFill.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        } else if (tnm == "pattFill") {
                            lastTag.push(tnm)
                            pf = PattFill.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        } else if (tnm == "blipFill") {
                            lastTag.push(tnm)
                            bf = BlipFill.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) { // shouldn't get here
                        lastTag.pop()
                        break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("FillGroup.parseOOXML: $e")
            }

            return FillGroup(bf, gpf, grpf, pf, sf)
        }
    }
}

/**
 * gradFill (Gradient Fill)
 *
 *
 * This element defines a gradient fill.
 *
 *
 * parents: many children: gsLst, SHADEPROPERTIES, tileRect
 */
// TODO: finish SHADEPROPERTIES
internal class GradFill : OOXMLElement {
    private var g: GsLst? = null
    // private ShadeProps sp= null;
    private var tr: TileRect? = null
    private var attrs: HashMap<String, String>? = null

    override// attributes
    // if (sp!=null) ooxml.append(sp.getOOXML());
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:gradFill")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (g != null)
                ooxml.append(g!!.ooxml)
            if (tr != null)
                ooxml.append(tr!!.ooxml)
            ooxml.append("</a:gradFill>")
            return ooxml.toString()
        }

    constructor(g: GsLst, /* ShadeProps sp, */tr: TileRect, attrs: HashMap<String, String>) {
        this.g = g
        // this.sp= sp;
        this.tr = tr
        this.attrs = attrs
    }

    constructor(gf: GradFill) {
        this.g = gf.g
        // this.sp= gf.sp;
        this.tr = gf.tr
        this.attrs = gf.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return GradFill(this)
    }

    companion object {

        private val serialVersionUID = 8965776942160065286L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): GradFill {
            val attrs = HashMap<String, String>()
            var g: GsLst? = null
            // ShadeProps sp= null;
            var tr: TileRect? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "gradFill") { // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                            }
                        } else if (tnm == "tileRect") {
                            lastTag.push(tnm)
                            tr = TileRect.parseOOXML(xpp, lastTag)
                        } else if (tnm == "gsLst") {
                            lastTag.push(tnm)
                            g = GsLst.parseOOXML(xpp, lastTag, bk)
                            /*
                         * } else if (tnm.equals("lin") || tnm.equals("path")) {
                         * lastTag.push(tnm); sp= (shadeProps)
                         * shadeProps.parseOOXML(xpp, lastTag).clone();
                         */
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "gradFill") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("gradFill.parseOOXML: $e")
            }

            return GradFill(g, /* sp, */tr, attrs)
        }
    }

}

/**
 * blipFill (Picture Fill)
 *
 *
 * This element specifies the type of picture fill that the picture object will
 * have. Because a picture has a picture fill already by default, it is possible
 * to have two fills specified for a picture object. An example of this is shown
 * below.
 *
 *
 * parents: many children: blip, srcRect, FillModeProperties
 */
internal class BlipFill : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var blip: Blip? = null
    private var srcRect: SrcRect? = null
    private var fillMode: FillMode? = null
    private var ns: String? = "xdr"    //	default namespace

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (ns == null)
                Logger.logErr("Error: BlipFill Namespace is null")
            ooxml.append("<$ns:blipFill")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append(">")
            if (blip != null)
                ooxml.append(blip!!.ooxml)
            if (srcRect != null)
                ooxml.append(srcRect!!.ooxml)
            if (fillMode != null)
                ooxml.append(fillMode!!.ooxml)
            ooxml.append("</$ns:blipFill>")
            return ooxml.toString()
        }

    /**
     * return the id for the embedded picture (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed attribute for this blip (the id for the embedded picture)
     *
     * @param embed
     */
    var embed: String?
        get() = if (blip != null) blip!!.embed else null
        set(embed) {
            if (blip == null)
                blip = Blip()
            blip!!.embed = embed
        }

    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (blip != null) blip!!.link else null
        set(link) {
            if (blip == null)
                blip = Blip()
            blip!!.link = link
        }

    constructor() {    // no-parameter constructor, set common defaults
        srcRect = SrcRect()
        fillMode = FillMode(null, false)
    }

    constructor(ns: String, attrs: HashMap<String, String>, b: Blip, s: SrcRect, f: FillMode) {
        this.ns = ns
        this.attrs = attrs
        this.blip = b
        this.srcRect = s
        this.fillMode = f
    }

    constructor(bf: BlipFill) {
        this.ns = bf.ns
        this.attrs = bf.attrs
        this.blip = bf.blip
        this.srcRect = bf.srcRect
        this.fillMode = bf.fillMode
    }

    override fun cloneElement(): OOXMLElement {
        return BlipFill(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2030570462677450734L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): BlipFill {
            val attrs = HashMap<String, String>()
            var b: Blip? = null
            var s: SrcRect? = null
            var f: FillMode? = null
            var ns: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "blipFill") { // get attributes
                            ns = xpp.prefix
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                        } else if (tnm == "blip") {
                            lastTag.push(tnm)
                            b = Blip.parseOOXML(xpp, lastTag)
                        } else if (tnm == "srcRect") {
                            lastTag.push(tnm)
                            s = SrcRect.parseOOXML(xpp, lastTag)
                        } else if (tnm == "stretch" || tnm == "tile") {
                            lastTag.push(tnm)
                            f = FillMode.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "blipFill") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("blipFill.parseOOXML: $e")
            }

            return BlipFill(ns, attrs, b, s, f)
        }
    }
}

/**
 * solidFill (Solid Fill) This element specifies a solid color fill. The shape
 * is filled entirely with the specified color
 *
 *
 * parents: many children: COLORCHOICE
 */
internal class SolidFill : OOXMLElement {
    private var color: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:solidFill>")
            if (this.color != null)
                ooxml.append(this.color!!.ooxml)
            ooxml.append("</a:solidFill>")
            return ooxml.toString()
        }

    constructor() { // no-param constructor, set up common defaults
        this.color = ColorChoice(null, SrgbClr("000000", null), null, null, null)
    }

    constructor(c: ColorChoice) {
        this.color = c
    }

    constructor(s: SolidFill) {
        this.color = s.color
    }

    /**
     * creates a solid fill of a specific color
     *
     * @param clr hex color string without #
     */
    constructor(clr: String?) {
        var clr = clr
        if (clr == null) clr = "FFFFFF"
        this.color = ColorChoice(null, SrgbClr(clr, null), null, null, null)
    }

    override fun cloneElement(): OOXMLElement {
        return SolidFill(this)
    }

    /**
     * set the color for this solid fill in hex (html) string format
     *
     * @param clr
     */
    fun setColor(clr: String) {
        this.color = ColorChoice(null, SrgbClr(clr, null), null, null, null)
    }

    fun getColor(): Int {
        return color!!.color
    }

    companion object {

        private val serialVersionUID = 3341509200573989744L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): SolidFill {
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "hslClr" || tnm == "prstClr"
                                || tnm == "schemeClr"
                                || tnm == "scrgbClr" || tnm == "srgbClr"
                                || tnm == "sysClr") { // get attributes
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk)
                                    .cloneElement() as ColorChoice

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "solidFill") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("solidFill.parseOOXML: $e")
            }

            return SolidFill(c)
        }
    }
}

/**
 * pattFill (Pattern Fill) This element specifies a pattern fill. A repeated
 * pattern is used to fill the object.
 *
 *
 * parent: many children: bgClr, fgClr
 */
internal class PattFill : OOXMLElement {
    private var prst: String? = null
    private var bg: BgClr? = null
    private var fg: FgClr? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:pattFill prst=\"$prst\">")
            if (fg != null)
                ooxml.append(fg!!.ooxml)
            if (bg != null)
                ooxml.append(bg!!.ooxml)
            ooxml.append("</a:pattFill>")
            return ooxml.toString()
        }

    constructor(prst: String, bg: BgClr, fg: FgClr) {
        this.prst = prst
        this.bg = bg
        this.fg = fg
    }

    constructor(p: PattFill) {
        this.prst = p.prst
        this.bg = p.bg
        this.fg = p.fg
    }

    override fun cloneElement(): OOXMLElement {
        return PattFill(this)
    }

    companion object {

        private val serialVersionUID = -1052627959661249692L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): PattFill {
            var prst: String? = null
            var bg: BgClr? = null
            var fg: FgClr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pattFill") { // get attributes
                            prst = xpp.getAttributeValue(0)
                        } else if (tnm == "bgClr") {
                            lastTag.push(tnm)
                            bg = BgClr.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "fgClr") {
                            lastTag.push(tnm)
                            fg = FgClr.parseOOXML(xpp, lastTag, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pattFill") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("pattFill.parseOOXML: $e")
            }

            return PattFill(prst, bg, fg)
        }
    }
}

/**
 * grpFill (Group Fill) This element specifies a group fill. When specified,
 * this setting indicates that the parent element is part of a group and should
 * inherit the fill properties of the group.
 *
 *
 * parent: many children: CT_GROUPFILL: ??? contains nothing?
 */
internal class GrpFill : OOXMLElement {

    override// TODO: no attributes or children?
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:grpFill/>")
            return ooxml.toString()
        }

    override fun cloneElement(): OOXMLElement {
        return GrpFill()
    }

    companion object {
        private val serialVersionUID = 2388879629485740996L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): GrpFill {
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "grpFill") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("grpFill.parseOOXML: $e")
            }

            return GrpFill()
        }
    }
}

/**
 * bgClr background color
 *
 *
 * parent: pattFill
 */
internal class BgClr : OOXMLElement {
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:bgClr>")
            if (this.colorChoice != null)
                ooxml.append(this.colorChoice!!.ooxml)
            ooxml.append("</a:bgClr>")
            return ooxml.toString()
        }


    constructor(c: ColorChoice) {
        this.colorChoice = c
    }

    constructor(s: BgClr) {
        this.colorChoice = s.colorChoice
    }

    override fun cloneElement(): OOXMLElement {
        return BgClr(this)
    }

    companion object {

        private val serialVersionUID = -879409152334931909L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): BgClr {
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "hslClr" || tnm == "prstClr"
                                || tnm == "schemeClr"
                                || tnm == "scrgbClr" || tnm == "srgbClr"
                                || tnm == "sysClr") { // get attributes
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk) as ColorChoice


                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "bgClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("bgClr.parseOOXML: $e")
            }

            return BgClr(c)
        }
    }
}

/**
 * fgClr Foreground Color
 *
 *
 * parent: pattFill
 */
internal class FgClr : OOXMLElement {
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:fgClr>")
            if (this.colorChoice != null)
                ooxml.append(this.colorChoice!!.ooxml)
            ooxml.append("</a:fgClr>")
            return ooxml.toString()
        }

    constructor(c: ColorChoice) {
        this.colorChoice = c
    }

    constructor(s: FgClr) {
        this.colorChoice = s.colorChoice
    }

    override fun cloneElement(): OOXMLElement {
        return FgClr(this)
    }

    companion object {
        private val serialVersionUID = 6836994790529289731L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): FgClr {
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "hslClr" || tnm == "prstClr"
                                || tnm == "schemeClr"
                                || tnm == "scrgbClr" || tnm == "srgbClr"
                                || tnm == "sysClr") { // get attributes
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk) as ColorChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fgClr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("fgClr.parseOOXML: $e")
            }

            return FgClr(c)
        }
    }
}

/**
 * blip (Blip) This element specifies the existence of an image (binary large
 * image or picture) and contains a reference to the image data.
 *
 *
 * parent: blipFill, buFill children: MANY
 */
// TODO: HANDLE THE MANY CHILDREN
internal class Blip : OOXMLElement {
    private var attrs: HashMap<String, String>? = HashMap()

    // TODO: cstate= "print"
    override// attributes
    // TODO: HANDLE CHILDREN
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    /**
     * return the id for the embedded picture (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed attribute for this blip (the id for the embedded picture)
     *
     * @param embed
     */
    var embed: String?
        get() = if (attrs != null && attrs!!["r:embed"] != null) {
            attrs!!["r:embed"]
        } else null
        set(embed) {
            attrs!!["r:embed"] = embed
        }

    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (attrs != null && attrs!!["link"] != null) {
            attrs!!["link"]
        } else null
        set(link) {
            attrs!!["link"] = link
        }

    constructor() {}

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(b: Blip) {
        this.attrs = b.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Blip(this)
    }

    companion object {

        private val serialVersionUID = 5188967633123620513L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): Blip {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "blip") { // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                if (xpp.getAttributePrefix(i) != null)
                                    attrs[xpp.getAttributePrefix(i) + ":"
                                            + xpp.getAttributeName(i)] = xpp
                                            .getAttributeValue(i)
                                else
                                    attrs[xpp.getAttributeName(i)] = xpp
                                            .getAttributeValue(i)
                            }
                            // } else if (tnm.equals("CHILDELEMENT")) {
                            // lastTag.push(tnm);
                            // layout = (layout) layout.parseOOXML(xpp,
                            // lastTag).clone();
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "blip") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("blip.parseOOXML: $e")
            }

            return Blip(attrs)
        }
    }
}

/**
 * tileRect (Tile Rectangle) This element specifies a rectangular region of the
 * shape to which the gradient is applied. This region is then tiled across the
 * remaining area of the shape to complete the fill. The tile rectangle is
 * defined by percentage offsets from the sides of the shape's bounding box.
 * parent: gradFill children: none
 */
internal class TileRect : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:tileRect")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(t: TileRect) {
        this.attrs = t.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return TileRect(this)
    }

    companion object {

        private val serialVersionUID = 5380575948049571420L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): TileRect {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "tileRect") { // get attributes
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "tileRect") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("tileRect.parseOOXML: $e")
            }

            return TileRect(attrs)
        }
    }
}

/**
 * srcRect (Source Rectangle) This element specifies the portion of the blip
 * used for the fill
 *
 *
 * parent: blipFill children: NONE
 */
internal class SrcRect : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:srcRect")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(s: SrcRect) {
        this.attrs = s.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return SrcRect(this)
    }

    companion object {

        private val serialVersionUID = -6407800173040857433L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): SrcRect {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "srcRect") { // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "srcRect") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("srcRect.parseOOXML: $e")
            }

            return SrcRect(attrs)
        }
    }
}

/**
 * fillRect (Fill Rectangle) This element specifies a fill rectangle. When
 * stretching of an image is specified, a source rectangle, srcRect, is scaled
 * to fit the specified fill rectangle. parent: stretch children: NONE
 */
internal class FillRect : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:fillRect")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(f: FillRect) {
        this.attrs = f.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return FillRect(this)
    }

    companion object {

        private val serialVersionUID = 7200764163180402065L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): FillRect {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "fillRect") { // get attributes
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fillRect") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("fillRect.parseOOXML: $e")
            }

            return FillRect()
        }
    }
}

/**
 * FillMode Choice of either tile or stretch elements
 *
 *
 * tile (Tile) This element specifies that a BLIP should be tiled to fill the
 * available space. This element defines a "tile" rectangle within the bounding
 * box. The image is encompassed within the tile rectangle, and the tile
 * rectangle is tiled across the bounding box to fill the entire area. stretch
 * (Stretch) This element specifies that a BLIP should be stretched to fill the
 * target rectangle. The other option is a tile where a BLIP is tiled to fill
 * the available area.
 *
 *
 * parent: blipFill choice of : stretch or tile
 */
internal class FillMode : OOXMLElement {
    // Since both "child" elements are so small, just output OOXML "by hand"
    private var attrs: HashMap<String, String>? = null
    private var tile = false

    override// Since both "child" elements are so small, just output OOXML "by hand"
    // attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (this.tile)
                ooxml.append("<a:tile")
            else
                ooxml.append("<a:stretch><a:fillRect")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            if (this.tile)
                ooxml.append("/>")
            else
                ooxml.append("/></a:stretch>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(attrs: HashMap<String, String>, tile: Boolean) {
        this.attrs = attrs
        this.tile = tile
    }

    constructor(f: FillMode) {
        this.attrs = f.attrs
        this.tile = f.tile
    }

    override fun cloneElement(): OOXMLElement {
        return FillMode(this)
    }

    companion object {

        private val serialVersionUID = 967269629502516244L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): FillMode {
            val attrs = HashMap<String, String>()
            var tile = false
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "tile") {
                            tile = true
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                            }
                        } else if (tnm == "fillRect") { // only child of
                            // stretch
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp
                                        .getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "tile" || endTag == "stretch") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("FillMode.parseOOXML: $e")
            }

            return FillMode(attrs, tile)
        }
    }
}

/**
 * gsLst (Gradient Stop List) The list of gradient stops that specifies the
 * gradient colors and their relative positions in the color band.
 *
 *
 * parent: gradFill children: gs
 */
internal class GsLst : OOXMLElement {
    private var gs: ArrayList<Gs>? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:gsLst")
            if (!gs!!.isEmpty()) {
                ooxml.append(">")
                for (i in gs!!.indices) {
                    ooxml.append(gs!![i].ooxml)
                }
                ooxml.append("</a:gsLst>")
            } else
                ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(g: ArrayList<Gs>) {
        this.gs = g
    }

    constructor(gl: GsLst) {
        this.gs = gl.gs
    }

    override fun cloneElement(): OOXMLElement {
        return GsLst(this)
    }

    companion object {

        private val serialVersionUID = 6576320251327916221L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): GsLst {
            val g = ArrayList<Gs>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "gs") {
                            lastTag.push(tnm)
                            g.add(Gs.parseOOXML(xpp, lastTag, bk))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "gsLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("gsLst.parseOOXML: $e")
            }

            return GsLst(g)
        }
    }
}

/**
 * gs (Gradient stops) This element defines a gradient stop. A gradient stop
 * consists of a position where the stop appears in the color band.
 *
 *
 * parent: gsLst children: COLORCHOICE
 */
internal class Gs : OOXMLElement {
    private var pos: String? = null
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:gs pos=\"" + this.pos + "\">")
            ooxml.append(colorChoice!!.ooxml)
            ooxml.append("</a:gs>")
            return ooxml.toString()
        }

    constructor(pos: String, c: ColorChoice) {
        this.pos = pos
        this.colorChoice = c
    }

    constructor(g: Gs) {
        this.pos = g.pos
        this.colorChoice = g.colorChoice
    }

    override fun cloneElement(): OOXMLElement {
        return Gs(this)
    }

    companion object {

        private val serialVersionUID = 7626866241477598159L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): Gs {
            var pos: String? = null
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "gs") { // get attributes
                            pos = xpp.getAttributeValue(0)
                        } else if (tnm == "hslClr" || tnm == "prstClr"
                                || tnm == "schemeClr"
                                || tnm == "scrgbClr" || tnm == "srgbClr"
                                || tnm == "sysClr") { // get attributes
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk) as ColorChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "gs") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("gs.parseOOXML: $e")
            }

            return Gs(pos, c)
        }
    }

}
