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
import java.util.HashMap
import java.util.Stack

/**
 * Geometry Group: one of prstGeom (Preset shapes) or custGeom (custom shapes)
 */
//TODO: Finish custGeom child elements ahLst, cxnLst, rect (+ finish path element)
class GeomGroup : OOXMLElement {
    private var p: PrstGeom? = null
    private var c: CustGeom? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (this.p != null)
                ooxml.append(p!!.ooxml)
            else if (this.c != null) ooxml.append(c!!.ooxml)
            return ooxml.toString()
        }

    constructor() {}

    constructor(p: PrstGeom, c: CustGeom) {
        this.p = p
        this.c = c
    }

    constructor(g: GeomGroup) {
        this.p = g.p
        this.c = g.c
    }

    override fun cloneElement(): OOXMLElement {
        return GeomGroup(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7202561792070909825L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            var p: PrstGeom? = null
            var c: CustGeom? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "prstGeom") {
                            lastTag.push(tnm)
                            p = PrstGeom.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        } else if (tnm == "custGeom") {
                            lastTag.push(tnm)
                            c = CustGeom.parseOOXML(xpp, lastTag)
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
                Logger.logErr("GeomGroup.parseOOXML: $e")
            }

            return GeomGroup(p, c)
        }
    }

}

internal class PrstGeom : OOXMLElement {
    private var prst: String? = null
    private var avLst: AvLst? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:prstGeom prst=\"$prst\">")
            if (avLst != null) ooxml.append(avLst!!.ooxml)
            ooxml.append("</a:prstGeom>")
            return ooxml.toString()
        }

    constructor() {    // no-param constructor, set up common defaults
        prst = "rect"
        avLst = AvLst()
    }

    constructor(prst: String, a: AvLst) {
        this.prst = prst
        this.avLst = a
    }

    constructor(p: PrstGeom) {
        this.prst = p.prst
        this.avLst = p.avLst
    }

    override fun cloneElement(): OOXMLElement {
        return PrstGeom(this)
    }

    companion object {

        private val serialVersionUID = 8327708502983472577L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): PrstGeom {
            var prst: String? = null
            var a: AvLst? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "prstGeom") {        // get attributes
                            prst = xpp.getAttributeValue(0)
                        } else if (tnm == "avLst") {
                            lastTag.push(tnm)
                            a = AvLst.parseOOXML(xpp, lastTag) as AvLst
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "prstGeom") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("prstGeom.parseOOXML: $e")
            }

            return PrstGeom(prst, a)
        }
    }

}


/**
 * custGeom (Custom Geometry)
 * This element specifies the existence of a custom geometric shape. This shape will consist of a series of lines and
 * curves described within a creation path. In addition to this there may also be adjust values, guides, adjust
 * handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
 *
 *
 * parent:  soPr
 * children:  avLst, gdLst, ahLst, cxnLst, rect, pathLst (REQ)
 */
// TODO: Finish child elements ahLst, cxnLst
internal class CustGeom : OOXMLElement {
    private var pathLst: PathLst? = null
    private var gdLst: GdLst? = null
    private var avLst: AvLst? = null
    private var cxnLst: CxnLst? = null
    private var rect: Rect? = null

    override// avLst
    // gdLst
    // TODO: ahLst
    // cxnLst
    // rect
    // pathLst
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:custGeom>")
            if (avLst != null) ooxml.append(avLst!!.ooxml)
            if (gdLst != null) ooxml.append(gdLst!!.ooxml)
            ooxml.append("<a:ahLst/>")
            if (cxnLst != null) ooxml.append(cxnLst!!.ooxml)
            if (rect != null) ooxml.append(rect!!.ooxml)
            if (pathLst != null) ooxml.append(pathLst!!.ooxml)
            ooxml.append("</a:custGeom>")
            return ooxml.toString()
        }

    constructor(p: PathLst, g: GdLst, a: AvLst, cx: CxnLst, r: Rect) {
        this.pathLst = p
        this.gdLst = g
        this.avLst = a
        this.cxnLst = cx
        this.rect = r
    }

    constructor(c: CustGeom) {
        this.pathLst = c.pathLst
        this.gdLst = c.gdLst
        this.avLst = c.avLst
        this.cxnLst = c.cxnLst
        this.rect = c.rect
    }

    override fun cloneElement(): OOXMLElement {
        return CustGeom(this)
    }

    companion object {

        private val serialVersionUID = 4036207867619551810L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): CustGeom {
            var p: PathLst? = null
            var g: GdLst? = null
            var a: AvLst? = null
            var cx: CxnLst? = null
            var r: Rect? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pathLst") {        // REQ
                            lastTag.push(tnm)
                            p = PathLst.parseOOXML(xpp, lastTag)
                        } else if (tnm == "gdLst") {
                            lastTag.push(tnm)
                            g = GdLst.parseOOXML(xpp, lastTag)
                        } else if (tnm == "avLst") {
                            lastTag.push(tnm)
                            a = AvLst.parseOOXML(xpp, lastTag) as AvLst
                        } else if (tnm == "cxnLst") {
                            lastTag.push(tnm)
                            cx = CxnLst.parseOOXML(xpp, lastTag)
                        } else if (tnm == "rect") {
                            lastTag.push(tnm)
                            r = Rect.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "custGeom") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("custGeom.parseOOXML: $e")
            }

            return CustGeom(p, g, a, cx, r)
        }
    }
}

/**
 * pathLst (List of Shape Paths)
 * This element specifies the entire path that is to make up a single geometric shape. The pathLst can consist of
 * many individual paths within it.
 *
 *
 * parent:  custGeom
 * children path (multiple)
 */
internal class PathLst : OOXMLElement {
    private var path: ArrayList<Path>? = null
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:pathLst")
            val ii = attrs!!.keys.iterator()
            while (ii.hasNext()) {
                val key = ii.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (path != null) {
                for (i in path!!.indices)
                    ooxml.append(path!![i].ooxml)
            }
            ooxml.append("</a:pathLst>")
            return ooxml.toString()
        }


    constructor(attrs: HashMap<String, String>, p: ArrayList<Path>) {
        this.attrs = attrs
        this.path = p
    }

    constructor(pl: PathLst) {
        this.attrs = pl.attrs
        this.path = pl.path
    }

    override fun cloneElement(): OOXMLElement {
        return PathLst(this)
    }

    companion object {

        private val serialVersionUID = -1996347204024728000L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): PathLst {
            val attrs = HashMap<String, String>()
            val p = ArrayList<Path>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pathLst") {        // get attributes
                            for (i in 0 until xpp.attributeCount)
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                        } else if (tnm == "path") {        // get one or more path children
                            lastTag.push(tnm)
                            p.add(Path.parseOOXML(xpp, lastTag))

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pathLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("pathLst.parseOOXML: $e")
            }

            return PathLst(attrs, p)
        }
    }
}

/**
 *
 */
internal class Path : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:path")
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

    constructor(p: Path) {
        this.attrs = p.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Path(this)
    }

    companion object {

        private val serialVersionUID = 6906237439620322589L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): Path {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "path") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "moveTo") {
                        } else if (tnm == "lnTo") {
                        } else if (tnm == "close") {
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "path") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("path.parseOOXML: $e")
            }

            return Path(attrs)
        }
    }
}

/**
 * gdLst (List of Shape Guides)
 * This element specifies all the guides that will be used for this shape. A guide is specified by the gd element and
 * defines a calculated value that may be used for the construction of the corresponding shape.
 */
internal class GdLst : OOXMLElement {
    private var gds: ArrayList<Gd>? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:gdLst>")
            if (gds != null) {
                for (i in gds!!.indices)
                    ooxml.append(gds!![i].ooxml)
            }
            ooxml.append("</a:gdLst>")
            return ooxml.toString()
        }

    constructor(gds: ArrayList<Gd>) {
        this.gds = gds
    }

    constructor(g: GdLst) {
        this.gds = g.gds
    }

    override fun cloneElement(): OOXMLElement {
        return GdLst(this)
    }

    companion object {

        private val serialVersionUID = -7852193131141462744L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): GdLst {
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
                        if (endTag == "gdLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("gdLst.parseOOXML: $e")
            }

            return GdLst(gds)
        }
    }
}

/**
 * ahLst (List of Shape Adjust Handles)
 * This element specifies the adjust handles that will be applied to a custom geometry. These adjust handles will
 * specify points within the geometric shape that can be used to perform certain transform operations on the
 * shape.
 * [Example: Consider the scenario where a custom geometry, an arrow in this case, has been drawn and adjust
 * handles have been placed at the top left corner of both the arrow head and arrow body. The user interface can
 * then be made to transform only certain parts of the shape by using the corresponding adjust handle
 *
 *
 * parent: custGeom
 * children:  one or more of [ahXY, ahPolar (both REQ)]
 */


/**
 * rect (Shape Text Rectangle)
 *
 * This element specifies the rectangular bounding box for text within a custGeom shape. The default for this
 * rectangle is the bounding box for the shape. This can be modified using this elements four attributes to inset or
 * extend the text bounding box.
 *
 * parent: custGeom
 * children: none
 */
internal class Rect : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:rect")
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

    constructor(r: Rect) {
        this.attrs = r.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Rect(this)
    }

    companion object {

        private val serialVersionUID = 2790708601254975676L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): Rect {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "rect") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "CHILDELEMENT") {
                            lastTag.push(tnm)
                            //layout = (layout) layout.parseOOXML(xpp, lastTag).clone();

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "rect") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("rect.parseOOXML: $e")
            }

            return Rect(attrs)
        }
    }
}

/**
 * cxnLst (List of Shape Connection Sites)
 * This element specifies all the connection sites that will be used for this shape. A connection site is specified by
 * defining a point within the shape bounding box that can have a cxnSp element attached to it. These connection
 * sites are specified using the shape coordinate system that is specified within the ext transform element.
 *
 * parents:		custGeom
 * children:	cxn	(0 or more)
 */
internal class CxnLst : OOXMLElement {
    private var cxns: ArrayList<Cxn>? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:cxnLst>")
            if (cxns != null) {
                for (i in cxns!!.indices)
                    ooxml.append(cxns!![i].ooxml)
            }
            ooxml.append("</a:cxnLst>")
            return ooxml.toString()
        }

    constructor(cxns: ArrayList<Cxn>) {
        this.cxns = cxns
    }

    constructor(c: CxnLst) {
        this.cxns = c.cxns
    }

    override fun cloneElement(): OOXMLElement {
        return CxnLst(this)
    }

    companion object {

        private val serialVersionUID = -562847539163221621L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): CxnLst {
            var cxns: ArrayList<Cxn>? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cxn") {
                            lastTag.push(tnm)
                            if (cxns == null) cxns = ArrayList()
                            cxns.add(Cxn.parseOOXML(xpp, lastTag))

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cxnLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cxnLst.parseOOXML: $e")
            }

            return CxnLst(cxns)
        }
    }
}


/**
 * cxn (Shape Connection Site)
 * This element specifies the existence of a connection site on a custom shape. A connection site allows a cxnSp to
 * be attached to this shape. This connection will be maintined when the shape is repositioned within the
 * document. It should be noted that this connection is placed within the shape bounding box using the transform
 * coordinate system which is also called the shape coordinate system, as it encompasses theentire shape. The
 * width and height for this coordinate system are specified within the ext transform element.
 *
 * parents: 	cxnLst
 * children:    pos	REQ
 */
internal class Cxn : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var pos: Pos? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:cxn")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (pos != null) ooxml.append(pos!!.ooxml)
            ooxml.append("</a:cxn>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>, p: Pos) {
        this.attrs = attrs
        this.pos = p
    }

    constructor(c: Cxn) {
        this.attrs = c.attrs
        this.pos = c.pos
    }

    override fun cloneElement(): OOXMLElement {
        return Cxn(this)
    }

    companion object {

        private val serialVersionUID = -4193511102420582252L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): Cxn {
            val attrs = HashMap<String, String>()
            var p: Pos? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cxn") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)    // ang REQ
                            }
                        } else if (tnm == "pos") {
                            lastTag.push(tnm)
                            p = Pos.parseOOXML(xpp, lastTag)

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cxn") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("cxn.parseOOXML: $e")
            }

            return Cxn(attrs, p)
        }
    }
}


/**
 * pos (Shape Position Coordinate)
 * Specifies a position coordinate within the shape bounding box. It should be noted that this coordinate is placed
 * within the shape bounding box using the transform coordinate system which is also called the shape coordinate
 * system, as it encompasses the entire shape. The width and height for this coordinate system are specified within
 * the ext transform element.
 *
 * parents:  cxn, ahPolar, ahXY
 * children: none
 */
internal class Pos : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:pos")
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

    constructor(p: Pos) {
        this.attrs = p.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Pos(this)
    }

    companion object {

        private val serialVersionUID = 5500991309750603125L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): Pos {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pos") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pos") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("pos.parseOOXML: $e")
            }

            return Pos(attrs)
        }
    }
}

