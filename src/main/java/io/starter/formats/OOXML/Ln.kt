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

import java.util.HashMap
import java.util.Stack

/**
 * ln (Outline)
 *
 *
 * This OOXML/DrawingML element specifies an outline style that can be applied to a number of different objects such as shapes and
 * text. The line allows for the specifying of many different types of outlines including even line dashes and bevels
 *
 *
 * parent:  many
 * children: bevel, custDash, gradFill, headEnd, miter, noFill, pattFill, prstDash, round, solidFill, tailEnd
 * attributes:  w (line width) cap (line ending cap) cmpd (compound line type) algn (stroke alignment)
 */
// TODO: Finish custDash
class Ln : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var fill: FillGroup? = null
    private var join: JoinGroup? = null
    private var dash: DashGroup? = null
    private var h: HeadEnd? = null
    private var t: TailEnd? = null


    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:ln")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (fill != null) ooxml.append(fill!!.ooxml)
            if (dash != null) ooxml.append(dash!!.ooxml)
            if (join != null) ooxml.append(join!!.ooxml)
            if (h != null) ooxml.append(h!!.ooxml)
            if (t != null) ooxml.append(t!!.ooxml)
            ooxml.append("</a:ln>")
            return ooxml.toString()
        }

    /**
     * return line width in 2003-style units, 0 for none
     *
     * @return
     */
    /**
     * set the line width
     *
     * @param w
     */
    // Specifies the width to be used for the underline stroke.
    // If this attribute is omitted, then a value of 0 is assumed.
    // 1 pt = 12700 EMUs.
    // default
    var width: Int
        get() = if (attrs != null && attrs!!["w"] != null) Integer.valueOf(attrs!!["w"]) / 12700 else 0
        set(w) {
            if (attrs == null)
                attrs = HashMap()
            attrs!!["w"] = w.toString()

        }

    val color: Int
        get() = if (fill == null) -1 else fill!!.color

    /**
     * convert 2007 preset dasing scheme to 2003 line stype
     * <br></br>0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none,
     * 6= dk gray pattern, 7= med. gray, 8= light gray
     *
     * @return
     */
    // solid
    // ? none = 5 ???
    // solid
    val lineStyle: Int
        get() {
            if (dash != null) {
                val style = dash!!.presetDashingScheme
                if (style == "solid")
                    return 0
                else if (style == "dash" ||
                        style == "sysDash" ||
                        style == "lgDash")
                    return 1
                else if (style == "sysDot")
                    return 2
                else if (style == "dashDot" ||
                        style == "sysDashDot" ||
                        style == "lgDashDot")
                    return 3
                else if (style == "sysDashDashDot" || style == "lgDashDotDot")
                    return 4
            }
            return 0
        }

    /**
     * create a new solid line in width w with color clr
     *
     * @param w   int specifies the width of a line in EMUs. 1 pt = 12700 EMUs
     * @param clr String hex color without the #
     */
    constructor(w: Int, clr: String) {
        this.fill = FillGroup(null, null, null, null, SolidFill(clr))
        this.dash = DashGroup(PrstDash("solid"))
        this.width = w
    }

    constructor() {    // no-param constructor, set up common defaults
        this.fill = FillGroup(null, null, null, null, SolidFill())
        this.join = JoinGroup("800000", true, false, false)
        this.h = HeadEnd()
        this.t = TailEnd()
    }

    constructor(attrs: HashMap<String, String>, fill: FillGroup, join: JoinGroup, dash: DashGroup, h: HeadEnd, t: TailEnd) {
        this.attrs = attrs
        this.fill = fill
        this.join = join
        this.dash = dash
        this.t = t
        this.h = h
    }

    constructor(l: Ln) {
        this.attrs = l.attrs
        this.fill = l.fill
        this.join = l.join
        this.dash = l.dash
        this.t = l.t
        this.h = l.h
    }

    override fun cloneElement(): OOXMLElement {
        return Ln(this)
    }

    fun setColor(clr: String) {
        if (fill == null)
            this.fill = FillGroup(null, null, null, null, SolidFill())
        fill!!.setColor(clr)
    }

    companion object {

        private val serialVersionUID = -161619607936083688L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            val attrs = HashMap<String, String>()
            var fill: FillGroup? = null
            var join: JoinGroup? = null
            var dash: DashGroup? = null
            var h: HeadEnd? = null
            var t: TailEnd? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "ln") {        // get ln attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "noFill" ||
                                tnm == "solidFill" ||
                                tnm == "pattFill" ||
                                tnm == "gradFill") {
                            lastTag.push(tnm)
                            fill = FillGroup.parseOOXML(xpp, lastTag, bk) as FillGroup
                        } else if (tnm == "bevel" ||
                                tnm == "round" ||
                                tnm == "miter") {
                            lastTag.push(tnm)
                            join = JoinGroup.parseOOXML(xpp, lastTag)
                        } else if (tnm == "prstDash" || tnm == "custDash") {
                            lastTag.push(tnm)
                            dash = DashGroup.parseOOXML(xpp, lastTag)
                        } else if (tnm == "headEnd") {
                            lastTag.push(tnm)
                            h = HeadEnd.parseOOXML(xpp, lastTag)
                        } else if (tnm == "tailEnd") {
                            lastTag.push(tnm)
                            t = TailEnd.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "ln") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("ln.parseOOXML: $e")
            }

            return Ln(attrs, fill, join, dash, h, t)
        }
    }
}


/**
 * line dash group:  prstDash, custDash
 */
// TODO: Finish custDash
internal class DashGroup : OOXMLElement {
    private var prstDash: PrstDash? = null

    override// TODO: finish custDash
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (prstDash != null) ooxml.append(prstDash!!.ooxml)
            return ooxml.toString()
        }

    /**
     * returns the preset dashing scheme for the line, if any
     * <br></br>One of:
     * <br></br>dash, dashDot, lgDash, lgDashDot, lgDashDotDot, solid, sysDash, sysDashDot,
     * sysDashDotDot, sysDot
     *
     * @return
     */
    val presetDashingScheme: String?
        get() = if (prstDash != null) prstDash!!.presetDashingScheme else null

    constructor(p: PrstDash) {
        this.prstDash = p
    }

    constructor(d: DashGroup) {
        this.prstDash = d.prstDash
    }

    override fun cloneElement(): OOXMLElement {
        return DashGroup(this)
    }

    companion object {

        private val serialVersionUID = -6892326040716070609L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): DashGroup {
            var p: PrstDash? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (/*tnm.equals("custDash") ||*/
                                tnm == "prstDash") {
                            p = PrstDash.parseOOXML(xpp, lastTag)
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "custDash" || endTag == "prstDash") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("custDash.parseOOXML: $e")
            }

            return DashGroup(p)
        }
    }
}

/**
 * choice of miter, bevel or round
 *
 *
 * parent: ln
 */
/* since each child element is so simple, we will just store which element it is rather than create separate object */
internal class JoinGroup : OOXMLElement {
    private var miter: Boolean = false
    private var round: Boolean = false
    private var bevel: Boolean = false
    private var miterVal: String? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (miter) {
                ooxml.append("<a:miter lim=\"$miterVal\"/>")
            } else if (round) {
                ooxml.append("<a:round/>")
            } else if (bevel) {
                ooxml.append("<a:bevel/>")
            }
            return ooxml.toString()
        }

    constructor(a: String, m: Boolean, r: Boolean, b: Boolean) {
        this.miterVal = a
        this.miter = m
        this.round = r
        this.bevel = b
    }

    constructor(j: JoinGroup) {
        this.miterVal = j.miterVal
        this.miter = j.miter
        this.round = j.round
        this.bevel = j.bevel
    }

    override fun cloneElement(): OOXMLElement {
        return JoinGroup(this)
    }

    companion object {

        private val serialVersionUID = -6107424300366896696L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): JoinGroup {
            var miter = false
            var round = false
            var bevel = false
            var a: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "bevel") {
                            bevel = true
                        } else if (tnm == "round") {
                            round = true
                        } else if (tnm == "miter") {
                            miter = true
                            a = xpp.getAttributeValue(0)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "bevel" ||
                                endTag == "round" ||
                                endTag == "miter") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("JoinGroup.parseOOXML: $e")
            }

            return JoinGroup(a, miter, round, bevel)
        }
    }
}

/**
 * headEnd (Line Head/End Style)
 *
 *
 * This element specifies decorations which can be added to the head of a line.
 */
internal class HeadEnd : OOXMLElement {
    private var len: String? = null
    private var type: String? = null
    private var w: String? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:headEnd")
            if (len != null) ooxml.append(" len=\"$len\"")
            if (type != null) ooxml.append(" type=\"$type\"")
            if (w != null) ooxml.append(" w=\"$w\"")
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor() {

    }

    constructor(len: String, type: String, w: String) {
        this.len = len
        this.type = type
        this.w = w
    }

    constructor(te: HeadEnd) {
        this.len = te.len
        this.type = te.type
        this.w = te.w
    }

    override fun cloneElement(): OOXMLElement {
        return HeadEnd(this)
    }

    companion object {
        private val serialVersionUID = -6744308104003922477L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): HeadEnd {
            var len: String? = null
            var type: String? = null
            var w: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "headEnd") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                if (n == "len")
                                    len = xpp.getAttributeValue(i)
                                else if (n == "type")
                                    type = xpp.getAttributeValue(i)
                                else if (n == "w")
                                    w = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "headEnd") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("headEnd.parseOOXML: $e")
            }

            return HeadEnd(len, type, w)
        }
    }
}

/**
 * tailEnd (Tail line end style)
 *
 *
 * This element specifies decorations which can be added to the tail of a line.
 */
internal class TailEnd : OOXMLElement {
    private var len: String? = null
    private var type: String? = null
    private var w: String? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:tailEnd")
            if (len != null) ooxml.append(" len=\"$len\"")
            if (type != null) ooxml.append(" type=\"$type\"")
            if (w != null) ooxml.append(" w=\"$w\"")
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(len: String, type: String, w: String) {
        this.len = len
        this.type = type
        this.w = w
    }

    constructor(te: TailEnd) {
        this.len = te.len
        this.type = te.type
        this.w = te.w
    }

    override fun cloneElement(): OOXMLElement {
        return TailEnd(this)
    }

    companion object {
        private val serialVersionUID = -5587427916156543370L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): TailEnd {
            var len: String? = null
            var type: String? = null
            var w: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "tailEnd") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                if (n == "len")
                                    len = xpp.getAttributeValue(i)
                                else if (n == "type")
                                    type = xpp.getAttributeValue(i)
                                else if (n == "w")
                                    w = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "tailEnd") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("tailEnd.parseOOXML: $e")
            }

            return TailEnd(len, type, w)
        }
    }
}


internal class PrstDash : OOXMLElement {
    /**
     * returns the preset dashing scheme  for the line, if any
     * <br></br>One of:
     * <br></br>dash, dashDot, lgDash, lgDashDot, lgDashDotDot, solid, sysDash, sysDashDot,
     * sysDashDotDot, sysDot
     *
     * @return
     */
    var presetDashingScheme: String? = null
        private set

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:prstDash val=\"$presetDashingScheme\"/>")
            return ooxml.toString()
        }

    constructor() {}

    constructor(`val`: String) {
        this.presetDashingScheme = `val`
    }

    constructor(p: PrstDash) {
        this.presetDashingScheme = p.presetDashingScheme
    }

    override fun cloneElement(): OOXMLElement {
        return PrstDash(this)
    }

    companion object {

        private val serialVersionUID = -4645986946936173151L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): PrstDash {
            var `val`: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "prstDash") {        // get val attribute
                            `val` = xpp.getAttributeValue(0)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "prstDash") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("prstDash.parseOOXML: $e")
            }

            return PrstDash(`val`)
        }
    }
}
