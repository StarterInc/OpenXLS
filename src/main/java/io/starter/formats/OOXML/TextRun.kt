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
 * textRun group, either r (regular text), br (line break) or fld (text Field)
 * parent:  p
 * children: either r, br or fld
 */
//TODO: Finish rPr children highlight TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver,
class TextRun : OOXMLElement {
    private var run: r? = null
    private var brk: Br? = null
    private var f: Fld? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (run != null)
                ooxml.append(run!!.ooxml)
            else if (brk != null)
                ooxml.append(brk!!.ooxml)
            else if (f != null) ooxml.append(f!!.ooxml)
            return ooxml.toString()
        }

    val title: String?
        get() {
            if (run != null) return run!!.title
            return if (f != null) f!!.title else null
        }

    /**
     * return the text properties for this text run
     *
     * @return
     */
    val textProperties: HashMap<String, String>
        get() = if (run != null) run!!.textProperties else HashMap()

    constructor(run: r, brk: Br, f: Fld) {
        this.run = run
        this.brk = brk
        this.f = f
    }

    constructor(r: TextRun) {
        this.run = r.run
        this.brk = r.brk
        this.f = r.f
    }

    /**
     * create a new regular text text run (OOXML element r)
     *
     * @param s
     */
    constructor(s: String) {
        this.run = r(s, null)
    }

    override fun cloneElement(): OOXMLElement {
        return TextRun(this)
    }

    companion object {

        private val serialVersionUID = -6224636879471246452L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var run: r? = null
            var brk: Br? = null
            var f: Fld? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "r") {
                            lastTag.push(tnm)
                            run = r.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        } else if (tnm == "br") {
                            lastTag.push(tnm)
                            brk = Br.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        } else if (tnm == "fld") {
                            lastTag.push(tnm)
                            f = Fld.parseOOXML(xpp, lastTag, bk)
                            lastTag.pop()
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {    // shouldn't get here
                        lastTag.pop()
                        break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("textRun.parseOOXML: $e")
            }

            return TextRun(run, brk, f)
        }
    }
}

/**
 * OOXML element r, text run, sub-element of p (paragraph)
 *
 *
 * children:  rPr, t (actual text string)
 */
internal class r : OOXMLElement {
    var title: String? = ""
        private set        // t element just contains string
    private var rp: RPr? = null

    override// text run
    val ooxml: String
        get() {
            if (title == null || title == "") return ""

            title = io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(title).toString()
            val ooxml = StringBuffer()
            ooxml.append("<a:r>")
            if (rp != null) ooxml.append(rp!!.ooxml)
            ooxml.append("<a:t>$title</a:t>")
            ooxml.append("</a:r>")
            return ooxml.toString()
        }

    val textProperties: HashMap<String, String>
        get() = if (rp != null) rp!!.textProperties else HashMap()

    constructor(title: String, rp: RPr) {
        this.rp = rp
        this.title = title
    }

    constructor(run: r) {
        this.rp = run.rp
        this.title = run.title
    }

    override fun cloneElement(): OOXMLElement {
        return r(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 863254651451294443L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): r {
            var t: String? = ""
            var rp: RPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "t") {        // t element of text run -- the title string we are interested in
                            t = io.starter.formats.XLS.OOXMLAdapter.getNextText(xpp)
                        } else if (tnm == "rPr") {    // text run properties
                            lastTag.push(tnm)
                            rp = RPr.parseOOXML(xpp, lastTag, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "r") {
                            lastTag.pop()    // pop this tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("r.parseOOXML: $e")
            }

            return r(t, rp)
        }
    }
}

/**
 * OOXML element br, vertical break, sub-element of p (paragraph)
 *
 *
 * children:  rPr
 */
internal class Br : OOXMLElement {
    private var rp: RPr? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:br>")
            if (rp != null) ooxml.append(rp!!.ooxml)
            ooxml.append("</a:br>")
            return ooxml.toString()
        }

    constructor(rp: RPr) {
        this.rp = rp
    }

    constructor(b: Br) {
        this.rp = b.rp
    }

    override fun cloneElement(): OOXMLElement {
        return Br(this)
    }

    companion object {
        private val serialVersionUID = -1724086871866480013L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): Br {
            var rp: RPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "rPr") {    // text run properties
                            lastTag.push(tnm)
                            rp = RPr.parseOOXML(xpp, lastTag, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "br") {
                            lastTag.pop()    // pop this tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("br.parseOOXML: $e")
            }

            return Br(rp)
        }
    }
}

/**
 * OOXML element fld, text field, sub-element of p (paragraph)
 *
 *
 * children:  pPr, rPr, t (actual text string)
 */
internal class Fld : OOXMLElement {
    var title: String? = null
        private set        // t element just contains string
    private var rp: RPr? = null
    private var p: PPr? = null
    var id: String
    var type: String? = null

    override// text field
    val ooxml: String
        get() {
            if (title == null || title == "") return ""

            title = io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(title).toString()
            val ooxml = StringBuffer()
            ooxml.append("<a:fld")
            ooxml.append(" id=\"$id\"")
            if (type != null) ooxml.append(" type=\"$type\"")
            ooxml.append(">")
            if (rp != null) ooxml.append(rp!!.ooxml)
            if (p != null) ooxml.append(p!!.ooxml)
            ooxml.append("<a:t>$title</a:t>")
            ooxml.append("</a:fld>")
            return ooxml.toString()
        }

    constructor(id: String, type: String, title: String, rp: RPr, p: PPr) {
        this.id = id
        this.type = type
        this.rp = rp
        this.title = title
        this.p = p
    }

    constructor(f: Fld) {
        this.id = f.id
        this.type = f.type
        this.rp = f.rp
        this.title = f.title
        this.p = f.p
    }

    override fun cloneElement(): OOXMLElement {
        return Fld(this)
    }

    companion object {

        private val serialVersionUID = -7060602732912595402L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): Fld {
            var t: String? = ""
            var id = ""
            var type = ""
            var p: PPr? = null
            var rp: RPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "fld") {
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                if (n == "id") {
                                    id = xpp.getAttributeValue(i)
                                } else if (n == "type") {
                                    type = xpp.getAttributeValue(i)
                                }
                            }
                        } else if (tnm == "t") {        // t element -- the title string we are interested in
                            t = io.starter.formats.XLS.OOXMLAdapter.getNextText(xpp)
                        } else if (tnm == "rPr") {    // text run properties
                            lastTag.push(tnm)
                            rp = RPr.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "pPr") {    // text field properties
                            lastTag.push(tnm)
                            p = PPr.parseOOXML(xpp, lastTag, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fld") {
                            lastTag.pop()    // pop this tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("textRun.parseOOXML: $e")
            }

            return Fld(id, type, t, rp, p)
        }
    }

}

internal class RPr : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var l: Ln? = null
    private var fill: FillGroup? = null
    private var effect: EffectPropsGroup? = null
    private var latin: String? = null
    private var ea: String? = null
    private var cs: String? = null

    override// attributes
    // group fill
    // group effect
    // highlight
    // TEXTUNDERLINELINE
    // TEXTUNDERLINEFILL
    // sym
    // hlinkClick
    // hlinkMouseOver
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:rPr")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append(">")
            if (l != null) ooxml.append(l!!.ooxml)
            if (fill != null) ooxml.append(fill!!.ooxml)
            if (effect != null) ooxml.append(effect!!.ooxml)
            if (latin != null) ooxml.append("<a:latin typeface=\"$latin\"/>")
            if (ea != null) ooxml.append("<a:ea typeface=\"$ea\"/>")
            if (cs != null) ooxml.append("<a:cs typeface=\"$cs\"/>")
            ooxml.append("</a:rPr>")
            return ooxml.toString()
        }

    /**
     * return the text properties for this text run
     *
     * @return
     */
    // TODO: Fill, line ...
    val textProperties: HashMap<String, String>
        get() {
            val textprops = HashMap<String, String>()
            textprops.putAll(attrs!!)
            if (latin != null) textprops["latin_typeface"] = latin
            if (ea != null) textprops["ea_typeface"] = ea
            if (cs != null) textprops["cs_typeface"] = cs
            return textprops
        }

    constructor(attrs: HashMap<String, String>, l: Ln, fill: FillGroup, effect: EffectPropsGroup, latin: String, ea: String, cs: String) {
        this.attrs = attrs
        this.l = l
        this.fill = fill
        this.effect = effect
        this.latin = latin
        this.ea = ea
        this.cs = cs
    }

    constructor(rp: RPr) {
        this.attrs = rp.attrs
        this.l = rp.l
        this.fill = rp.fill
        this.effect = rp.effect
        this.latin = rp.latin
        this.ea = rp.ea
        this.cs = rp.cs
    }

    override fun cloneElement(): OOXMLElement {
        return RPr(this)
    }

    companion object {

        private val serialVersionUID = 228716184734751439L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): RPr {
            val attrs = HashMap<String, String>()
            var l: Ln? = null
            var fill: FillGroup? = null
            var effect: EffectPropsGroup? = null
            var latin: String? = null
            var ea: String? = null
            var cs: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "rPr") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "ln") {
                            lastTag.push(tnm)
                            l = Ln.parseOOXML(xpp, lastTag, bk) as Ln
                        } else if (tnm == "solidFill" ||
                                tnm == "noFill" ||
                                tnm == "gradFill" ||
                                tnm == "grpFill" ||
                                tnm == "pattFill" ||
                                tnm == "blipFill") {
                            lastTag.push(tnm)
                            fill = FillGroup.parseOOXML(xpp, lastTag, bk) as FillGroup
                        } else if (tnm == "effectLst" || tnm == "effectDag") {
                            lastTag.push(tnm)
                            effect = EffectPropsGroup.parseOOXML(xpp, lastTag) as EffectPropsGroup
                            // TODO: Eventually these will be objects
                        } else if (tnm == "latin") {
                            latin = xpp.getAttributeValue(0)
                        } else if (tnm == "ea") {
                            ea = xpp.getAttributeValue(0)
                        } else if (tnm == "cs") {
                            cs = xpp.getAttributeValue(0)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "rPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("rPr.parseOOXML: $e")
            }

            return RPr(attrs, l, fill, effect, latin, ea, cs)
        }
    }
}
