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

import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.FormatConstants
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.HashMap
import java.util.Stack

/**
 * p (Text Paragraph)
 *
 *
 * This element specifies the presence of a paragraph of text within the containing text body. The paragraph is the
 * highest level text separation mechanism within a text body. A paragraph may contain text paragraph properties
 * associated with the paragraph. If no properties are listed then properties specified in the defPPr element are
 * used.
 *
 *
 * parent:  r, t, txBody, txpr
 * children:  pPr, (r, br or fld), endParaRPr
 */
// TODO: Finish pPr Text Paragraph Properties -- MANY child elements not handled
//TODO: Finish endParaRPr children TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver,
class P : OOXMLElement {
    private var run: TextRun? = null
    private var ppr: PPr? = null
    private var ep: EndParaRPr? = null

    override// 20090526 KSC: order of children was wrong (Kaylan/Rajesh chart error)
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:p>")
            if (ppr != null) ooxml.append(ppr!!.ooxml)
            if (run != null)
                ooxml.append(run!!.ooxml)
            if (ep != null) ooxml.append(ep!!.ooxml)
            ooxml.append("</a:p>")
            return ooxml.toString()
        }

    val title: String?
        get() = if (run != null) run!!.title else ""

    /**
     * returns a map of all text properties defined for this paragraph
     *
     * @return
     */
    // algn- left, right, centered, just, distributed
    // defTabSz
    // fontAlgn (Font Alignment)
    // hangingPunct (Hanging Punctuation)  bool
    // indent (Indent)
    // lvl (Level)
    // marL (Left Margin)
    // marR (Right Margin)
    // rtl (Right To Left) bool
    /* This element contains all default run level text properties for the text runs within a containing paragraph. These
				10 properties are to be used when overriding properties have not been defined within the rPr element.*//*
             *	altLang (Alternative Language)
             *  b (Bold)		bool
             *  baseline (Baseline)
             *  bmk (Bookmark Link Target)
             *  cap (Capitalization)
             *  i (Italics)		bool
             *  kern (Kerning)
             *  kumimoji
             *  lang (Language ID)
             *  spc (Spacing)
             *  strike (Strikethrough)
             *  sz (Font Size)	size
             *  u (Underline)	underline style
             *
             *  PLUS-- fill, line, blipFill, cs/ea/latin font (attribute: typeface) ****
             */ val textProperties: HashMap<String, String>
        get() {
            val textprops = HashMap<String, String>()
            if (ppr != null) {
                textprops.putAll(ppr!!.textProperties)
                textprops.putAll(ppr!!.defaultTextProperties)
            }
            if (run != null) {
                textprops.putAll(run!!.textProperties)
            }
            return textprops
        }

    constructor(s: String) {
        this.run = TextRun(s)
        this.ppr = PPr(DefRPr(), null)
    }

    constructor(ppr: PPr, run: TextRun, ep: EndParaRPr) {
        this.ppr = ppr
        this.ep = ep
        this.run = run
    }

    constructor(p: P) {
        this.ppr = p.ppr
        this.ep = p.ep
        this.run = p.run
    }

    /**
     * create a default paragraph property from the specified information
     *
     * @param fontFace String font face e.g. "Arial"
     * @param sz       int size in 100 pts (e.g. font size 12.5 pts,, sz= 1250)
     * @param b        boolean true if bold
     * @param i        boolean true if italic
     * @param u        String underline.  One of the following Strings:  dash, dashHeavy, dashLong, dashLongHeavy, dbl, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, dotted
     * dottedHeavy, heavy, none, sng, wavy, wavyDbl, wavyHeavy, words (underline only words not spaces)
     * @param strike   String strike setting.  One of the following Strings: dblStrike, noStrike or sngStrike  or null if none
     * @param clr      String fill color in hex form without the #
     */
    constructor(fontFace: String, sz: Int, b: Boolean, i: Boolean, u: String, strike: String, clr: String) {
        this.ppr = PPr(fontFace, sz, b, i, u, strike, clr)
    }


    /**
     * create a default paragraph property from the specified Font and text s
     *
     * @param f Font
     */
    constructor(f: Font, s: String) {
        val u = f.underlineStyle
        var usty = "none"
        when (u) {
            FormatConstants.STYLE_UNDERLINE_SINGLE -> usty = "sng"
            FormatConstants.STYLE_UNDERLINE_DOUBLE -> usty = "dbl"
            FormatConstants.STYLE_UNDERLINE_SINGLE_ACCTG -> usty = "sng"
            FormatConstants.STYLE_UNDERLINE_DOUBLE_ACCTG -> usty = "dbl"
        }
        val strike = if (f.stricken) "sngStrike" else "noStrike"
        val clr = FormatHandle.colorToHexString(FormatHandle.getColor(f.color)).substring(1)
        this.ppr = PPr(f.fontName, f.fontHeightInPoints.toInt() * 100, f.bold, f.italic, usty, strike, clr)
        this.run = TextRun(s)
    }

    override fun cloneElement(): OOXMLElement {
        return P(this)
    }

    companion object {

        private val serialVersionUID = 6302706683933521698L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var ppr: PPr? = null
            var run: TextRun? = null
            var ep: EndParaRPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pPr") {        // paragraph-level text props
                            lastTag.push(tnm)
                            ppr = PPr.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "r" ||
                                tnm == "br" ||
                                tnm == "fld") {    // text run
                            lastTag.push(tnm)
                            run = TextRun.parseOOXML(xpp, lastTag, bk) as TextRun
                        } else if (tnm == "endParaRPr") {
                            lastTag.push(tnm)
                            ep = EndParaRPr.parseOOXML(xpp, lastTag, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "p") {
                            lastTag.pop()    // pop this tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("p.parseOOXML: $e")

            }

            return P(ppr, run, ep)
        }
    }
}

/**
 * pPr (Text Paragraph Properties)
 *
 *
 * This element contains all paragraph level text properties for the containing paragraph. These paragraph
 * properties should override any and all conflicting properties that are associated with the paragraph in question
 *
 *
 * parent: p, fld
 * children: many
 * attributes: many
 */
// TODO: Handle child elements: lnSpc, spcBef, spcAft, TEXTBULLETCOLOR, TEXTBULLETSIZE, TEXTBULLET, tabLst
internal class PPr : OOXMLElement {
    private var dp: DefRPr? = null
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:pPr")
            if (attrs != null) {
                val i = attrs!!.keys.iterator()
                while (i.hasNext()) {
                    val key = i.next()
                    val `val` = attrs!![key]
                    ooxml.append(" $key=\"$`val`\"")
                }
            }
            ooxml.append(">")
            if (dp != null) ooxml.append(dp!!.ooxml)
            ooxml.append("</a:pPr>")
            return ooxml.toString()
        }

    /**
     * return the default run properties for this para
     *
     * @return
     */
    val defaultTextProperties: HashMap<String, String>
        get() = dp!!.textProperties

    /**
     * return the text properties defined by PPr:
     * algn- left, right, centered, just, distributed
     * defTabSz
     * fontAlgn (Font Alignment)
     * hangingPunct (Hanging Punctuation)  bool
     * indent (Indent)
     * lvl (Level)
     * marL (Left Margin)
     * marR (Right Margin)
     * rtl (Right To Left) bool
     *
     * @return
     */
    val textProperties: HashMap<String, String>
        get() = if (attrs != null) attrs else HashMap()

    constructor(dp: DefRPr, attrs: HashMap<String, String>) {
        this.dp = dp
        this.attrs = attrs
    }

    constructor(p: PPr) {
        this.dp = p.dp
        this.attrs = p.attrs
    }

    /**
     * create a default paragraph property from the specified information
     *
     * @param fontFace String font face e.g. "Arial"
     * @param sz       int size in 100 pts (e.g. font size 12.5 pts,, sz= 1250)
     * @param b        boolean true if bold
     * @param i        boolean true if italic
     * @param u        String underline.  One of the following Strings:  dash, dashHeavy, dashLong, dashLongHeavy, dbl, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, dotted
     * dottedHeavy, heavy, none, sng, wavy, wavyDbl, wavyHeavy, words (underline only words not spaces)
     * @param strike   String strike setting.  One of the following Strings: dblStrike, noStrike or sngStrike  or null if none
     * @param clr      String fill color in hex form without the #
     */
    constructor(fontFace: String, sz: Int, b: Boolean, i: Boolean, u: String, strike: String, clr: String) {
        this.dp = DefRPr(fontFace, sz, b, i, u, strike, clr)
    }

    override fun cloneElement(): OOXMLElement {
        return PPr(this)
    }

    companion object {

        private val serialVersionUID = -6909210948618654877L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): PPr {
            var dp: DefRPr? = null
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pPr") {        // t element of text run -- the title string we are interested in
                            for (i in 0 until xpp.attributeCount) {    // align, defTabSz, fontAlgn, hangingPunct, indent, lvl, rtl, marL, marR ...
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "defRPr") {        // default text properties
                            lastTag.push(tnm)
                            dp = DefRPr.parseOOXML(xpp, lastTag, bk) as DefRPr
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pPr") {
                            lastTag.pop()    // pop this element
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("paraText.parseOOXML: $e")
            }

            return PPr(dp, attrs)
        }
    }
}


/**
 * endParaRPr (End Paragraph Run Properties)
 *
 *
 * This element specifies the text run properties that are to be used if another run is inserted after the last run
 * specified. This effectively saves the run property state so that it may be applied when the user enters additional
 * text. If this element is omitted, then the application may determine which default properties to apply. It is
 * recommended that this element be specified at the end of the list of text runs within the paragraph so that an
 * orderly list is maintained.
 *
 *
 * parent: p
 * children: ln, FillGroup, EffectGroup, highlight, TEXTUNDERLINE, TEXTUNDERLINEFILL, latin, ea, cs, sym, hlinkClick, hlinkMouseOver,
 */
// TODO: Finish children TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver,
internal class EndParaRPr : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var l: Ln? = null
    private var fill: FillGroup? = null
    private var effect: EffectPropsGroup? = null
    private var latin: String? = null
    private var ea: String? = null
    private var cs: String? = null    // really children but only have 1 attribute and no children

    override// attributes
    // group fill
    // sym
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:endParaRPr")
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
            if (latin != null) ooxml.append("<a:latin typeface=\"$latin\"/>")
            if (ea != null) ooxml.append("<a:ea typeface=\"$ea\"/>")
            if (cs != null) ooxml.append("<a:cs typeface=\"$cs\"/>")
            ooxml.append("</a:endParaRPr>")
            return ooxml.toString()
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

    constructor(ep: EndParaRPr) {
        this.attrs = ep.attrs
        this.l = ep.l
        this.fill = ep.fill
        this.effect = ep.effect
        this.latin = ep.latin
        this.ea = ep.ea
        this.cs = ep.cs
    }

    override fun cloneElement(): OOXMLElement {
        return EndParaRPr(this)
    }

    companion object {

        private val serialVersionUID = -7094231887468090281L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): EndParaRPr {
            val attrs = HashMap<String, String>()
            var fill: FillGroup? = null
            var effect: EffectPropsGroup? = null
            var l: Ln? = null
            var latin: String? = null
            var ea: String? = null
            var cs: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "endParaRPr") {        // get attributes
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
                        if (endTag == "endParaRPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("endParaRPr.parseOOXML: $e")
            }

            return EndParaRPr(attrs, l, fill, effect, latin, ea, cs)
        }
    }
}

