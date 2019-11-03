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
 * defRPr (Default Text Run Properties)
 *
 *
 * This element contains all default run level text properties for the text runs within a containing paragraph. These
 * properties are to be used when overriding properties have not been defined within the rPr element
 *
 *
 * parent: many, including  pPr
 * children:  ln, FILLS, EFFECTS, highlight, TEXTUNDERLINE, TEXTUNDERLINEFILL, latin, ea, cs, sym, hlinkClick, hlinkMouseOver
 * many attributes
 */
// TODO: FINISH CHILD ELEMENTS highlight TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver
class DefRPr : OOXMLElement {
    private var fillGroup: FillGroup? = null
    private var effect: EffectPropsGroup? = null
    private var line: Ln? = null
    private var attrs: HashMap<String, String>? = null
    private var latin: String? = null
    private var ea: String? = null
    private var cs: String? = null    // really children but only have 1 attribute and no children

    /**
     * return all the text properties in hashmap from
     *
     * @return
     */
    // TODO: Fill, line ...
    val textProperties: HashMap<String, String>
        get() {
            val textprops = HashMap<String, String>()
            textprops.putAll(attrs!!)
            textprops["latin_typeface"] = latin
            textprops["ea_typeface"] = ea
            textprops["cs_typeface"] = cs
            return textprops
        }

    override// group fill
    // group effect
    // highlight
    // TEXTUNDERLINELINE
    // TEXTUNDERLINEFILL
    // hLinkClick
    // hLinkMouseOver
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:defRPr")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (line != null) ooxml.append(line!!.ooxml)
            if (fillGroup != null) ooxml.append(fillGroup!!.ooxml)
            if (effect != null) ooxml.append(effect!!.ooxml)
            if (latin != null) ooxml.append("<a:latin typeface=\"$latin\"/>")
            if (ea != null) ooxml.append("<a:ea typeface=\"$ea\"/>")
            if (cs != null) ooxml.append("<a:cs typeface=\"$cs\"/>")
            ooxml.append("</a:defRPr>")
            return ooxml.toString()
        }

    /**
     * create a default default Run Properties
     * for BIFF8 compatibility
     * qaa
     */
    constructor() {
        this.attrs = HashMap()
        this.attrs!!["sz"] = "900"
        this.attrs!!["b"] = "1"
        this.attrs!!["i"] = "0"
        this.attrs!!["u"] = "none"
        this.attrs!!["strike"] = "noStrike"
        this.attrs!!["baseline"] = "0"
        this.fillGroup = FillGroup(null, null, null, null, SolidFill())
        this.latin = "Arial"
        this.ea = "Arial"
        this.cs = "Arial"
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
        this.attrs = HashMap()
        this.attrs!!["sz"] = sz.toString()
        this.attrs!!["b"] = if (b) "1" else "0"
        this.attrs!!["i"] = if (i) "1" else "0"
        this.attrs!!["u"] = u
        this.attrs!!["strike"] = strike
        this.attrs!!["baseline"] = "0"
        this.fillGroup = FillGroup(null, null, null, null, SolidFill(clr))
        this.latin = fontFace
        this.ea = fontFace
        this.cs = fontFace
    }

    constructor(fill: FillGroup, effect: EffectPropsGroup, l: Ln, attrs: HashMap<String, String>, latin: String, ea: String, cs: String) {
        this.fillGroup = fill
        this.effect = effect
        this.line = l
        this.latin = latin
        this.ea = ea
        this.cs = cs
        this.attrs = attrs
    }

    constructor(dp: DefRPr) {
        this.fillGroup = dp.fillGroup
        this.effect = dp.effect
        this.line = dp.line
        this.latin = dp.latin
        this.ea = dp.ea
        this.cs = dp.cs
        this.attrs = dp.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return DefRPr(this)
    }

    companion object {

        private val serialVersionUID = 6764149567499222506L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var fill: FillGroup? = null
            var effect: EffectPropsGroup? = null
            var l: Ln? = null
            var latin: String? = null
            var ea: String? = null
            var cs: String? = null
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "defRPr") {        // default text properties
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
                        if (endTag == "defRPr") {
                            lastTag.pop()    // pop this tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("defPr.parseOOXML: $e")
            }

            return DefRPr(fill, effect, l, attrs, latin, ea, cs)
        }
    }
}

