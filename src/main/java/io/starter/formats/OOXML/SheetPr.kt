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

/**
 * sheetPr (Sheet Properties)
 *
 *
 * Sheet-level properties.
 *
 *
 * parent:  worksheet, dialogsheet
 * children:  tabColor, outlinePr, pageSetUpPr
 */
// TODO: Finish pageSetUpPr  + input into 2003 sheet settings **
class SheetPr : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var tab: TabColor? = null
    private var outlinePr: OutlinePr? = null
    private var pageSetupPr: PageSetupPr? = null

    override// attributes
    // ordered sequence:
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<sheetPr")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }

            ooxml.append(">")
            if (tab != null) ooxml.append(tab!!.ooxml)
            if (outlinePr != null) ooxml.append(outlinePr!!.ooxml)
            if (pageSetupPr != null) ooxml.append(pageSetupPr!!.ooxml)
            ooxml.append("</sheetPr>")
            return ooxml.toString()
        }

    /**
     * return the codename used to link this sheet to vba code
     *
     * @return
     */
    /**
     * set the codename used to link this sheet to vba code
     *
     * @param codename
     */
    var codename: String?
        get() = if (attrs != null) attrs!!["codeName"] else null
        set(codename) {
            if (attrs == null)
                attrs = HashMap()
            attrs!!["codeName"] = codename
        }


    constructor(attrs: HashMap<String, String>, tab: TabColor, op: OutlinePr, pr: PageSetupPr) {
        this.attrs = attrs
        this.tab = tab
        this.outlinePr = op
        this.pageSetupPr = pr
    }

    constructor(sp: SheetPr) {
        this.attrs = sp.attrs
        this.tab = sp.tab
        this.outlinePr = sp.outlinePr
        this.pageSetupPr = sp.pageSetupPr
    }

    override fun cloneElement(): OOXMLElement {
        return SheetPr(this)
    }

    companion object {

        private val serialVersionUID = 1781567781060400234L

        /**
         * outlinePr
         * pageSetUpPr
         * tabColor
         * codeName
         * enableFormatConditionsCalculation
         * filterMode
         * published
         * syncHorizontal
         * syncRef
         * syncVertical
         * transitionEntry
         * transitionEvaluation
         *
         * @param xpp
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser): SheetPr {
            val attrs = HashMap<String, String>()
            var tab: TabColor? = null
            var op: OutlinePr? = null
            var pr: PageSetupPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "sheetPr") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "pageSetUpPr") { // Page setup properties of the worksheet
                            pr = PageSetupPr.parseOOXML(xpp)
                        } else if (tnm == "outlinePr") {
                            op = OutlinePr.parseOOXML(xpp)
                        } else if (tnm == "tabColor") {
                            tab = TabColor.parseOOXML(xpp)

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "sheetPr") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("sheetPr.parseOOXML: $e")
            }

            return SheetPr(attrs, tab, op, pr)
        }
    }
}

/**
 * (Sheet Tab Color)
 * Background color of the sheet tab.
 *
 *
 * parent: 	sheetPr
 * children: none
 */
internal class TabColor : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<tabColor")
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

    constructor(t: TabColor) {
        this.attrs = t.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return TabColor(this)
    }

    companion object {

        private val serialVersionUID = -2862996863147633555L


        fun parseOOXML(xpp: XmlPullParser): TabColor {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "tabColor") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "tabColor") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("tabColor.parseOOXML: $e")
            }

            return TabColor(attrs)
        }
    }
}

/**
 * outlinePr
 * Sheet Group Outline Settings
 * all attributes are optional, default values:
 * applyStyles="0"	showOutlineSymbols="1"	summaryBelow="1"	summaryRight="1"
 */
internal class OutlinePr : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<outlinePr")
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

    constructor(op: OutlinePr) {
        this.attrs = op.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return OutlinePr(this)
    }

    companion object {

        private val serialVersionUID = 3030511803286369045L


        fun parseOOXML(xpp: XmlPullParser): OutlinePr {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "outlinePr") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "outlinePr") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("OutlinePr.parseOOXML: $e")
            }

            return OutlinePr(attrs)
        }
    }
}

/**
 * pageSetUpPr (Page Setup Properties)
 * Page setup properties of the worksheet
 * attributes:	autoPageBreaks (default=true), fitToPage (default=false)
 */
internal class PageSetupPr : OOXMLElement {
    private var autoPageBreaks = true
    private var fitToPage = false

    override// attributes
    // if not default
    // if not default
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<pageSetUpPr")
            if (!autoPageBreaks)
                ooxml.append(" autoPageBreaks=\"0\"")
            if (fitToPage)
                ooxml.append(" fitToPage=\"1\"")
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(autoPageBreaks: Boolean, fitToPage: Boolean) {
        this.autoPageBreaks = autoPageBreaks
        this.fitToPage = fitToPage
    }

    constructor(pr: PageSetupPr) {
        this.autoPageBreaks = pr.autoPageBreaks
        this.fitToPage = pr.fitToPage
    }

    override fun cloneElement(): OOXMLElement {
        return PageSetupPr(this)
    }

    companion object {

        private val serialVersionUID = 3030511803286369045L

        fun parseOOXML(xpp: XmlPullParser): PageSetupPr {
            var autoPageBreaks = true
            var fitToPage = false
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pageSetUpPr") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "fitToPage")
                                    fitToPage = xpp.getAttributeValue(i) == "1"
                                else if (nm == "autoPageBreaks")
                                    autoPageBreaks = xpp.getAttributeValue(i) == "1"
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pageSetUpPr") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("PageSetupPr: $e")
            }

            return PageSetupPr(autoPageBreaks, fitToPage)
        }
    }
}


