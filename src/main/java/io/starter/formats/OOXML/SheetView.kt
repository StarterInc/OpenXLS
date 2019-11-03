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

/**
 * sheetView (Worksheet View)
 *
 *
 * A single sheet view definition. When more than 1 sheet view is defined in the file, it means that when opening
 * the workbook, each sheet view corresponds to a separate window within the spreadsheet application, where
 * each window is showing the particular sheet. containing the same workbookViewId value, the last sheetView
 * definition is loaded, and the others are discarded. When multiple windows are viewing the same sheet, multiple
 * sheetView elements (with corresponding workbookView entries) are saved.
 *
 *
 * parent:		sheetViews
 * children:  	pane, selection, pivotSelection
 */
// TODO: finish pivotSelection
class SheetView : OOXMLElement {
    private var attrs = HashMap<String, Any>()
    private var pane: Pane? = null
    private var selections = ArrayList<Selection>()

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<sheetView")
            val i = attrs.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs[key] as String
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (pane != null) ooxml.append(pane!!.ooxml)
            if (selections.size > 0) {
                for (j in selections.indices) {
                    ooxml.append(selections[j].ooxml)
                }
            }
            ooxml.append("</sheetView>")
            return ooxml.toString()
        }

    constructor() {

    }

    constructor(attrs: HashMap<String, Any>, p: Pane, selections: ArrayList<Selection>) {
        this.attrs = attrs
        this.pane = p
        this.selections = selections
    }

    constructor(s: SheetView) {
        this.attrs = s.attrs
        this.pane = s.pane
        this.selections = s.selections
    }

    override fun cloneElement(): OOXMLElement {
        return SheetView(this)
    }


    /**
     * return the attribute value for key null if not found
     *
     * @param key
     * @return
     */
    fun getAttr(key: String): Any {
        return this.attrs[key]
    }

    /**
     * set the atttribute value for key to val
     *
     * @param key
     * @param val
     */
    fun setAttr(key: String, `val`: Any) {
        this.attrs[key] = `val`
    }

    /**
     * remove a previously set attribute, if found
     *
     * @param key
     */
    fun removeAttr(key: String) {
        this.attrs.remove(key)
    }

    fun removeSelection() {
        this.removeAttr("tabSelected")
        selections = ArrayList()
    }

    /**
     * return the attribute value for key
     * in String form "" if not found
     *
     * @param key
     * @return
     */
    fun getAttrS(key: String): String {
        val o = this.attrs[key] ?: return ""
        return o.toString()
    }

    companion object {

        private val serialVersionUID = 8750051341951797617L

        fun parseOOXML(xpp: XmlPullParser): OOXMLElement {
            val attrs = HashMap<String, Any>()
            var p: Pane? = null
            val selections = ArrayList<Selection>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "sheetView") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "pane") {
                            p = Pane.parseOOXML(xpp)
                        } else if (tnm == "selection") {
                            selections.add(Selection.parseOOXML(xpp))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "sheetView") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("sheetView.parseOOXML: $e")
            }

            return SheetView(attrs, p, selections)
        }
    }

}

/**
 * pane (View Pane)
 * Worksheet view pane
 *
 *
 * parent:  sheetView, customSheetView
 * children: none
 */
internal class Pane : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<pane")
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

    constructor(p: Pane) {
        this.attrs = p.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Pane(this)
    }

    companion object {
        private val serialVersionUID = 5570779997661362205L


        fun parseOOXML(xpp: XmlPullParser): Pane {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pane") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "pane") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("pane.parseOOXML: $e")
            }

            return Pane(attrs)
        }
    }
}

/**
 * selection (Selection)
 * Worksheet view selection.
 *
 *
 * parent: 	 sheetView, customSheetView
 * children: none
 */
internal class Selection : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<selection")
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

    constructor(s: Selection) {
        this.attrs = s.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Selection(this)
    }

    companion object {

        private val serialVersionUID = -5411798327743116154L

        fun parseOOXML(xpp: XmlPullParser): Selection {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "selection") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "selection") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("selection.parseOOXML: $e")
            }

            return Selection(attrs)
        }
    }
}

