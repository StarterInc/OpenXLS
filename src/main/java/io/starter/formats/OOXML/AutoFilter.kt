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
 * autoFilter (AutoFilter Settings)
 * AutoFilter temporarily hides rows based on a filter criteria, which is applied column by column to a table of data
 * in the worksheet. This collection expresses AutoFilter settings.
 *
 *
 * parent: 		worksheet, table, filter, customSheetView
 * children: 	filterColumn (0+), sortState
 * attributes:	ref
 */
// TODO: finish sortState
// TODO: finish filterColumn children filters->filter, dataGroupItem
class AutoFilter : OOXMLElement {
    private var ref: String? = null
    private var filterColumns: ArrayList<FilterColumn>? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<autoFilter")
            if (ref != null) ooxml.append(" ref=\"$ref\"")
            ooxml.append(">")
            if (filterColumns != null) {
                for (i in filterColumns!!.indices)
                    ooxml.append(filterColumns!![i].ooxml)
            }
            ooxml.append("</autoFilter>")
            return ooxml.toString()
        }

    constructor(ref: String, f: ArrayList<FilterColumn>) {
        this.ref = ref
        this.filterColumns = f
    }

    constructor(a: AutoFilter) {
        this.ref = a.ref
        this.filterColumns = a.filterColumns
    }

    override fun cloneElement(): OOXMLElement {
        return AutoFilter(this)
    }

    companion object {

        private val serialVersionUID = 7111401348177004218L


        fun parseOOXML(xpp: XmlPullParser): OOXMLElement {
            var ref: String? = null
            var f: ArrayList<FilterColumn>? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "autoFilter") {        // get ref attribute
                            if (xpp.attributeCount == 1) ref = xpp.getAttributeValue(0)
                        } else if (tnm == "sortState") {
                        } else if (tnm == "filterColumn") {
                            if (f == null) f = ArrayList()
                            f.add(FilterColumn.parseOOXML(xpp))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "autoFilter") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("autoFilter.parseOOXML: $e")
            }

            return AutoFilter(ref, f)
        }
    }
}

/**
 * filterColumn (AutoFilter Column)
 * The filterColumn collection identifies a particular column in the AutoFilter range and specifies filter information
 * that has been applied to this column. If a column in the AutoFilter range has no criteria specified, then there is
 * no corresponding filterColumn collection expressed for that column
 *
 *
 * parent: 		autoFilter
 * children:	CHOICE OF: colorFilter, customFilters, dynamicFilter, filters, iconFilter, top10
 * attributes:	colId REQ, hiddenButton, showButton
 */
internal class FilterColumn : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var filter: Any? = null        // CHOICE of filter

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<filterColumn ")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            if (filter is ColorFilter) ooxml.append((filter as ColorFilter).ooxml)
            if (filter is CustomFilters) ooxml.append((filter as CustomFilters).ooxml)
            if (filter is DynamicFilter) ooxml.append((filter as DynamicFilter).ooxml)
            if (filter is Filters) ooxml.append((filter as Filters).ooxml)
            if (filter is IconFilter) ooxml.append((filter as IconFilter).ooxml)
            if (filter is Top10) ooxml.append((filter as Top10).ooxml)
            ooxml.append("</filterColumn>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>, filter: Any) {
        this.attrs = attrs
        this.filter = filter
    }

    constructor(f: FilterColumn) {
        this.attrs = f.attrs
        this.filter = f.filter
    }

    override fun cloneElement(): OOXMLElement {
        return FilterColumn(this)
    }

    companion object {

        private val serialVersionUID = 5005589034415840928L


        fun parseOOXML(xpp: XmlPullParser): FilterColumn {
            val attrs = HashMap<String, String>()
            var filter: Any? = null        // CHOICE of filter
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "filterColumn") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "colorFilter") {
                            filter = ColorFilter.parseOOXML(xpp)
                        } else if (tnm == "customFilters") {
                        } else if (tnm == "dynamicFilter") {
                        } else if (tnm == "filters") {
                        } else if (tnm == "iconFilter") {
                        } else if (tnm == "top10") {
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "filterColumn") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("filterColumn.parseOOXML: $e")
            }

            return FilterColumn(attrs, filter)
        }
    }
}

/**
 * colorFilter (Color Filter Criteria)
 * This element specifies the color to filter by and whether to use the cell's fill or font color in the filter criteria. If
 * the cell's font or fill color does not match the color specified in the criteria, the rows corresponding to those cells
 * are hidden from view.
 *
 *
 * parent: 	filterColumn
 * children: none
 */
internal class ColorFilter : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<colorFilter")
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

    constructor(c: ColorFilter) {
        this.attrs = c.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return ColorFilter(this)
    }

    companion object {

        private val serialVersionUID = 7077951504723033275L

        fun parseOOXML(xpp: XmlPullParser): ColorFilter {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "colorFilter") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "colorFilter") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("colorFilter.parseOOXML: $e")
            }

            return ColorFilter(attrs)
        }
    }
}

/**
 * dynamicFilter (Dynamic Filter)
 * This collection specifies dynamic filter criteria. These criteria are considered dynamic because they can change,
 * either with the data itself (e.g., "above average") or with the current system date (e.g., show values for "today").
 * For any cells whose values do not meet the specified criteria, the corresponding rows shall be hidden from view
 * when the filter is applied.
 *
 *
 * parent: 	filterColumn
 * children: none
 */
internal class DynamicFilter : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<dynamicFilter")
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

    constructor(d: DynamicFilter) {
        this.attrs = d.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return DynamicFilter(this)
    }

    companion object {

        private val serialVersionUID = -473171074711686551L


        fun parseOOXML(xpp: XmlPullParser): DynamicFilter {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "dynamicFilter") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "dynamicFilter") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("dynamicFilter.parseOOXML: $e")
            }

            return DynamicFilter(attrs)
        }
    }
}

/**
 * iconFilter (Icon Filter)
 *
 *
 * This element specifies the icon set and particular icon within that set to filter by. For any cells whose icon does
 * not match the specified criteria, the corresponding rows shall be hidden from view when the filter is applied.
 *
 *
 * parent: filterColumn
 * children: none
 */
internal class IconFilter : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<iconFilter")
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

    constructor(i: IconFilter) {
        this.attrs = i.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return IconFilter(this)
    }

    companion object {

        private val serialVersionUID = -5897037678209125965L


        fun parseOOXML(xpp: XmlPullParser): IconFilter {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "iconFilter") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "iconFilter") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("iconFilter.parseOOXML: $e")
            }

            return IconFilter(attrs)
        }
    }
}

/**
 * customFilters (Custom Filters)
 * When there is more than one custom filter criteria to apply (an 'and' or 'or' joining two criteria), then this
 * element groups the customFilter elements together.
 */
internal class CustomFilters : OOXMLElement {
    private var and = false
    private var custfilter: Array<CustomFilter>? = null

    override// shouln't be!
    // shouldn't be!
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<customFilters")
            if (this.and) ooxml.append(" and=\"1\"")
            ooxml.append(">")
            if (custfilter != null) {
                if (custfilter!![0] != null)
                    ooxml.append(custfilter!![0].ooxml)
                if (custfilter!![1] != null) ooxml.append(custfilter!![1].ooxml)
            }
            ooxml.append("</customFilters>")
            return ooxml.toString()
        }

    constructor(and: Boolean, custfilter: Array<CustomFilter>) {
        this.and = and
        this.custfilter = custfilter
    }

    constructor(c: CustomFilters) {
        this.and = c.and
        this.custfilter = c.custfilter
    }

    override fun cloneElement(): OOXMLElement {
        return CustomFilters(this)
    }

    companion object {

        private val serialVersionUID = -2491942158519963335L


        fun parseOOXML(xpp: XmlPullParser): CustomFilters {
            var and = false
            val custfilter = arrayOfNulls<CustomFilter>(2)
            var idx = 0
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "customFilters") {        // get attributes
                            if (xpp.attributeCount == 1) and = xpp.getAttributeValue(0) == "1"
                        } else if (tnm == "customFilter") {    // 1-2
                            custfilter[idx++] = CustomFilter.parseOOXML(xpp)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "customFilters") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("customFilters.parseOOXML: $e")
            }

            return CustomFilters(and, custfilter)
        }
    }
}

/**
 * customFilter (Custom Filter Criteria)
 * A custom AutoFilter specifies an operator and a value. There can be at most two customFilters specified, and in
 * that case the parent element specifies whether the two conditions are joined by 'and' or 'or'. For any cells
 * whose values do not meet the specified criteria, the corresponding rows shall be hidden from view when the
 * filter is applied.
 *
 *
 * parent: customFilters
 * children: none
 */
internal class CustomFilter : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<customFilter")
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

    constructor(c: CustomFilter) {
        this.attrs = c.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return CustomFilter(this)
    }

    companion object {

        private val serialVersionUID = 7995078604042667255L

        fun parseOOXML(xpp: XmlPullParser): CustomFilter {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "customFilter") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "customFilter") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("customFilter.parseOOXML: $e")
            }

            return CustomFilter(attrs)
        }
    }
}

/**
 * top10 (Top 10)
 * This element specifies the top N (percent or number of items) to filter by.
 *
 *
 * parent: filterColumn
 * children: none
 */
internal class Top10 : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<top10")
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

    constructor(t: Top10) {
        this.attrs = t.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Top10(this)
    }

    companion object {

        private val serialVersionUID = 77735498689922082L


        fun parseOOXML(xpp: XmlPullParser): Top10 {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "top10") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "top10") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("top10.parseOOXML: $e")
            }

            return Top10(attrs)
        }
    }
}


/**
 * filters (Filter Criteria)
 * When multiple values are chosen to filter by, or when a group of date values are chosen to filter by, this element
 * groups those criteria together.
 *
 *
 * parent:  filterColumn
 * children: filter (0+), dateGroupItem (0+)
 */
// TODO: finish children
internal class Filters : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    // TODO: one or more filter elements
    // TODO: one or more dateGroupItem elements
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<filters")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            ooxml.append("</filters>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(f: Filters) {
        this.attrs = f.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return Filters(this)
    }

    companion object {

        private val serialVersionUID = 921424089049938924L

        fun parseOOXML(xpp: XmlPullParser): Filters {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "filters") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "filter") {
                        } else if (tnm == "dateGroupItem") {
                            //layout = (layout) layout.parseOOXML(xpp, lastTag).clone();
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "filters") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("filters.parseOOXML: $e")
            }

            return Filters(attrs)
        }
    }
}

	
