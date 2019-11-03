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

import io.starter.OpenXLS.PivotTableHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.Boundsheet
import io.starter.formats.XLS.SxStreamID
import io.starter.formats.XLS.Sxvd
import io.starter.formats.XLS.Sxview
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserException
import org.xmlpull.v1.XmlPullParserFactory

import java.io.InputStream

class PivotTableDefinition : OOXMLElement {
    internal var ptview: Sxview? = null

    override// TODO: Finish
    val ooxml: String?
        get() = null

    /*
     * item types enum -- add an underscore because "default" is not a valid entry
     */
    internal enum class ITEMTYPES {
        _data,
        _default,
        _sum,
        _countA,
        _count,
        _avg,
        _max,
        _min,
        _product,
        _stdDev,
        _stdDevP,
        _var,
        _varP,
        _grand,
        _blank
    }
    /**
     * Changes a range in a PivotTable to expand until it includes the cell
     * address from CellHandle.
     *
     * Example:
     *
     * CellHandle cell = new Cellhandle("D4") PivotTable = SUM(A1:B2)
     * addCellToRange("A1:B2",cell); would change the PivotTable to look like
     * "SUM(A1:D4)"
     *
     * Returns false if PivotTable does not contain the PivotTableLoc range.
     *
     * @param String
     * - the Cell Range as a String to add the Cell to
     * @param CellHandle
     * - the CellHandle to add to the range
     *
     * public boolean addCellToRange(String PivotTableLoc, CellHandle
     * handle) throws PivotTableNotFoundException{ int[]
     * PivotTableaddr = ExcelTools.getRangeRowCol(PivotTableLoc);
     * String handleaddr = handle.getCellAddress(); int[] celladdr =
     * ExcelTools.getRowColFromString(handleaddr);
     *
     * // check existing range and set new range vals if the new Cell
     * is outside if(celladdr[0] >
     * PivotTableaddr[2])PivotTableaddr[2] = celladdr[0];
     * if(celladdr[0] < PivotTableaddr[0])PivotTableaddr[0] =
     * celladdr[0]; if(celladdr[1] >
     * PivotTableaddr[3])PivotTableaddr[3] = celladdr[1];
     * if(celladdr[1] < PivotTableaddr[1])PivotTableaddr[1] =
     * celladdr[1]; String newaddr =
     * ExcelTools.formatRange(PivotTableaddr); boolean b =
     * this.changePivotTableLocation(PivotTableLoc, newaddr); return
     * b; }
     */

    /**
     * Returns the "SXVIEW" record for this PivotTable.
     *
     * @return
     *
     * public Sxview getPt() { return pt; }
     */

    /**
     * @param sxview public void setPt(Sxview sxview) { pt= sxview; }
     */

    override fun cloneElement(): OOXMLElement? {
        return null
    }

    companion object {

        private val serialVersionUID = -5070227633357072878L


        /**
         * NOT COMPLETED DO NOT USE
         * parse a pivotTable OOXML element <br></br>
         * Top-level attributes  * Location information  * Collection of fields
         *  * Fields on the row axis  * Items on the row axis (specific values)
         *  * Fields on the column axis  * Items on the column axis (specific
         * values)  * Fields on the report filter region  * Fields in the values
         * region  * Style information <br></br>
         * Outline of the XML for a pivotTableDefinition (sequence)  *
         * pivotTableDefinition  * location  * pivotFields  * rowFields  *
         * rowItems  * colFields  * colItems  * pageFields  * dataFields  *
         * conditionalFormats  * pivotTableStyleInfo
         *
         * <br></br>
         *
         * <pre>
         * A PivotTable report that has more than one row field has one inner row field,
         * the one closest to the data area.
         *
         * Any other row fields are outer row fields
         *
         * Items in the outermost row field are displayed only once,
         * but items in the rest of the row fields are repeated as needed
         *
         * Page fields allow you to filter the entire PivotTable report to
         * display data for a single item or all the items.
         *
         * Data fields provide the data values to be summarized. Usually data fields contain numbers,
         * which are combined with the Sum summary function, but data fields can also contain text,
         * in which case the PivotTable report uses the Count summary function.
         *
         * If a report has more than one data field, a single field button named Data
         * appears in the report for access to all of the data fields.
        </pre> *
         *
         * @param bk
         * @param sheet
         * @param ii
         */
        fun parseOOXML(bk: WorkBookHandle, /*Object cacheid, */sheet: Boundsheet, ii: InputStream): PivotTableHandle {
            var ptview: Sxview? = null

            try {
                val factory = XmlPullParserFactory.newInstance()
                factory.isNamespaceAware = true
                val xpp = factory.newPullParser()
                xpp.setInput(ii, null) // using XML 1.0 specification
                var eventType = xpp.eventType

                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pivotTableDefinition") { // get attributes
                            var cacheId = ""
                            var tablename = ""
                            for (z in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(z)
                                val v = xpp.getAttributeValue(z)
                                if (nm == "name")
                                    tablename = v
                                else if (nm == "cacheId") {
                                    cacheId = v
                                }
                                //dataOnRows
                                //applyNumberFormats
                                //applyBorderFormats
                                //applyFontFormats
                                //applyPatternFormats
                                //applyAlignmentFormats
                                //applyWidthHeightFormats
                                //dataCaption
                                //updatedVersion
                                //showMemberPropertyTips
                                //useAutoFormatting
                                //itemPrintTitles
                                //createdVersion
                                //indent
                                //compact
                                //compactData
                                //gridDropZones
                            }
                            // KSC: TESTING!!!
                            cacheId = "0"
                            val cid = Integer.valueOf(cacheId).toShort()
                            val ptstream = bk.workBook!!.getPivotStream(cid + 1)
                            ptview = sheet.addPivotTable(ptstream!!.cellRange!!.toString(), bk, cid + 1, tablename)
                            ptview!!.dataName = "Values"    // default
                            ptview.iCache = cid
                        } else if (tnm == "location") {
                            parseLocationOOXML(xpp, ptview, bk)
                        } else if (tnm == "pivotFields") { // Represents the collection of fields that appear on the PivotTable.
                            parsePivotFields(xpp, ptview)
                        } else if (tnm == "pageFields") {    // Represents the collection of items in the page or report filter region of the PivotTable.
                            //						short count = Integer.valueOf(xpp.getAttributeValue(0)).shortValue();
                            //						ptview.setCDimPg(count);
                        } else if (tnm == "pageField") { // count: # of pageField elements
                            parsePageFieldOOXML(xpp, ptview)
                        } else if (tnm == "dataFields") {    // Represents the collection of items in the data region of the PivotTable.
                            //						short count = Integer.valueOf(xpp.getAttributeValue(0)).shortValue();
                            //						ptview.setCDimData(count);
                        } else if (tnm == "dataField") { // count: # of dataField elements
                            parseDataFieldOOXML(xpp, ptview)
                        } else if (tnm == "colFields" || tnm == "rowFields") {    // the collection of fields on the column or row axis
                            parseFieldOOXML(xpp, ptview)
                        } else if (tnm == "rowItems" || tnm == "colItems") {        // the collection of column items or row items
                            parseLineItemOOXML(xpp, ptview)
                        } else if (tnm == "formats") {
                            parseFormatsOOXML(xpp, ptview)
                        } else if (tnm == "chartFormats") {
                            parseFormatsOOXML(xpp, ptview)
                        } else if (tnm == "pivotTableStyleInfo") {
                        }
                    } else if (eventType == XmlPullParser.END_TAG) { // go to end of file
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("PivotTableDefinition.parseOOXML: $e")
            }

            return PivotTableHandle(ptview, bk)
        }

        /**
         * parses a pivotTableDefinition location firstDataCol, firstDataRow,
         * firstHeaderRow, ref [all req], colPageCount, rowPageCount
         *
         * @param xpp
         * @param pth
         * @return
         * @throws XmlPullParserException
         */
        @Throws(XmlPullParserException::class)
        private fun parseLocationOOXML(xpp: XmlPullParser, ptview: Sxview?, bk: WorkBookHandle) {
            try {
                for (i in 0 until xpp.attributeCount) {
                    val nm = xpp.getAttributeName(i)
                    val v = xpp.getAttributeValue(i)
                    if (nm.equals("ref", ignoreCase = true))
                    // req; Specifies the first row of the actual PivotTable (NOT the data)
                        ptview!!.setLocation(v)
                    else if (nm.equals("firstDataCol", ignoreCase = true))
                    // req
                    // Specifies the first column of the PivotTable data, relative to the top left cell in the ref value.
                        ptview!!.colFirstData = Integer.valueOf(v).toShort()
                    else if (nm.equals("firstDataRow", ignoreCase = true))
                    // req
                    // Specifies the first row of the PivotTable data, relative to the top left cell in the ref value.
                        ptview!!.rwFirstData = Integer.valueOf(v).toShort()
                    else if (nm.equals("firstHeaderRow", ignoreCase = true))
                    // req
                    // Specifies the first row of the PivotTable header relative to the top left cell in the ref value.
                        ptview!!.rwFirstHead = Integer.valueOf(v).toShort()
                    else if (nm.equals("rowPageCount", ignoreCase = true))
                    else if (nm.equals("colPageCount", ignoreCase = true))
                    ;// def= 0
                    // Specifies the number of columns per page for this PivotTable that the filter area will occupy. By default
                    // there is a single column of filter fields per page and the fields occupy as many rows as there are fields.
                    // def= 0
                    // Specifies the number of rows per page for this PivotTable that the filter area will occupy. By default there is a
                    // single column of filter fields per page and the fields occupy as many rows as there are fields.
                }
            } catch (e: Exception) {
                throw XmlPullParserException("PivotTableHandle.parseLocation:")
            }

        }

        /**
         * parses the pivotFields element of the pivotTableDefinition parent
         * <br></br>Represents the collection of fields that appear on the PivotTable.
         *
         * @param xpp
         * @throws XmlPullParserException
         */
        @Throws(XmlPullParserException::class)
        private fun parsePivotFields(xpp: XmlPullParser, ptview: Sxview?) {
            try {
                var eventType = xpp.eventType
                val elname = xpp.name
                var fcount = 0
                var curAxis: Sxvd? = null                        // up to 4 axes:  ROW, COL, PAGE or DATA
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pivotField") { // Represents a single field in the PivotTable. This complex type contains information about the field, including the collection of items in the field.
                            curAxis = null
                            for (i in 0 until xpp.attributeCount) {        // TODO: HANDLE ALL ATTRIBUTES *****
                                val nm = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (nm.equals("axis", ignoreCase = true))
                                // Specifies the region of the PivotTable that this field is displayed
                                    curAxis = ptview!!.addPivotFieldToAxis(axisLookup(v), fcount++)    // axisPage, axisRow, axisCol
                                else if (nm.equals("showAll", ignoreCase = true))
                                else if (nm.equals("defaultSubtotal", ignoreCase = true))
                                else if (nm.equals("numFmtId", ignoreCase = true))
                                else if (nm.equals("dataField", ignoreCase = true)) {    // Specifies a boolean value that indicates whether this field appears in the data region of the PivotTable.
                                    if (v == "1" && curAxis == null)
                                        curAxis = ptview!!.addPivotFieldToAxis(axisLookup("axisValues"), fcount++)        // DATA axis
                                } else if (nm.equals("multipleItemSelectionAllowed", ignoreCase = true))
                                else if (nm.equals("sortType", ignoreCase = true))
                                ;// ascending, descending or manual
                                // Specifies a boolean value that indicates whether the field can have multiple items selected in the page field.
                                // Specifies the identifier of the number format to apply to this field.
                                // Specifies a boolean value that indicates whether the default subtotal aggregation // function is displayed for this field.
                                // Specifies a boolean value that indicates whether to show all items for this field.
                                // A value of off, 0, or false indicates items be shown according to user specified criteria
                            }

                        } else if (tnm == "items") {    // Represents the collection of items in a PivotTable field. The items in the collection are ordered by index. Items
                            // represent the unique entries from the field in the source data.
                            parsePivotItemOOXML(xpp, ptview, curAxis)
                        } else if (tnm == "pivotArea") { // parent= autoSortScope, which has no attributes or other children
                            parsePivotAreaOOXML(xpp)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == elname)
                            break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                throw XmlPullParserException("parsePivotFields:$e")
            }

        }

        /**
         * parses pivot items -- Represents the collection of items in a PivotTable
         * field. The items in the collection are ordered by index. Items represent
         * the unique entries from the field in the source data The order in which
         * the items are listed is the order they would appear on a particular axis
         * (row or column, for example). parent= pivotField
         *
         * @param xpp
         * @throws XmlPullParserException
         */
        @Throws(XmlPullParserException::class)
        private fun parsePivotItemOOXML(xpp: XmlPullParser, ptview: Sxview?, axis: Sxvd?) {
            try {
                var eventType = xpp.eventType
                val elname = xpp.name
                //int count = Integer.valueOf(xpp.getAttributeValue(0)).intValue();	// count # of item elements
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "item") { // Represents a single item in PivotTable field.
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (nm == "c")
                                else if (nm == "d")
                                else if (nm == "e")
                                else if (nm == "f")
                                else if (nm == "h")
                                else if (nm == "m")
                                else if (nm == "n")
                                else if (nm == "s")
                                else if (nm == "sd")
                                else if (nm == "t")
                                // Specifies the type of this item. A value of 'default' indicates the subtotal or total item.
                                    if (v == "default")
                                        ptview!!.addPivotItem(axis!!, 1, -1)
                                    else
                                        Logger.logWarn("PivitItem: Unknown type")    // REMOVE WHEN TESTED
                                else if (nm == "x")
                                // Specifies the item index in pivotFields collection in the PivotCache. Applies only non- OLAP PivotTables.
                                    ptview!!.addPivotItem(axis!!, 0, Integer.valueOf(v).toInt())// Specifies a boolean value that indicates whether the details are hidden for this item.
                                // Specifies a boolean value that indicates whether the item has a character value.
                                // Specifies the user caption of the item.
                                // Specifies a boolean value that indicate whether the item has a missing value.
                                // Specifies a boolean value that indicates whether the item is hidden.
                                // Specifies a boolean value that indicates whether this item is a calculated member.
                                // Specifies a boolean value that indicates whether attribute hierarchies nested next to each other on a PivotTable row or column will offer drilling "across" each other or not.
                                // Specifies a boolean value that indicates whether this item has been expanded in the PivotTable view.
                                // Specifies a boolean value that indicates whether the approximate number of child items for this item is greater than zero.
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == elname)
                            break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                throw XmlPullParserException("parsePivotItemOOXML:$e")
            }

        }

        /**
         * parses the format element of the pivotTableDefinition.  Represents the collection of formats applied to PivotTable.
         *
         * @param xpp
         * @throws XmlPullParserException
         */
        @Throws(XmlPullParserException::class)
        private fun parseFormatsOOXML(xpp: XmlPullParser, ptview: Sxview?) {
            try {
                var eventType = xpp.eventType
                val elname = xpp.name
                //int count = Integer.valueOf(xpp.getAttributeValue(0)).intValue();	// count # of item elements
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "format") {    // parent= formats
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                //String v = xpp.getAttributeValue(i);
                                if (nm == "dxfId") { // Specifies the identifier of the format the application is currently using for the PivotTable. Formatting information is written to the styles part. See the Styles section (§3.8) for more information on formats.

                                } else if (nm == "action") {
                                    /* 	Specifies the formatting behavior for the area indicated in the pivotArea element. The
									default value for this attribute is "formatting," which indicates that the specified cells
									have some formatting applied. The format is specified in the dxfId attribute. If the
									formatting is cleared from the cells, then the value of this attribute becomes "blank."
								 */
                                }
                            }
                        } else if (tnm == "chartFormat") {    // parent= chartFormats
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                //String v = xpp.getAttributeValue(i);
                                if (nm == "chart") { // Specifies the index of the chart part to which the formatting applies.
                                } else if (nm == "format") {    // Specifies the index of the pivot format that is currently in use. This index corresponds to a dxf element in the Styles part.
                                } else if (nm == "series") {    // Specifies a boolean value that indicates whether format applies to a series. (default=false)
                                }
                            }
                        } else if (tnm == "pivotArea")
                            parsePivotAreaOOXML(xpp)
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == elname)
                            break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                throw XmlPullParserException("parsePivotAreaOOXML:$e")
            }

        }

        /**
         * parse pivotArea element:  Rule describing a PivotTable selection (format, pivotField ...)
         *
         * @param xpp
         * @throws XmlPullParserException
         */
        @Throws(XmlPullParserException::class)
        private fun parsePivotAreaOOXML(xpp: XmlPullParser) {
            try {
                var eventType = xpp.eventType
                val elname = xpp.name
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pivotArea") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                //String v = xpp.getAttributeValue(i);
                                if (nm == "field")
                                else if (nm == "type")
                                else if (nm == "dataOnly")
                                else if (nm == "labelOnly")
                                else if (nm == "outline")
                                else if (nm == "axis")
                                else if (nm == "fieldPosition")
                                ;// Position of the field within the axis to which this rule applies.
                                //The region of the PivotTable to which this rule applies.
                                // Flag indicating whether the rule refers to an area that is in outline mode.
                                // Flag indicating whether only the item labels for an item selection are selected and does not include the data values (in the data area of the view).
                                // Flag indicating whether only the data values (in the data area of the view) for an item selection are selected and does not include the item labels.
                                // all, button, data, none, normal, origin, topRight
                                // Index of the field that this selection rule refers to.
                                // grandRow, grandCol, cacheIndex, offset, collapsedLevelsAreSubtotals
                            }
                        } else if (tnm == "references") { // Represents the set of selected fields and the selected items within those fields
                            // count
                        } else if (tnm == "reference") {
                            /*
						 * <attribute name="field" use="optional" type="xsd:unsignedInt"/>
8 <attribute name="count" type="xsd:unsignedInt"/>
9 <attribute name="selected" type="xsd:boolean" default="true"/>
10 <attribute name="byPosition" type="xsd:boolean" default="false"/>
11 <attribute name="relative" type="xsd:boolean" default="false"/>
12 <attribute name="defaultSubtotal" type="xsd:boolean" default="false"/>
13 <attribute name="sumSubtotal" type="xsd:boolean" default="false"/>
14 <attribute name="countASubtotal" type="xsd:boolean" default="false"/>
15 <attribute name="avgSubtotal" type="xsd:boolean" default="false"/>
16 <attribute name="maxSubtotal" type="xsd:boolean" default="false"/>
17 <attribute name="minSubtotal" type="xsd:boolean" default="false"/>
18 <attribute name="productSubtotal" type="xsd:boolean" default="false"/>
19 <attribute name="countSubtotal" type="xsd:boolean" default="false"/>
20 <attribute name="stdDevSubtotal" type="xsd:boolean" default="false"/>
21 <attribute name="stdDevPSubtotal" type="xsd:boolean" default="false"/>
22 <attribute name="varSubtotal" type="xsd:boolean" default="false"/>
23 <attribute name="varPSubtotal" type="xsd:boolean" default="false"/>
						 */
                        } else if (tnm == "x") {
                            //int index= parseItemIndexOOXML(xpp);
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == elname)
                            break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                throw XmlPullParserException("parsePivotAreaOOXML:")
            }

        }

        /**
         * element i parent= rowItem or colItem
         * <br></br>the collection of items in the row axis -- index corresponds to that in the location range
         * <br></br>OR
         * <br></br>the collection of column items-- index corresponds to that in the location range
         * <br></br>The first * collection represents all item values for the first column in the column axis area
         * <br></br>The first <x> in the first * corresponds to the first field in the columns area
         * or
         * <br></br>Represents the collection of items in the row region of the PivotTable.
         *
         * @param xpp
         * @throws XmlPullParserException
         *</x>* */
        @Throws(XmlPullParserException::class)
        private fun parseLineItemOOXML(xpp: XmlPullParser, ptview: Sxview?) {
            try {
                // number of row or col lines == cRw or cCol of SxVIEW
                //int linecount= Integer.valueOf(xpp.getAttributeValue(0)).intValue();	// parent element= rowItems or colItems
                var eventType = xpp.eventType
                val elname = xpp.name
                val isRowItems = elname == "rowItems"

                val enumType = ITEMTYPES::class.java
                var type = 0
                var repeat = 0
                var indexes: ShortArray? = null
                val nIndexes = (if (isRowItems) ptview!!.cDimRw else ptview!!.cDimCol).toInt()
                var index = 0
                var nLines = 0
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "i") { // a colItem or rowItem
                            indexes = ShortArray(nIndexes)
                            nLines = nIndexes
                            index = 0
                            repeat = 0
                            type = repeat
                            for (i in 0 until xpp.attributeCount) {    // i, r, t
                                val nm = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (nm == "i") {    // Specifies a zero-based index indicating the referenced data item it in a data field with multiple data items.
                                } else if (nm == "r") {// Specifies the number of items to repeat from the previous row item. Note: The first item has no @r explicitly written. Since a default of "0" is specified in the schema, for any item
                                    // whose @r is missing, a default value of "0" is implied.
                                    repeat = Integer.valueOf(v).toInt()
                                    index += repeat
                                    nLines -= repeat
                                } else if (nm == "t") {// Specifies the type of the item. Value of 'default' indicates a grand total as the last row item value
                                    // default= data, avg, blank, count, countA, data, grand, max, min, product, stdDev, stdDevP, sum, var, varP
                                    type = Enum.valueOf<ITEMTYPES>(enumType!!, "_$v").ordinal
                                    if (type != 0 || type != 0xE) {
                                        //									index++;	stil confused on this ...
                                        nLines--
                                    }
                                }

                            }
                        } else if (tnm == "x") {
                            indexes[index++] = parseItemIndexOOXML(xpp).toShort()            // v
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endname = xpp.name
                        if (endname == elname)
                            break
                        else if (endname == "i") {
                            if (isRowItems)
                                ptview.addPivotLineToROWAxis(repeat, nLines, type, indexes)
                            else
                                ptview.addPivotLineToCOLAxis(repeat, nLines, type, indexes)
                        }

                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                throw XmlPullParserException("parseItemOOXML:$e")
            }

        }

        /**
         * Represents a generic field that can appear either on the column or the
         * row region of the PivotTable. There will be as many <x> elements as there
         * are item values in any particular column or row. attribute: x: Specifies
         * the index to a pivotField item value
         *
         * @param xpp
         * @throws XmlPullParserException
        </x> */
        @Throws(XmlPullParserException::class)
        private fun parseFieldOOXML(xpp: XmlPullParser, ptview: Sxview?) {
            try {
                //int fieldcount = Integer.valueOf(xpp.getAttributeValue(0)).intValue();
                var eventType = xpp.eventType
                val elname = xpp.name
                /*			if (elname.equals("rowFields"))
				ptview.setCDimRw((short)fieldcount);
			else
				ptview.setCDimCol((short)fieldcount);*/
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "field") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "x") {
                                    val v = xpp.getAttributeValue(i)
                                    if (elname == "rowFields")
                                        ptview!!.addRowField(Integer.valueOf(v))
                                    else
                                        ptview!!.addColField(Integer.valueOf(v))
                                }
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == elname)
                            break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                throw XmlPullParserException("parseFieldOOXML:$e")
            }

        }

        /**
         * This element represents an array of indexes to cached shared item values
         * <br></br>element x
         * <br></br>child of reference (...pivotField), i (rowField, colField)
         *
         * @param xpp
         * @throws XmlPullParserException
         * @return int index
         */
        private fun parseItemIndexOOXML(xpp: XmlPullParser): Int {
            var v = 0
            try {
                if (xpp.attributeCount > 0)
                    v = Integer.valueOf(xpp.getAttributeValue(0)).toInt()    // index
            } catch (e: Exception) {
            }

            return v
        }

        private fun parseDataFieldOOXML(xpp: XmlPullParser, ptview: Sxview?) {
            /*
 <attribute name="subtotal"
		 * type="ST_DataConsolidateFunction" default="sum"/> 9
		 * <attribute name="showDataAs" type="ST_ShowDataAs"
		 * default="normal"/> 10 <attribute name="baseField"
		 * type="xsd:int" default="-1"/> 11 <attribute
		 * name="baseItem" type="xsd:unsignedInt"
		 * default="1048832"/> 12 <attribute name="numFmtId"
		 * type="ST_NumFmtId" use="optional"/>
		 */
            var fieldIndex = 0
            var aggregateFunction: String? = null
            var name: String? = null
            for (z in 0 until xpp.attributeCount) {
                val nm = xpp.getAttributeName(z)
                val v = xpp.getAttributeValue(z)
                if (nm == "name")
                    name = v
                else if (nm == "fld")
                    fieldIndex = Integer.valueOf(xpp.getAttributeValue(z))
                else if (nm == "subtotal")
                // default= "sum"
                    aggregateFunction = v
            }
            // TODO:
            // showDataAs, baseItem, baseField		--> display format
            // numFmtId
            ptview!!.addDataField(fieldIndex, aggregateFunction, name)
        }

        /**
         * parse the pageField element, which defines a pivot field on the PAGE axis
         *
         * @param xpp
         * @param ptview
         */
        private fun parsePageFieldOOXML(xpp: XmlPullParser, ptview: Sxview?) {
            var fieldIndex = 0
            var itemIndex = 0x7FFD
            for (i in 0 until xpp.attributeCount) {
                val nm = xpp.getAttributeName(i)
                if (nm == "fld") {
                    fieldIndex = Integer.valueOf(xpp.getAttributeValue(i))
                } else if (nm == "item") {
                    itemIndex = Integer.valueOf(xpp.getAttributeValue(i))
                }
            }
            ptview!!.addPageField(fieldIndex, itemIndex)
        }

        private fun axisLookup(axis: String): Int {
            if (axis == "axisRow")
                return Sxvd.AXIS_ROW.toInt()
            else if (axis == "axisCol")
                return Sxvd.AXIS_COL.toInt()
            else if (axis == "axisPage")
                return Sxvd.AXIS_PAGE.toInt()
            else if (axis == "axisValues")
                return Sxvd.AXIS_DATA.toInt()
            return Sxvd.AXIS_NONE.toInt()

        }
    }

}
