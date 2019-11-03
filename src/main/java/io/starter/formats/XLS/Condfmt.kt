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
package io.starter.formats.XLS

import io.starter.OpenXLS.ExcelTools
import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.CfRule
import io.starter.formats.OOXML.Dxf
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList


/**
 * **Condfmt:  Conditional Formatting Range Information 0x1B0**<br></br>
 *
 *
 * This record stores a conditional format, including conditions and formatting info.
 *
 *
 * And, no it does not just point to an Xf record because that would be easy.
 *
 *
 *
 *
 * OFFSET       NAME            SIZE        CONTENTS
 * -----
 * 4                ccf                 2           Number of Conditional formats
 * 6                grbit               2           Option flags (not a byte?) [1 = Conditionally formatted cells need recalculation or redraw]
 * 8                rwFirst             2           First row to conditionally format (0 based)
 * 10               rwLast              2           Last row to conditionally format (0 based)
 * 12               colFirst            2           First column to conditionally format (0 based)
 * 14               colLast             2           Last column to conditionally format (0 based)
 * 16               sqrefCount          2           Count of sqrefs *
 * 18               rgbSqref            var         Array of sqref structures
 *
 *
 *
 *
 * Sqref Structures
 *
 *
 * OFFSET       NAME            SIZE        CONTENTS
 * -----
 * 0            rwFirst             2           First row in reference
 * 2            rwLast              2           Last row in reference
 * 4            colFirst            2           First column in reference
 * 6            colLast             2           Last column in reference
 *
 * @see Cf
 */
/**
 * default constructor
 */
class Condfmt : io.starter.formats.XLS.XLSRecord() {

    /**
     * @return Returns the formatHandle.
     */
    /**
     * @param formatHandle The formatHandle to set.
     */
    var formatHandle: FormatHandle? = null
    internal var grbit: Short = 0        //      Option flags (not a byte?)

    private var ccf: Int = 0
    internal var refs: DiscontiguousRefStruct? = null
    /**
     * returns the rules associated with this record
     *
     * @return
     */
    val rules = ArrayList() // 2003-version Cf recs OR OOXML cfRules TODO: eventually will generate Cf records instead
    /**
     * This is an overall not ideal situation where we want a formatid that we can use in sheetster
     * to identify this conditional format
     *
     *
     * // TODO: Perfect this algorithm!! :)  cfxe should be constant for this Condfmt
     * // ... if address changes?  if sheet # changes
     *
     * @return Returns the cfxe.
     */
    /**
     * @param cfxe The cfxe to set.
     */
    // base cxfe on cell address
    var cfxe = -1
        get() {
            val rc = refs!!.rowColBounds
            this.cfxe = 50000 + this.sheet!!.sheetNum * 10000 + ByteTools.readShort(rc[0], rc[1]).toInt()
            return field
        } // a fake ixfe for use by OpenXLS to track formats
    internal var isdirty = false    // if any changes to underlying record is made, set to true

    /**
     * Return all ranges as strings
     *
     * @return
     */
    val allRanges: Array<String>
        get() = refs!!.refs

    /**
     * Returns the entire range this conditional format refers to  in Row[0]Col[0]Row[n]Col[n] format.
     *
     * @return
     */
    val encompassingRange: IntArray
        get() = refs!!.rowColBounds

    /**
     * get the bounding range of this conditional format
     *
     * @return
     */
    val boundingRange: String
        get() {
            val rowcols = refs!!.rowColBounds
            return ExcelTools.formatRangeRowCol(rowcols)
        }


    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)


    /**
     * returns EXML for the Conditional Format
     *
     *
     *
     * <ConditionalFormatting>
     * <Range>R12C2:R16C2</Range>
     * <Condition>
     * <Qualifier>Between</Qualifier>
     * <Value1>2</Value1>
     * <Value2>4</Value2>
     * <Format Style='color:#002060;font-weight:700;text-line-through:none;
    border:.5pt solid windowtext;background:#00B0F0'></Format>
    </Condition> *
    </ConditionalFormatting> *
     *
     * @return XML string for this record
     */
    val xml: String
        get() = getXML(false)

    /**
     * returns XMLSS for the Conditional Format
     *
     *
     *
     * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
     * <Range>R12C2:R16C2</Range>
     * <Condition>
     * <Qualifier>Between</Qualifier>
     * <Value1>2</Value1>
     * <Value2>4</Value2>
     * <Format Style='color:#002060;font-weight:700;text-line-through:none;
    border:.5pt solid windowtext;background:#00B0F0'></Format>
    </Condition> *
    </ConditionalFormatting> *
     *
     * @return XML string for this record
     */
    val xmlss: String
        get() = getXML(true)

    /**
     * set dirty flag to rebuild condfmt record
     * used when updated the underlying ranges wit
     */
    fun setDirty() {
        isdirty = true
    }

    /**
     * initialize the condfmt record
     *
     *
     * Please note that the sqref structure is not initialized in this location, but is required for cfmt functionality.
     *
     *
     * It happens on parse after worksheet is set
     */
    override fun init() {
        super.init()
        rw = 0
        ccf = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()     // SHOULD BE # cf's but appears to be 1+ ??
        grbit = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())   // SHOULD BE 1 to recalc but has been 3, 5, ...??
    }

    /**
     * As the init() call occurs before worksheet is set upon this conditionalformat record,
     * we have to initialze the references after init in order for referenceTracker to work correctly
     */
    fun initializeReferences() {
        data = this.getData()
        val sqrefCount = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt()).toInt()
        val sqrefdata = ByteArray(sqrefCount * 8)
        System.arraycopy(data!!, 14, sqrefdata, 0, sqrefdata.size)
        refs = DiscontiguousRefStruct(sqrefdata, this)
    }

    /**
     * add a new CF rule to this conditional format
     */
    fun addRule(c: Cf) {
        if (rules.indexOf(c) == -1) {
            rules.add(c)
        }
        c.condfmt = this
    }

    /**
     * update data for streaming
     *
     * @param loc
     */
    private fun updateRecord() {
        if (!isdirty) return
        // get the size of our output
        val outdata = ByteArray(refs!!.numRefs * 8 + 14)
        var tmp = ByteTools.shortToLEBytes(this.rules.size.toShort())
        var offset = 0
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]

        tmp = ByteTools.shortToLEBytes(grbit)
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]

        val rowcols = refs!!.rowColBounds
        tmp = ByteTools.shortToLEBytes(rowcols[0].toShort())
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]
        tmp = ByteTools.shortToLEBytes(rowcols[2].toShort())
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]
        tmp = ByteTools.shortToLEBytes(rowcols[1].toShort())
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]
        tmp = ByteTools.shortToLEBytes(rowcols[3].toShort())
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]

        tmp = ByteTools.shortToLEBytes(refs!!.numRefs.toShort())
        outdata[offset++] = tmp[0]
        outdata[offset++] = tmp[1]

        val sqrefbytes = refs!!.recordData
        System.arraycopy(sqrefbytes, 0, outdata, offset, sqrefbytes.size)
        this.setData(outdata)
    }

    /**
     * Add a location to the conditional format record, this
     * is a string representation that can be either a cell, ie "A1", or a range
     * ie "A1:A12";
     *
     * @param location string representing the added range
     */
    fun addLocation(location: String) {
        refs!!.addRef(location)
        isdirty = true
    }

    /**
     * Set this cf to a new enclosing cell range
     * This should only be used for inital creation of a conditional format
     * record or when all other internal ranges should be cleared as it removes
     * all others
     *
     * @param range
     */
    fun resetRange(range: String) {
        refs = DiscontiguousRefStruct(range, this)
        isdirty = true
    }

    /**
     * update the bytes
     */
    override fun preStream() {
        this.updateRecord()
    }

    /**
     * Add a cell range to this conditional format.
     *
     * @param string
     */
    private fun addCellRange(range: String) {
        refs!!.addRef(range)
        isdirty = true
        updateRecord()
    }

    /**
     * returns EXML (XMLSS) for the Conditional Format
     *
     *
     *
     * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
     * <Range>R12C2:R16C2</Range>
     * <Condition>
     * <Qualifier>Between</Qualifier>
     * <Value1>2</Value1>
     * <Value2>4</Value2>
     * <Format Style='color:#002060;font-weight:700;text-line-through:none;
    border:.5pt solid windowtext;background:#00B0F0'></Format>
    </Condition> *
    </ConditionalFormatting> *
     *
     * @return
     */
    fun getXML(useXMLSSNameSpace: Boolean): String {

        var ns = ""
        if (useXMLSSNameSpace)
            ns = "xmlns=\"urn:schemas-microsoft-com:office:excel\""

        val xml = StringBuffer("<ConditionalFormatting$ns>")
        // cf's
        val its = this.rules.iterator()
        while (its.hasNext()) {
            val c = its.next() as Cf
            xml.append(c.xml)
        }
        xml.append("</ConditionalFormatting>")
        return xml.toString()
    }

    /**
     * generate the proper OOXML to define this set of Conditional Formatting
     *
     * @return
     */
    fun getOOXML(bk: WorkBookHandle, priority: IntArray): String {
        this.updateRecord()
        val ooxml = StringBuffer()
        ooxml.append("<conditionalFormatting")
        if (this.refs != null) {
            ooxml.append(" sqref=\"")
            val refStrs = refs!!.refs
            for (i in refStrs.indices) {
                if (i > 0) ooxml.append(" ")
                ooxml.append(refStrs[i])
            }
            ooxml.append("\"")
        }
        ooxml.append(">")
        // cf's
        // NOTE:  cf.getDxfId/setDxfId links this conditional formatting rule with the proper incremental style
        // NOTE:  cfRules must have a valid dxfId or the output file will open with errors
        // NOTE:  for now, dxfs can only be saved from the original styles.xml;
        var dxfs: ArrayList<*>? = this.workBook!!.dxfs
        if (dxfs == null) {
            dxfs = ArrayList()
            this.workBook!!.dxfs = dxfs
        }
        if (rules != null) {
            for (i in rules.indices) {
                ooxml.append((rules.get(i) as Cf).getOOXML(bk, priority[0]++, dxfs))
            }
        }
        ooxml.append("</conditionalFormatting>")
        return ooxml.toString()
    }

    /**
     * Checks if the conditional format contains the row/col passed in
     *
     * @param rowColFromString
     * @return
     */
    operator fun contains(rowColFromString: IntArray): Boolean {
        return refs!!.containsReference(rowColFromString)
    }

    /**
     * clear out object referencse
     */
    override fun close() {
        super.close()
        refs = null
        if (rules != null) {
            for (i in rules.indices) {
                var cf: Cf? = rules.get(i)
                cf!!.close()
                cf = null

            }
        }
    }

    companion object {

        private val serialVersionUID = -7923448634000437926L

        /**
         * Create a Condfmt record & populate with prototype bytes
         *
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val cf = Condfmt()
                cf.opcode = XLSConstants.CONDFMT
                cf.setData(cf.PROTOTYPE_BYTES)
                cf.init()
                return cf
            }


        /**
         * OOXML conditionalFormatting (Conditional Formatting)
         * A Conditional Format is a format, such as cell shading or font color,
         * that a spreadsheet application can
         * automatically apply to cells if a specified condition is true.
         * This collection expresses conditional formatting rules
         * applied to a particular cell or range.
         *
         * parent:   worksheet
         * children: cfRule  (1 or more)
         * attributes: pivot (flag indicating this cf is assoc with a pivot table), sqref
         */


        /**
         * create one or more Data Validation records based on OOXML input
         */
        // TODO: finish pivot option, create Cf recs on each cfRule
        fun parseOOXML(xpp: XmlPullParser, wb: WorkBookHandle, bs: Boundsheet): Condfmt? {
            var condfmt: Condfmt? = null
            var dxfs: ArrayList<*>? = wb.workBook!!.dxfs
            if (dxfs == null) dxfs = ArrayList()    // shouldn't!
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "conditionalFormatting") {      // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (n == "sqref") {    // series of references
                                    condfmt = bs.createCondfmt("", wb)    //(Condfmt) Condfmt.getPrototype();
                                    condfmt!!.initializeReferences()
                                    val ranges = StringTool.splitString(v, " ")
                                    for (z in ranges.indices) {
                                        condfmt.addCellRange(bs.sheetName + "!" + ranges[z])
                                    }
                                } else if (n == "pivot") {
                                    // ???
                                }
                            }
                            // create a Cf record based upon cfRule info
                        } else if (tnm == "cfRule") {  // one or more
                            val cfRule = CfRule.parseOOXML(xpp).cloneElement() as CfRule
                            val cf = bs.createCf(condfmt!!)    // creates a new cf rule and links to the current condfmt
                            cf.setOperator(Cf.translateOperator(cfRule.operator))    // set the cf rule operator	(greater than, equals ...)
                            cf.setType(Cf.translateOOXMLType(cfRule.type))                // set the cf rule type (cell is, exrpression ...)
                            if (cf.type.toInt() == 3)
                            // containsText
                                cf.setContainsText(cfRule.containsText)
                            if (cfRule.formula1 != null)
                                cf.setCondition1(cfRule.formula1)
                            if (cfRule.formula2 != null)
                                cf.setCondition2(cfRule.formula2!!)
                            val dxfId = cfRule.dxfId
                            if (dxfId > -1) {    // it's not required to have a dxf
                                val dxf = dxfs!![dxfId] as Dxf    // dxf= differential format, contains the specific styles to define this cf rule
                                Cf.setStylePropsFromDxf(dxf, cf)
                                //	                        String dxfStyleString= dfx.getStyleProps();	// returns a string representation of the dxf or differential styles
                                //	                        Cf.setStylePropsFromString(dxfStyleString,cf);	// set the dxf styles to the cf rule
                            }
                            // original code that didn't input CfRules into Cf's just stored CfRule objects ...                        condfmt.cfRules.add((CfRule.parseOOXML(xpp).cloneElement()));
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "conditionalFormatting") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("Condfmt.parseOOXML: $e")
            }

            if (condfmt != null)
                bs.addConditionalFormat(condfmt)   // add this conditional format to the sheet
            return condfmt
        }
    }

}
