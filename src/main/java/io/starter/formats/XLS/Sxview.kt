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
import io.starter.formats.XLS.SxAddl.SxcView
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException
import java.util.ArrayList


/**
 * **SXVIEW B0h: This record contains top-level pivot table information.**<br></br>
 *
 * <pre>
 * ref (8 bytes): A Ref8U structure that specifies the PivotTable report body. For more information, see Location and Body.
 * rwFirstHead (2 bytes): An RwU structure that specifies the first row of the row area.
 * MUST be 1 if none of the axes are assigned in this PivotTable view.
 * Otherwise, the value MUST be greater than or equal to ref.rwFirst.
 *
 * rwFirstData (2 bytes): An RwU structure that specifies the first row of the data area.
 * MUST be 1 if none of the axes are assigned in this PivotTable view.
 * Otherwise, it MUST be equal to the value as specified by the following formula:
 *
 * rwFirstData = rwFirstHead + cDimCol
 *
 * colFirstData (2 bytes): A ColU structure that specifies the first column of the data area.
 * It MUST be 1 if none of the axes are assigned in this PivotTable view.
 * Otherwise, the value MUST be greater than or equal to ref.colFirst, and if the value of cDimCol or cDimData is not zero,
 * it MUST be less than or equal to ref.colLast.
 *
 * iCache (2 bytes): A signed integer that specifies the zero-based index of an SXStreamID record in the Globals substream.
 * MUST be greater than or equal to zero and less than the number of SXStreamID records in the Globals substream.
 *
 * reserved (2 bytes): MUST be zero, and MUST be ignored.
 *
 * sxaxis4Data (2 bytes): An SXAxis structure that specifies the default axis for the data field.
 * Either the sxaxis4Data.sxaxisRw field MUST be 1 or the sxaxis4Data.sxaxisCol field MUST be 1.
 * The sxaxis4Data.sxaxisPage field MUST be 0 and the sxaxis4Data.sxaxisData field MUST be 0.
 *
 * ipos4Data (2 bytes): A signed integer that specifies the row or column position for the data field in the PivotTable view.
 * The sxaxis4Data field specifies whether this is a row or column position.
 * MUST be greater than or equal to -1 and less than or equal to 0x7FFF. A value of -1 specifies the default position.
 *
 * cDim (2 bytes): A signed integer that specifies the number of pivot fields in the PivotTable view.
 * MUST equal the number of Sxvd records following this record.
 * MUST equal the number of fields in the associated PivotCache specified by iCache.
 *
 * cDimRw (2 bytes): An unsigned integer that specifies the number of fields on the row axis of the PivotTable view.
 * MUST be less than or equal to 0x7FFF. MUST equal the number of array elements in the SxIvd record in this PivotTable view that contain row items.
 *
 * cDimCol (2 bytes): An unsigned integer that specifies the number of fields on the column axis of the PivotTable view.
 * MUST be less than or equal to 0x7FFF.
 * MUST equal the number of array elements in the SxIvd record in this PivotTable view that contain column items.
 *
 * cDimPg (2 bytes): An unsigned integer that specifies the number of page fields in the PivotTable view.
 * MUST be less than or equal to 0x7FFF.
 * MUST equal the number of array elements in the SXPI record in this PivotTable view.
 *
 * cDimData (2 bytes): A signed integer that specifies the number of data fields in the PivotTable view.
 * MUST be greater than or equal to zero and less than or equal to 0x7FFF.
 * MUST equal the number of SXDI records in this PivotTable view.
 *
 * cRw (2 bytes): An unsigned integer that specifies the number of pivot lines in the row area of the PivotTable view.
 * MUST be less than or equal to 0x7FFF.
 * MUST equal the number of array elements in the first SXLI record in this PivotTable view.
 *
 * cCol (2 bytes): An unsigned integer that specifies the number of pivot lines in the column area of the PivotTable view.
 * MUST equal the number of array elements in the second SXLI record in this PivotTable view.
 *
 * A - fRwGrand (1 bit): A bit that specifies whether the PivotTable contains grand totals for rows.
 * MUST be 0 if none of the axes have been assigned in this PivotTable view.
 *
 * B - fColGrand (1 bit): A bit that specifies whether the PivotTable contains grand totals for columns.
 * MUST be 1 if none of the axes are assigned in this PivotTable view.
 *
 * C - unused1 (1 bit): Undefined and MUST be ignored.
 *
 * D - fAutoFormat (1 bit): A bit that specifies whether the PivotTable has AutoFormat applied.
 *
 * E - fAtrNum (1 bit): A bit that specifies whether the PivotTable has number AutoFormat applied.
 *
 * F - fAtrFnt (1 bit): A bit that specifies whether the PivotTable has font AutoFormat applied.
 *
 * G - fAtrAlc (1 bit): A bit that specifies whether the PivotTable has alignment AutoFormat applied.
 *
 * H - fAtrBdr (1 bit): A bit that specifies whether the PivotTable has border AutoFormat applied.
 *
 * I - fAtrPat (1 bit): A bit that specifies whether the PivotTable has pattern AutoFormat applied.
 *
 * J - fAtrProc (1 bit): A bit that specifies whether the PivotTable has width/height AutoFormat applied.
 *
 * unused2 (6 bits): Undefined and MUST be ignored.
 *
 * itblAutoFmt (2 bytes): An AutoFmt8 structure that specifies the PivotTable AutoFormat.
 * If the value of itblAutoFmt in the associated SXViewEx9 record is not 1, this field is overridden by the value of itblAutoFmt in the associated SXViewEx9.
 *
 * cchTableName (2 bytes): An unsigned integer that specifies the length, in characters, of stTable.
 * MUST be greater than or equal to zero and less than or equal to 0x00FF.
 *
 * cchDataName (2 bytes): An unsigned integer that specifies the length, in characters of stData.
 * MUST be greater than zero and less than or equal to 0x00FE.
 *
 * stTable (variable): An XLUnicodeStringNoCch structure that specifies the name of the PivotTable.
 * The length of this field is specified by cchTableName.
 *
 * stData (variable): An XLUnicodeStringNoCch structure that specifies the name of the data field.
 * The length of this field is specified by cchDataName.
 *
</pre> *
 */
class Sxview : XLSRecord(), XLSConstants {
    internal var rwFirst: Short = 0x0 // First Row of Pivot Table
    internal var rwLast: Short = 0x0 // Last Row of Pivot Table
    internal var colFirst: Short = 0x0 // First Column of Pivot Table
    internal var colLast: Short = 0x0 // Last Column of Pivot Table
    internal var rwFirstHead: Short = 0x0 // First Row containing Pivot Table headings
    internal var rwFirstData: Short = 0x0 // First Row containing Pivot Table data
    internal var colFirstData: Short = 0x0 // First Column containing Pivot Table data
    internal var iCache: Short = 0x0 // Index to the Cache
    //  offset 20, length 2 is reserved, must be 0 (zero)
    /**
     * retieves the default axis for the data field:
     *  * 1- row
     *  * 2= col
     *  * 4= page
     *  * 8= data
     */
    var sxaxis4Data: Short = 0x0
        internal set // Default axis for a data field
    internal var ipos4Data: Short = 0x0 // Default position for a data field
    internal var cDim: Short = 0x0 // Number of fields
    /**
     * returns the number of pivot fields on the ROW axis
     *
     * @return
     */
    var cDimRw: Short = 0x0
        internal set // Number of row fields;
    /**
     * returns the number of pivot fields on the COL axis
     *
     * @return
     */
    var cDimCol: Short = 0x0
        internal set // Number of column fields
    /**
     * returns the number of fields on the page axis
     *
     * @return
     */
    var cDimPg: Short = 0x0
        internal set // Number of page fields;
    /**
     * returns the number of pivot fields on the data axis
     *
     * @return
     */
    var cDimData: Short = 0x0
        internal set // Number of Data fields
    /**
     * returns the number of pivot items or lines on the ROW axis
     *
     * @return
     */
    var cRw: Short = 0x0
        internal set // Number of data rows
    /**
     * returns the number of pivot items or lines on the COL axis
     *
     * @return
     */
    var cCol: Short = 0x0
        internal set // Number of data columns
    internal var grbit1: Byte = 0x0 // if you dont know what this is....
    internal var grbit2: Byte = 0x0 // if you dont know what this is....
    internal var itblAutoFmt: Short = 0x0 // Index to the Pivot Table autoformat
    internal var cchName: Short = 0x0 // Length the Pivot Table name
    internal var cchData: Short = 0x0 // Length of the data field name
    internal var rgch: ByteArray// PivotTableName followed by the name of a data field

    /* The following member variables are populated from parsing the
       grbit field
    */
    internal var fRwGrand = false  // = 1 if the Pivot Table contains grand totals for rows
    /**
     * specifies whether the PivotTable contains grand totals for columns. MUST be 1 if none of the axes are assigned in this PivotTable view.
     */
    var fColGrand = false
        internal set // = 1 if the Pivot Table contains grand totals for Columns
    // bit 2, mask 0004h is reserved, must be 0 (zero)
    internal var fAutoFormat = false // = 1 if the Pivot table has an autoformat applied
    internal var fWH = false // = 1 if the width/height autoformat is applied
    internal var fFont = false // = 1 if the font autoformat is applied
    internal var fAlign = false // = 1 if the alignment autoformat is applied
    internal var fBorder = false // = 1 if the border autoformat is applied
    internal var fPattern = false // = 1 if the pattern autoformat is applied
    internal var fNumber = false // = 1 if the number autoformat is applied

    /*
        The following member variables are derived from the rgch field
    */
    internal var PivotTableName: String? = null
    internal var DataFieldName: String? = null

    private val subRecs = ArrayList()

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, /* ref */
            0, 0, /* rwfirstHead */
            0, 0, /* rwfirstData */
            0, 0, /* colfirstHead */
            0, 0, /* iCache */
            0, 0, /* reserved */
            1, 0, /* saxis4Data */
            -1, -1, /* ipos4Data */
            0, 0, /* cDim */
            0, 0, /* cDimRw */
            0, 0, /* cDimCol */
            0, 0, /* cDimPg */
            0, 0, /* cDimData */
            0, 0, /* cRw */
            0, 0, /* cCol */
            11, 2, /* grbit + reserved */
            1, 0, /* itblAutoFmt */
            0, 0, /* cchTableName */
            0, 0)/* cchDataName */


    /**
     * returns the number of pivot fields (==columns in the pivot table data range)
     *
     * @return
     */
    /**
     * sets the number of pivot fields (i.e. columns in the source data) for this pivot table
     * <br></br>NOTE: this method wipes any existing data of the pivot table
     *
     * @param s
     */
    // number of pivot fields MUST==# of SxVd records
    // remove ALL field-related records
    // reset variables
    // now add the required records for each field
    // for each pivot field (which goes on an axis)
    var nPivotFields: Short
        get() = cDim
        set(s) {
            cDim = s
            val b = ByteTools.shortToLEBytes(cDim)
            System.arraycopy(b, 0, data!!, 22, 2)
            run {
                var i = 0
                while (i < subRecs.size) {
                    val br = subRecs.get(i)
                    if (br.opcode != XLSConstants.SXEX) {
                        this.sheet!!.removeRecFromVec(br)
                        subRecs.removeAt(i)
                        i--
                    } else
                        break
                    i++
                }
            }
            cDimRw = 0
            cDimCol = 0
            cDimPg = 0
            cDimData = 0
            cRw = 0
            cCol = 0
            var zz = this.recordIndex + 1
            for (i in 0 until cDim) {
                val svd = Sxvd.prototype as Sxvd?
                svd!!.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(zz++, svd)
                this.subRecs.add(i * 2, svd)
                val svdex = SxVdEX.prototype as SxVdEX?
                svdex!!.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(zz++, svdex)
                this.subRecs.add(i * 2 + 1, svdex)
            }
        }

    var fwh: Boolean
        get() = fWH
        set(b) {
            if (b != fWH) {
                if (b == false) {
                    grbit1 = (grbit1 and 0xEF).toByte()
                } else {
                    grbit1 = (grbit1 or 0x10).toByte()
                }
                this.initGrbit()
            }
        }

    /**
     * return the name of the Pivot Table.
     */
    /**
     * Sets the name of the Pivot Table.
     */
    // also set associated qsitag pivot view name
    // find SXADDL_SxView_SxDID record and set PivotTableView name - must match this view tablename
    var tableName: String?
        get() = PivotTableName
        set(s) {
            PivotTableName = s
            this.buildRgch()
            val qsi = getSubRec(XLSConstants.QSISXTAG.toInt(), -1) as QsiSXTag?
            if (qsi != null)
                qsi.name = s
            for (i in subRecs.size - 1 downTo 1) {
                val br = subRecs.get(i)
                if (br.opcode == XLSConstants.SXADDL) {
                    if ((br as SxAddl).recordId === SxAddl.SxcView.sxdId) {
                        (br as SxAddl).setViewName(s)
                        break
                    }
                }
            }

        }

    /**
     * Sets the name of the Data field
     */
    var dataName: String?
        get() = DataFieldName
        set(s) {
            DataFieldName = s
            this.buildRgch()
        }

    override fun init() {
        super.init()
        if (this.length <= 0) {  // Is this record populated?
            if (DEBUGLEVEL > -1) Logger.logInfo("no data in SXVIEW")
        } else { // parse out all the fields
            rwFirst = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
            rwLast = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
            colFirst = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
            colLast = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
            rwFirstHead = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
            rwFirstData = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
            colFirstData = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
            iCache = ByteTools.readShort(this.getByteAt(14).toInt(), this.getByteAt(15).toInt())

            // 16 & 17 - reserved must be zero

            sxaxis4Data = ByteTools.readShort(this.getByteAt(18).toInt(), this.getByteAt(19).toInt())
            ipos4Data = ByteTools.readShort(this.getByteAt(20).toInt(), this.getByteAt(21).toInt())
            cDim = ByteTools.readShort(this.getByteAt(22).toInt(), this.getByteAt(23).toInt())
            cDimRw = ByteTools.readShort(this.getByteAt(24).toInt(), this.getByteAt(25).toInt())
            cDimCol = ByteTools.readShort(this.getByteAt(26).toInt(), this.getByteAt(27).toInt())
            cDimPg = ByteTools.readShort(this.getByteAt(28).toInt(), this.getByteAt(29).toInt())
            cDimData = ByteTools.readShort(this.getByteAt(30).toInt(), this.getByteAt(31).toInt())
            cRw = ByteTools.readShort(this.getByteAt(32).toInt(), this.getByteAt(33).toInt())
            cCol = ByteTools.readShort(this.getByteAt(34).toInt(), this.getByteAt(35).toInt())
            grbit1 = this.getByteAt(37)
            grbit2 = this.getByteAt(36)


            this.initGrbit() // note the manual hibyting
            itblAutoFmt = ByteTools.readShort(this.getByteAt(38).toInt(), this.getByteAt(39).toInt())
            cchName = ByteTools.readShort(this.getByteAt(40).toInt(), this.getByteAt(41).toInt())
            cchData = ByteTools.readShort(this.getByteAt(42).toInt(), this.getByteAt(43).toInt())
            val fullnamelen = cchName.toInt() + cchData.toInt()
            rgch = ByteArray(fullnamelen)
            var pos = 44
            if (cchName > 0) {
                //A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
                // 0x0  All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
                // 0x1  All the characters in the string are saved as double-byte characters in rgb.
                // reserved (7 bits): MUST be zero, and MUST be ignored.

                val encoding = this.getByteAt(pos++)

                val tmp = this.getBytesAt(pos, cchName * (encoding + 1))
                try {
                    if (encoding.toInt() == 0)
                        PivotTableName = String(tmp!!, XLSConstants.DEFAULTENCODING)
                    else
                        PivotTableName = String(tmp!!, XLSConstants.UNICODEENCODING)
                } catch (e: UnsupportedEncodingException) {
                    Logger.logInfo("encoding PivotTable name in Sxview: $e")
                }

                pos += cchName * (encoding + 1)
            }
            if (cchData > 0) {
                val encoding = this.getByteAt(pos++)
                val tmp = this.getBytesAt(pos, cchData * (encoding + 1))
                try {
                    if (encoding.toInt() == 0)
                        DataFieldName = String(tmp!!, XLSConstants.DEFAULTENCODING)
                    else
                        DataFieldName = String(tmp!!, XLSConstants.UNICODEENCODING)
                } catch (e: UnsupportedEncodingException) {
                    Logger.logInfo("encoding PivotTable name in Sxview: $e")
                }

            }
        }
        if (DEBUGLEVEL > 3)
            Logger.logInfo("SXVIEW: name:" + this.tableName + " iCache:" + iCache + " cDim:" + cDim + " cDimRw:" + cDimRw + " cDimCol:" + cDimCol + " cDimPg:" + cDimPg + " cDimData:" + cDimData + " cRw:" + cRw + " cCol:" + cCol + " datafieldname:" + DataFieldName)
    }

    /**
     * specifies the default axis for the data field:
     *  * 1- row
     *  * 2= col
     *  * 4= page
     *  * 8= data
     *
     * @param s
     */
    protected fun setAxis4Data(s: Short) {
        sxaxis4Data = s
        val b = ByteTools.shortToLEBytes(sxaxis4Data)
        System.arraycopy(b, 0, data!!, 18, 2)
    }

    /**
     * A signed integer that specifies the row or column position for the data field in the PivotTable view.
     * <br></br>see getSxaxis4Data to determine whether this is a row or column position
     *
     * @param s
     */
    protected fun setIpos4Data(s: Short) {
        ipos4Data = s
        val b = ByteTools.shortToLEBytes(ipos4Data)
        System.arraycopy(b, 0, data!!, 20, 2)
    }

    /**
     * A signed integer that specifies the row or column position for the data field in the PivotTable view.
     * <br></br>see getSxaxis4Data to determine whether this is a row or column position
     */
    fun getIpos4Data(): Short {
        return ipos4Data
    }

    /**
     * adds the pivot field corresponding to cache field at index fieldIndex to the desired axis
     * <br></br>A pivot field is a cache field that has been added to the pivot table
     * <br></br>Pivot fields are defined by SXVD and associated records
     * <br></br>the SXVD record stores which axis the field is on
     * <br></br>there are cDim pivot fields on the pivot table view
     *  * 1= row
     *  * 2= col
     *  * 4= page
     *  * 8= data
     *
     * @param axis
     * @param fieldIndex 0-based pivot field index
     * @see SxVd.AXIS_ constants
     */
    fun addPivotFieldToAxis(axis: Int, fieldIndex: Int): Sxvd {
        var zz = this.recordIndex + 1
        val sxvdex = getSubRec(XLSConstants.SXVDEX.toInt(), fieldIndex) as SxVdEX?    // end of last pivot field set (PIVOTVD rule)
        if (sxvdex != null)
            zz = sxvdex.recordIndex + 1

        val sxvd = Sxvd.prototype as Sxvd?    // for each pivot field (which goes on an axis)
        sxvd!!.setSheet(this.sheet)
        this.sheet!!.sheetRecs.add(zz++, sxvd)
        this.subRecs.add(cDim * 2, sxvd)
        val svdex = SxVdEX.prototype as SxVdEX?
        svdex!!.setSheet(this.sheet)
        this.sheet!!.sheetRecs.add(zz++, svdex)
        this.subRecs.add(cDim * 2 + 1, svdex)
        cDim++
        sxvd.setAxis(axis)
        return sxvd

        /*
    	Sxvd sxvd= (Sxvd) getSubRec(SXVD, fieldIndex);
    	if (sxvd!=null) {
    		// got it
    		sxvd.setAxis(axis);
        	return sxvd;
        }
        return null;*/
    }

    /**
     * adds a pivot field  (0-based index) to the ROW axis
     *
     * @param fieldNumber
     */
    fun addRowField(fieldNumber: Int) {
        cDimRw++    // MUST==# row elements in SxIVD
        val b = ByteTools.shortToLEBytes(cDimRw)
        System.arraycopy(b, 0, data!!, 24, 2)
        var sxivd = getSubRec(XLSConstants.SXIVD.toInt(), if (cDimCol > 0) 1 else 0) as Sxivd?
        if (sxivd == null) {
            val zz = getPivotRecordInsertionIndexes(XLSConstants.SXIVD.toInt(), 1, -1)
            if (zz > 0) { // should!!!
                sxivd = Sxivd.prototype as Sxivd?
                sxivd!!.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxivd)
                this.subRecs.add(zz, sxivd)
            }
        }
        sxivd!!.addField(fieldNumber)
    }

    /**
     * adds a pivot field (0-based index) to the COL axis
     */
    fun addColField(fieldNumber: Int) {
        cDimCol++
        val b = ByteTools.shortToLEBytes(cDimCol)
        System.arraycopy(b, 0, data!!, 26, 2)
        var sxivd = getSubRec(XLSConstants.SXIVD.toInt(), 0) as Sxivd?
        if (sxivd == null) {
            val zz = getPivotRecordInsertionIndexes(XLSConstants.SXIVD.toInt(), 0, -1)
            if (zz > 0) { // should!!
                sxivd = Sxivd.prototype as Sxivd?
                sxivd!!.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxivd)
                this.subRecs.add(zz, sxivd)
            }
        }
        sxivd!!.addField(fieldNumber)
    }

    /**
     * adds a pivot field to the page axis
     *
     * @param strref
     */
    fun addPageField(fieldIndex: Int, itemIndex: Int) {
        cDimPg++
        val b = ByteTools.shortToLEBytes(cDimPg)
        System.arraycopy(b, 0, data!!, 28, 2)
        var sxpi = getSubRec(XLSConstants.SXPI.toInt(), -1) as SxPI?
        if (sxpi == null) {
            val zz = getPivotRecordInsertionIndexes(XLSConstants.SXPI.toInt(), -1, -1)
            if (zz > 0) { // should!!!
                sxpi = SxPI.prototype as SxPI?
                sxpi!!.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxpi)
                this.subRecs.add(zz, sxpi)
            }
        }
        sxpi!!.addPageField(fieldIndex, itemIndex)
    }

    /**
     * adds a pivot field to the DATA axis
     *
     * @param fieldIndex
     * @param aggregateFunction
     */
    fun addDataField(fieldIndex: Int, aggregateFunction: String, name: String) {
        cDimData++
        val b = ByteTools.shortToLEBytes(cDimData)
        System.arraycopy(b, 0, data!!, 30, 2)
        var sxdi: SxDI? = null
        /*        SxDI sxdi= (SxDI) getSubRec(SXDI, -1);
        if (sxdi==null) {           */
        val zz = getPivotRecordInsertionIndexes(XLSConstants.SXDI.toInt(), -1, -1)
        if (zz > 0) { // should!!!
            sxdi = SxDI.prototype as SxDI?
            sxdi!!.setSheet(this.sheet)
            this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxdi)
            this.subRecs.add(zz, sxdi)
        }
        /*	    }*/
        sxdi!!.addDataField(fieldIndex, aggregateFunction, name)
    }

    /**
     * adds a pivot item to the end of the list of items on this axis (ROW, COL, DATA or PAGE)
     *
     * @param axis      Axis int: ROW, COL, DATA or PAGE
     * @param itemType  one of:
     *  * 0x0000		itmtypeData	    A data value
     *  * 0x0001	    itmtypeDEFAULT	Default subtotal for the pivot field
     *  * 0x0002	    itmtypeSUM	    Sum of values in the pivot field
     *  * 0x0003		itmtypeCOUNTA	Count of values in the pivot field
     *  * 0x0004	    itmtypeAVERAGE	Average of values in the pivot field
     *  * 0x0005	    itmtypeMAX	    Max of values in the pivot field
     *  * 0x0006	    itmtypeMIN	    Min of values in the pivot field
     *  * 0x0007	    itmtypePRODUCT  Product of values in the pivot field
     *  * 0x0008	    itmtypeCOUNT	Count of numbers in the pivot field
     *  * 0x0009	    itmtypeSTDEV	Statistical standard deviation (estimate) of the pivot field
     *  * 0x000A	    itmtypeSTDEVP	Statistical standard deviation (entire population) of the pivot field
     *  * 0x000B		itmtypeVAR	    Statistical variance (estimate) of the pivot field
     *  * 0x000C	    itmtypeVARP	    Statistical variance (entire population) of the pivot field
     * @param cacheItem A cache item index in the cache field associated with the pivot field, as specified by Cache Items.
     */
    fun addPivotItem(axis: Sxvd, itemType: Int, cacheItem: Int) {
        val n = axis.numItems        // Axis record = Sxvd.  Identifies pivot field on axis. Pivot items records follow until SXVDEX.
        axis.numItems = n + 1
        val axisIndex = getSubRecIndex(axis)
        val zz = getPivotRecordInsertionIndexes(XLSConstants.SXVI.toInt(), n, axisIndex)    // get LAST pivot field on axis
        val sxvi = Sxvi.prototype as Sxvi?                // pivot item record
        sxvi!!.setItemType(itemType)
        sxvi.setCacheItem(cacheItem)
        sxvi.setSheet(this.sheet)
        this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxvi)
        subRecs.add(zz, sxvi)
        /*if (axis.getAxisType()==Sxvd.AXIS_ROW) {
     	} else if (axis.getAxisType()==Sxvd.AXIS_COL) {
     	} else if (axis.getAxisType()==Sxvd.AXIS_PAGE) {

     	} else if (axis.getAxisType()==Sxvd.AXIS_DATA) {

     	}*/
        if (cacheItem != -1) {
            val pc = this.workBook!!.pivotCache    // TODO should this be here or in Sxstream?
            pc!!.addCacheItem(iCache.toInt(), cacheItem)    // adds the cache item to the "used" list, as it were
        }
    }

    /**
     * sets the number of pivot items or lines on the ROW axis
     *
     * @param strref
     */
    fun addPivotLineToROWAxis(repeat: Int, nLines: Int, type: Int, indexes: ShortArray) {
        cRw++
        val b = ByteTools.shortToLEBytes(cRw)
        System.arraycopy(b, 0, data!!, 32, 2)
        /**
         * If the value of either of the cRw or cCol fields of the associated SxView is greater than zero,
         * then two records of this type MUST exist in the file for the associated SxView.
         * The first record contains row area pivot lines and the second record contains column area pivot lines.		 */
        var sxli = getSubRec(XLSConstants.SXLI.toInt(), 0) as Sxli?
        if (sxli == null) {
            val zz = getPivotRecordInsertionIndexes(XLSConstants.SXLI.toInt(), 0, -1)
            if (zz > 0) { // should!!!
                sxli = Sxli.getPrototype(this.workBook) as Sxli
                sxli.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxli)
                this.subRecs.add(zz, sxli)
                sxli.addField(repeat, nLines, type, indexes)
                sxli = Sxli.getPrototype(this.workBook) as Sxli
                sxli.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(subRecs.get(zz + 1).recordIndex, sxli)
                this.subRecs.add(zz + 1, sxli)
            }
        } else
            sxli.addField(repeat, nLines, type, indexes)
    }

    fun hasRowPivotItemsRecord(): Boolean {
        return getSubRec(XLSConstants.SXLI.toInt(), 0) != null
    }


    /**
     * set the number of pivot items or lines on the COL axis
     *
     * @param strref
     */
    fun addPivotLineToCOLAxis(repeat: Int, nLines: Int, type: Int, indexes: ShortArray) {
        cCol++
        val b = ByteTools.shortToLEBytes(cCol)
        System.arraycopy(b, 0, data!!, 34, 2)
        var sxli = getSubRec(XLSConstants.SXLI.toInt(), 1) as Sxli?
        if (sxli == null) {
            val zz = getPivotRecordInsertionIndexes(XLSConstants.SXLI.toInt(), 0, -1)
            if (zz == -1) { // shouldn't!
                sxli = Sxli.prototype as Sxli?
                sxli!!.setSheet(this.sheet)
                this.sheet!!.sheetRecs.add(subRecs.get(zz).recordIndex, sxli)
                this.subRecs.add(zz, sxli)
            }
        }
        sxli!!.addField(repeat, nLines, type, indexes)
    }

    /**
     * iCache links this pivot table to a pivot data cache
     *
     * @param s
     */
    fun setICache(s: Short) {
        iCache = s
        val b = ByteTools.shortToLEBytes(iCache)
        System.arraycopy(b, 0, data!!, 14, 2)
    }

    /**
     * iCache links this pivot table to a pivot data cache
     */
    fun getICache(): Short {
        return iCache
    }


    /**
     * Init the grbit, populate the member variables of the SXVIEW object.
     * This can be called both at initial init, and later to reup the boolean fields
     * after changes to the grbit.
     */
    private fun initGrbit() {
        fRwGrand = grbit1 and 0x1 == 0x1
        fColGrand = grbit1 and 0x2 == 0x2
        fAutoFormat = grbit1 and 0x8 == 0x8
        fWH = grbit1 and 0x10 == 0x10
        fFont = grbit1 and 0x20 == 0x20
        fAlign = grbit1 and 0x40 == 0x40
        fBorder = grbit1 and 0x80 == 0x80
        fPattern = grbit2 and 0x1 == 0x1
        fNumber = grbit2 and 0x2 == 0x2
        this.getData()[36] = grbit2
        this.getData()[37] = grbit1
    }

    /**
     * store associated records for ease of lookup
     *
     * @param r
     */
    fun addSubrecord(r: BiffRec) {
        subRecs.add(r)
    }
    /*
        The following methods all change or get grbit fields
    */

    /**
     * specifies whether the PivotTable contains grand totals for rows.
     * MUST be 0 if none of the axes have been assigned in this PivotTable view.
     */
    fun setFRwGrand(b: Boolean) {
        if (b != fRwGrand) {
            if (b == false) {
                grbit1 = (grbit1 and 0xFE).toByte()
            } else {
                grbit1 = (grbit1 or 0x1).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable contains grand totals for rows.
     * MUST be 0 if none of the axes have been assigned in this PivotTable view.
     */
    fun getFRwGrand(): Boolean {
        return fRwGrand
    }

    /**
     * specifies whether the PivotTable contains grand totals for columns.
     * MUST be 1 if none of the axes are assigned in this PivotTable view.
     *
     * @param b
     */
    fun setColGrand(b: Boolean) {
        if (b != fColGrand) {
            if (b == false) {
                grbit1 = (grbit1 and 0xFD).toByte()
            } else {
                grbit1 = (grbit1 or 0x2).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable has AutoFormat applied.
     *
     * @param b
     */
    fun setFAutoFormat(b: Boolean) {
        if (b != fAutoFormat) {
            if (b == false) {
                grbit1 = (grbit1 and 0xFB).toByte()
            } else {
                grbit1 = (grbit1 or 0x4).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable has AutoFormat applied.
     *
     * @return
     */
    fun getFAutoFormat(): Boolean {
        return fAutoFormat
    }

    /**
     * specifies whether the PivotTable has font AutoFormat applied.
     *
     * @param b
     */
    fun setFFont(b: Boolean) {
        if (b != fFont) {
            if (b == false) {
                grbit1 = (grbit1 and 0xDF).toByte()
            } else {
                grbit1 = (grbit1 or 0x20).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable has font AutoFormat applied.
     *
     * @return
     */
    fun getFFont(): Boolean {
        return fFont
    }

    /**
     * specifies whether the PivotTable has alignment AutoFormat applied.
     *
     * @param b
     */
    fun setFAlign(b: Boolean) {
        if (b != fAlign) {
            if (b == false) {
                grbit1 = (grbit1 and 0xBF).toByte()
            } else {
                grbit1 = (grbit1 or 0x40).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable has alignment AutoFormat applied.
     *
     * @return
     */
    fun getFAlign(): Boolean {
        return fAlign
    }

    /**
     * specifies whether the PivotTable has border AutoFormat applied.
     *
     * @param b
     */
    fun setFBorder(b: Boolean) {
        if (b != fBorder) {
            if (b == false) {
                grbit1 = (grbit1 and 0x7F).toByte()
            } else {
                grbit1 = (grbit1 or 0x80).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable has border AutoFormat applied.
     *
     * @return
     */
    fun getFBorder(): Boolean {
        return fBorder
    }

    /**
     * specifies whether the PivotTable has pattern AutoFormat applied.
     *
     * @param b
     */
    fun setFPattern(b: Boolean) {
        if (b != fPattern) {
            if (b == false) {
                grbit2 = (grbit2 and 0xFE).toByte()
            } else {
                grbit2 = (grbit2 or 0x1).toByte()
            }
            this.initGrbit()
        }
    }

    /**
     * specifies whether the PivotTable has pattern AutoFormat applied.
     *
     * @return
     */
    fun getFPattern(): Boolean {
        return fPattern
    }

    /**
     * @param b
     */
    fun setFNumber(b: Boolean) {
        if (b != fNumber) {
            if (b == false) {
                grbit2 = (grbit2 and 0xFE).toByte()
            } else {
                grbit2 = (grbit2 or 0x1).toByte()
            }
            this.initGrbit()
        }
    }

    fun getFNumber(): Boolean {
        return fNumber
    }


    /**
     * sets the location and size of this pivot table view
     *
     * @param range
     */
    fun setLocation(range: String) {
        val rc = ExcelTools.getRangeRowCol(range)
        setRwFirst(rc[0].toShort())
        setColFirst(rc[1].toShort())
        setRwLast(rc[2].toShort())
        setColLast(rc[3].toShort())
    }
    /*
        Set and get methods,  these all basically do the same thing, update the
        member variable, then update the body data with the new information.  Simple, but large XLSRecord
    */

    /**
     * sets the first row of the pivot table
     */
    fun setRwFirst(s: Short) {
        rwFirst = s
        val b = ByteTools.shortToLEBytes(rwFirst)
        System.arraycopy(b, 0, data!!, 0, 2)
    }

    /**
     * retrieves the first row of the pivot table
     */
    fun getRwFirst(): Short {
        return rwFirst
    }

    /**
     * sets the last row of the pivot table
     */
    fun setRwLast(s: Short) {
        rwLast = s
        val b = ByteTools.shortToLEBytes(rwLast)
        System.arraycopy(b, 0, data!!, 2, 2)
    }

    /**
     * retieves the first row of the pivot table
     */
    fun getRwLast(): Short {
        return rwLast
    }

    /**
     * sets the first column of the pivot table
     */
    fun setColFirst(s: Short) {
        colFirst = s
        val b = ByteTools.shortToLEBytes(colFirst)
        System.arraycopy(b, 0, data!!, 4, 2)
    }

    /**
     * retrieves the first column of the pivot table
     */
    fun getColFirst(): Short {
        return colFirst
    }

    /**
     * sets the last column of the pivot table
     */
    fun setColLast(s: Short) {
        colLast = s
        val b = ByteTools.shortToLEBytes(colLast)
        System.arraycopy(b, 0, data!!, 6, 2)
    }

    /**
     * sets the last column of the pivot table
     */
    fun getColLast(): Short {
        return colLast
    }

    fun setRwFirstHead(s: Short) {
        rwFirstHead = s
        val b = ByteTools.shortToLEBytes(rwFirstHead)
        System.arraycopy(b, 0, data!!, 8, 2)
    }

    fun getRwFirstHead(): Short {
        return rwFirstHead
    }

    fun setRwFirstData(s: Short) {
        rwFirstData = s
        val b = ByteTools.shortToLEBytes(rwFirstData)
        System.arraycopy(b, 0, data!!, 10, 2)
    }

    fun getRwFirstData(): Short {
        return rwFirstData
    }

    fun setColFirstData(s: Short) {
        colFirstData = s
        val b = ByteTools.shortToLEBytes(colFirstData)
        System.arraycopy(b, 0, data!!, 12, 2)
    }

    fun getColFirstData(): Short {
        return colFirstData
    }

    /*
        Probably want to use the individual grbit options instead of this, but
        hey,  I'm being thorough.
    */
    fun setGrbit(s: Short) {
        val b = ByteTools.shortToLEBytes(s)
        grbit2 = b[0]
        grbit1 = b[1]
        System.arraycopy(b, 0, data!!, 36, 2)
        this.initGrbit()
    }


    fun setItblAutoFmt(s: Short) {
        itblAutoFmt = s
        val b = ByteTools.shortToLEBytes(itblAutoFmt)
        System.arraycopy(b, 0, data!!, 38, 2)
    }

    fun getItblAutoFmt(): Short {
        return itblAutoFmt
    }

    /*
        Builds a new rgch field, changes the length of the entire record....
    */
    private fun buildRgch() {
        var data = ByteArray(44)
        System.arraycopy(this.getData()!!, 0, data, 0, 44)
        var strbytes = ByteArray(0)
        var databytes = ByteArray(0)
        try {
            if (PivotTableName != null)
                strbytes = PivotTableName!!.toByteArray(charset(XLSConstants.DEFAULTENCODING))
            if (DataFieldName != null)
                databytes = DataFieldName!!.toByteArray(charset(XLSConstants.DEFAULTENCODING))
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("encoding pivot table name in SXVIEW: $e")
        }

        //update the lengths:
        cchName = strbytes.size.toShort()
        var nm = ByteTools.shortToLEBytes(cchName)
        data[40] = nm[0]
        data[41] = nm[1]

        cchData = databytes.size.toShort()
        nm = ByteTools.shortToLEBytes(cchData)
        data[42] = nm[0]
        data[43] = nm[1]

        // now append variable-length string data
        val newrgch = ByteArray(cchName.toInt() + cchData.toInt() + 2)    // account for encoding bytes
        System.arraycopy(strbytes, 0, newrgch, 1, cchName.toInt())
        System.arraycopy(databytes, 0, newrgch, cchName + 2, cchData.toInt())

        data = ByteTools.append(newrgch, data)
        this.setData(data)
    }


    /**
     * utility method to retrieve the nth pivot table subrecord (SxVd, SxPI, SxDI ...)
     *
     * @param opcode opcode to look up
     * @param index  if > -1, the index or occurrence of the record to return
     * @return
     */
    private fun getSubRec(opcode: Int, index: Int): BiffRec? {
        var j = 0
        for (i in subRecs.indices) {
            val br = subRecs.get(i)
            if (br.opcode.toInt() == opcode) {
                if (index == -1 || j++ == index)
                    return br
            }
        }
        return null
    }

    /**
     * utility method to retrieve the subrecord index of the desired pivot table record
     *
     * @param br pivot table record to look up
     * @return
     */
    private fun getSubRecIndex(br: BiffRec): Int {
        var i = 0
        while (i < subRecs.size) {
            if (br === subRecs.get(i))
                break
            i++
        }
        return i
    }

    /**
     * utility method to retrieve a pivot table subrecord (SxVd, SxPI, SxDI ...) index
     *
     * @param opcode opcode to look up
     * @param index  if > -1, the index or occurrence of the record to return
     * @return index in subrecords of desired record
     */
    private fun getSubRecIndex(opcode: Int, index: Int): Int {
        var j = 0
        for (i in subRecs.indices) {
            val br = subRecs.get(i)
            if (br.opcode.toInt() == opcode) {
                if (index == -1 || j++ == index)
                    return i
            }
        }
        return -1
    }

    /**
     * i
     * finds the proper sheet rec insertion index for the desired opcode
     * as well as the proper insertion index into the subrecord array
     * <br></br>this method contains alot of intelligence regarding the order of pivot table records
     *
     * @param opcode one of SXVD, SXIVD, SXPI, SXDI, SXLI, or SXVDEX
     * @param index  either 0 (for ROW records) or 1 (for COL records)
     * @return int[]    {sheet record index, subrecord index} or -1's
     */
    private fun getPivotRecordInsertionIndexes(opcode: Int, index: Int, pivotFieldIndex: Int): Int {
        var i: Int
        var j = 0
        if (pivotFieldIndex < 0) {
            i = subRecs.size - 1
            while (i >= 0) {
                val br = subRecs.get(i)
                val bropcode = br.opcode.toInt()
                if (bropcode == opcode) { //
                    if (j < index) {
                        j++
                        i--
                        continue        // haven't found correct record jet
                    }
                    break
                } else if (opcode == XLSConstants.SXLI.toInt() && (bropcode == XLSConstants.SXDI.toInt() || bropcode == XLSConstants.SXPI.toInt() || bropcode == XLSConstants.SXIVD.toInt() || bropcode == XLSConstants.SXVDEX.toInt())) {
                    break
                } else if (opcode == XLSConstants.SXDI.toInt() && (bropcode == XLSConstants.SXPI.toInt() || bropcode == XLSConstants.SXIVD.toInt() || bropcode == XLSConstants.SXVDEX.toInt())) {
                    break
                } else if (opcode == XLSConstants.SXPI.toInt() && (bropcode == XLSConstants.SXIVD.toInt() || bropcode == XLSConstants.SXVDEX.toInt())) {
                    break
                } else if (opcode == XLSConstants.SXIVD.toInt() && bropcode == XLSConstants.SXVDEX.toInt()) {
                    break
                } else if (opcode == XLSConstants.SXVD.toInt() && bropcode == XLSConstants.SXEX.toInt()) {
                    break
                }
                i--
            }
            i++    // counter reverse order
        } else {    // only lookup within the current pivot field
            i = pivotFieldIndex + 1
            while (i < subRecs.size) {
                val br = subRecs.get(i)
                val bropcode = br.opcode.toInt()
                if (bropcode == opcode) { //
                    if (j < index) {
                        j++
                        i++
                        continue        // haven't found correct record jet
                    }
                    break
                } else if (bropcode == XLSConstants.SXVDEX.toInt()) {
                    break
                }
                i++
            }
        }
        return i
        /*	    if (i >= 0) {
	    	return new int[] {subRecs.get(i+j).getRecordIndex(), i};
	    }
	    return new int[] {-1, i};*/
    }


    /**
     * creates the basic, default records necessary to define a pivot table
     *
     * @param sheet string sheetname where pivot table is located
     * @return arraylist of records
     */
    fun addInitialRecords(sheet: Boundsheet): ArrayList<*> {
        /*
         * basic blank pivot table:
         * SXVIEW
         *  	after SXVIEW pivot field and item records will go:
         *  		for each pivot field:  SxVd, SxVi* n items, SxVDEx
         *  	SXIVD if cDimRw > 0
         *  	SXIVD if cDimCol > 0
         *  	SXPI if cDimPg > 0
         *  	SXDI if cDimData > 0
         *  	SXLI if cRw > 0
         *  	SXLI if cCol > 0
         * SXEX
         * QXISTAG
         * SXADDL records
         *
         */
        val initialrecs = ArrayList()
        this.setSheet(sheet)
        this.workBook = sheet.workBook
        // SXEX
        val sxex = SxEX.prototype as SxEX?
        addInit(initialrecs, sxex, sheet)
        // QSISXTAG
        val qsi = QsiSXTag.prototype as QsiSXTag?
        addInit(initialrecs, qsi, sheet)
        val sxv = SxVIEWEX9.prototype as SxVIEWEX9?
        addInit(initialrecs, sxv, sheet)
        // SXADDLs -required SxAddl records: stores additional PivotTableView info of a variety of types
        //		byte[] b= ByteTools.cLongToLEBytes(sid);      	b= ByteTosols.append(new byte[] {0, 0}, b); // add 2 reserved bytes
        var sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdId.sxd().toInt(), null)
        addInit(initialrecs, sa, sheet)
        //SXADDL_sxcView: record=sxdVer10Info data:[1, 65, 0, 0, 0, 0]
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdVer10Info.sxd().toInt(), null)
        addInit(initialrecs, sa, sheet)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdTableStyleClient.sxd().toInt(), byteArrayOf(0, 0, 0, 0, 0, 0, 51, 0, 0, 0))
        addInit(initialrecs, sa, sheet)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdVerUpdInv.sxd().toInt(), byteArrayOf(2, 0, 0, 0, 0, 0))
        addInit(initialrecs, sa, sheet)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdVerUpdInv.sxd().toInt(), byteArrayOf(-1, 0, 0, 0, 0, 0))
        addInit(initialrecs, sa, sheet)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdEnd.sxd().toInt(), null)
        addInit(initialrecs, sa, sheet)
        return initialrecs
    }

    /**
     * utility function to properly add a Pivot Table View subrec
     *
     * @param initialrecs
     * @param rec
     * @param addToInitRecords
     * @param sheet
     */
    private fun addInit(initialrecs: ArrayList<*>, rec: XLSRecord, sheet: Boundsheet) {
        initialrecs.add(rec)
        rec.setSheet(sheet)
        rec.workBook = this.workBook
        this.addSubrecord(rec)
    }

    override fun toString(): String {
        return "SXVIEW: name:" + this.tableName + " iCache:" + iCache + " cDim:" + cDim + " cDimRw:" + cDimRw + " cDimCol:" + cDimCol +
                " cDimPg:" + cDimPg + " cDimData:" + cDimData +
                " cRw:" + cRw + " cCol:" + cCol +
                " datafieldname:" + DataFieldName
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2639291289806138985L

        val prototype: XLSRecord?
            get() {
                val sx = Sxview()
                sx.opcode = XLSConstants.SXVIEW
                sx.setData(sx.PROTOTYPE_BYTES)
                sx.init()
                return sx
            }
    }
}

