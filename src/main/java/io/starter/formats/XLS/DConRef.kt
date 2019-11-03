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

import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException

/**
 * DConRef	0x51
 *
 *
 * The DConRef record specifies a range in this workbook or in an external
 * workbook that is a data source for a PivotTable or a data source for the data
 * consolidation settings of the associated sheet. If the range specified is in
 * an external workbook this record also specifies the path to the external
 * workbook.
 *
 *
 * ref (6 bytes): A RefU structure that specifies the range. If this record is part of an SXTBL production as specified in the
 * Globals Substream ABNF and this field has a rwFirst equal to 0 and a rwLast equal to 16383, this reference specifies all rows
 * within the columns specified by colFirst and colLast.
 *
 *
 * cchFile (2 bytes): An unsigned integer that specifies the count of characters in stFile. MUST be greater than or equal to 0x0002.
 *
 *
 * stFile (variable): A DConFile structure that specifies the workbook and sheet that contains the range specified in the ref field.
 *
 *
 * unused (variable): An array of bytes that is unused and MUST be ignored. MUST exist if and only if stFile specifies a self reference
 * (the value of stFile.stFile.rgb[0] is 2). If the value stFile.stFile.fHighByte is 0 the size of this array is 1.
 * If the value of stFile.stFile.fHighByte is 1 the size of this array is 2.
 *
 *
 *
 *
 * The RefU structure specifies a range of cells on the sheet.
 * rwFirst (2 bytes): A RwU structure that specifies the first row in the range. The value MUST be less than or equal to rwLast.
 * rwLast (2 bytes): A RwU structure that specifies the last row in the range.
 * colFirst (1 byte): A ColByteU structure that specifies the first column in the range. The value MUST be less than or equal to colLast.
 * colLast (1 byte): A ColByteU structure that specifies the last column in the range.
 *
 *
 * the DConFile structure specifies the workbook file or workbook file and sheet that contain a data source range.
 * This structure is used by the DConBin, DConRef and DConName records.
 * stFile (variable): An XLUnicodeStringNoCch that specifies the workbook file or workbook file and sheet that contain the range specified in the DConBin, DConRef or DConName record.
 * MUST be a string that conforms to the following ABNF grammar:
 * dcon-file = external-virt-path / self-reference
 * external-virt-path = volume / unc-volume / rel-volume  / transfer-protocol / startup / alt-startup / library /  simple-file-path-dcon
 * simple-file-path-dcon = %x0001 file-path
 * self-reference = %x0002 sheet-name
 * See VirtualPath for the definition of the volume, unc-volume, rel-volume, transfer-protocol, startup, alt-startup, library, file-path and sheet-name rules used in the ABNF grammar.  Note that the volume, unc-volume, rel-volume, transfer-protocol, startup, alt-startup, library, and file-path rules specify that an optional sheet name can be included.
 * If this structure is contained in a DConName or DConBin record and the defined name has a workbook scope, then this string MUST satisfy the external-virt-path rule and MUST NOT specify a sheet name.  Otherwise a sheet name MUST be specified.
 *
 *
 *
 *
 *
 *
 * "http://www.extentech.com">Extentech Inc.
 */
class DConRef : XLSRecord(), XLSConstants {
    private var rwFirst: Short = 0
    private var rwLast: Short = 0
    private var colFirst: Short = 0
    private var colLast: Short = 0
    private var cchFile: Short = 0
    private var fileName: String? = null
    private var refType: Byte = 0

    /**
     * returns the source range for a pivot table in r0c0r1c form
     *
     * @return int[]
     */
    val range: IntArray
        get() = intArrayOf(rwFirst.toInt(), colFirst.toInt(), rwLast.toInt(), colLast.toInt())

    /**
     * if this is a self-referentail data source i.e. in same workbook, return souce sheet name
     *
     * @return
     */
    /**
     * sets the source sheet for the pivot table
     *
     * @param sheetName
     */
    // cch
    // encoding
    // self-reference flag
    var sourceSheet: String?
        get() {
            if (refType.toInt() != 2)
                Logger.logWarn("External Data Sources are not supported")
            return fileName
        }
        set(sheetName) {
            cchFile = (sheetName.length.toShort() + 1).toShort()
            fileName = sheetName
            var data = ByteArray(10)
            System.arraycopy(this.getData()!!, 0, data, 0, 6)
            val b = ByteTools.shortToLEBytes(cchFile)
            data[6] = b[0]
            data[7] = b[1]
            data[8] = 0
            data[9] = 0x2
            try {
                data = ByteTools.append(sheetName.toByteArray(charset(XLSConstants.DEFAULTENCODING)), data)
                data = ByteTools.append(byteArrayOf(0), data)
            } catch (e: UnsupportedEncodingException) {
            }

            this.setData(data)
        }

    /**
     * returns the cell range for this pivot cache
     *
     * @return
     */
    /**
     * sets the cell range for the pivot cache
     *
     * @param cr
     */
    //this.getWorkBook());
    var cellRange: CellRange?
        get() {
            val range = fileName + "!" + ExcelTools.formatLocation(intArrayOf(rwFirst.toInt(), colFirst.toInt(), rwLast.toInt(), colLast.toInt()))
            try {
                return CellRange(range, null)
            } catch (e: CellNotFoundException) {
            }

            return null
        }
        set(cr) = try {
            val rc = cr.rangeCoords
            setRange(rc, cr.getSheet()!!.sheetName)
        } catch (e: CellNotFoundException) {
        }

    /**
     * init method
     */
    override fun init() {
        super.init()
        rwFirst = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        rwLast = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        colFirst = this.getByteAt(4).toShort()
        colLast = this.getByteAt(5).toShort()
        cchFile = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        if (cchFile > 0) {
            //A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
            // 0x0  All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
            // 0x1  All the characters in the string are saved as double-byte characters in rgb.
            // reserved (7 bits): MUST be zero, and MUST be ignored.

            val encoding = this.getByteAt(8)
            refType = this.getByteAt(9)    // 1= simple-file-path-dcon 2= self-reference

            if (refType.toInt() != 2)
            // TODO: handle external refs ...
                Logger.logWarn("PivotTable: External Data Sources are not supported")
            val tmp = this.getBytesAt(10, (cchFile - 1) * (encoding + 1))
            try {
                if (encoding.toInt() == 0)
                    fileName = String(tmp!!, XLSConstants.DEFAULTENCODING)
                else
                    fileName = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding PivotTable name in DCONREF: $e")
            }

        }
        if (DEBUGLEVEL > 3)
            Logger.logInfo("DCONREF: rwFirst:$rwFirst rwLast:$rwLast colFirst:$colFirst colLast:$colLast cchFile:$cchFile fileName:$fileName")
    }

    /**
     * sets the source range and sheet for the pivot table
     *
     * @param rc
     */
    fun setRange(rc: IntArray, sheetName: String?) {
        rwFirst = rc[0].toShort()
        colFirst = rc[1].toShort()
        rwLast = rc[2].toShort()
        colLast = rc[3].toShort()
        // update record
        val data = this.getData()
        var b = ByteTools.shortToLEBytes(rwFirst)
        data[0] = b[0]
        data[1] = b[1]
        b = ByteTools.shortToLEBytes(rwLast)
        data[2] = b[0]
        data[3] = b[1]
        data[4] = colFirst.toByte()
        data[5] = colLast.toByte()
        sourceSheet = sheetName
    }

    /**
     * sets the cell range for the pivot cache
     *
     * @param cr
     */
    fun setCellRange(cr: String) {
        val sheetname: String?
        if (cr.indexOf("!") != -1)
            sheetname = cr.split("!".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()[0]
        else
            sheetname = fileName
        val rc = ExcelTools.getRangeCoords(cr)
        setRange(rc, sheetname)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2639291289806138985L

        /**
         * create a new default DCONREF source data range record
         *
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val dr = DConRef()
                dr.opcode = XLSConstants.DCONREF
                dr.setData(byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0))
                dr.init()
                return dr
            }
    }
}
