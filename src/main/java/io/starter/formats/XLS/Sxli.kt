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

import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.Arrays


/**
 * SxLI: The SXLI record specifies pivot lines for the row area or column area
 * of a PivotTable view.
 *
 *
 * SXLI B5h: This record stores an array of variable-length SXLI structures
 * which describe row and col items.
 *
 *
 * Thre are 2 of these recs per Table -- one for rows, one for cols
 *
 *
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rgsxli      var     Array of SXLI structures
 *
 *
 * SXLI Structures have variable length but will always be at least 10 bytes long.
 *
 * offset  name        size    contents
 * ---
 * 0       cSic        2       count of identical items to previous
 * A signed integer that specifies the count of pivot item indexes in the beginning of the rgisxvi array that are identical to the same number of pivot item indexes
 * in the beginning of the rgisxvi array of the previous SXLIItem structure in the rgsxli array of the preceding SXLI record.
 * The value MUST be greater than or equal to 0 and less than the isxviMac field. If the fGrand field equals 1, then this value MUST be 0.
 * 2       itmtype     2       type:
 * 0x0 = data
 * 0x1 = default
 * 0x2 = SUM
 * 0x3 = COUNTA
 * 0x4 = COUNT
 * 0x5 = AVERAGE
 * 0x6 = MAX
 * 0x7 = MIN
 * 0x8 = PRODUCT
 * 0x9 = STDEV
 * 0xA = STDEVP
 * 0xB = VAR
 * 0xC = VARP
 * 0xD = GRAND TOTAL
 * 4       isxviMac    2	   Number of elements in the rgisxvi array that are displayed in this pivot line. MUST be greater than or equal to 0.
 * If the fGrand field equals 1, then the value of this field MUST be 1. If the fGrand field equals zero and the preceding SXLI record contains row area pivot items, then this value MUST be less than or equal to the cDimRw field of the preceding SxView. If the fGrand field equals zero and the preceding SXLI record contains column area pivot items, then this value MUST be less than or equal to the cDimCol field of the preceding SxView.
 * fMultiDataName (1 bit):  A bit that specifies whether the data field name is used for the total or the subtotal.
 * iData (8 bits): An unsigned integer that specifies a data item index as specified in Data Items, for an SXDI record specifying a data item used for a subtotal.
 * fSbt (1 bit): A bit that specifies whether this pivot line is a subtotal.
 * fBlock (1 bit): A bit that specifies whether this pivot line is a block total.
 * fGrand (1 bit): A bit that specifies whether this pivot line is a grand total.
 * fMultiDataOnAxis (1 bit): A bit that specifies whether a pivot line entry in this pivot line is a data item index.
 * 3 bits- unused
 * rgisxvi (variable): An array of 2-byte signed integers that specifies a pivot line entry.  --> if 0x7FFF means no pivot line/blank
 *
 * 6       grbit       2       option flags
 * 8       rgisxvi     2       Array of indices to SXVI records -- number is isxviMac+1
 *
</pre> *
 *
 *
 *
 * more info: rgsxli (variable): An array of SXLIItem.
 *
 *
 * Zero or two records of this type appear in the file for each PivotTable view
 * depending on the values of the cRw and cCol fields of the associated SxView
 * record.
 *
 *
 * If the value of either of the cRw or cCol fields of the associated SxView is
 * greater than zero, then two records of this type MUST exist in the file for
 * the associated SxView. The first record contains row area pivot lines and the
 * second record contains column area pivot lines.
 *
 *
 * The count of SXLIItem structures in rgsxli, which are row area pivot lines,
 * MUST equal the cRw field of SxView.
 *
 *
 * The count of SXLIItem structures in rgsxli, which are column area pivot
 * lines, MUST equal the cCol field of SxView.
 *
 *
 * The associated SxView record is the SxView record of the PivotTable view.
 *
 *
 * The SXLIItem structure specifies a pivot line in the row area or column area
 * of a PivotTable view.
 *
 *
 * cSic (2 bytes): A signed integer that specifies the count of pivot item
 * indexes in the beginning of the rgisxvi array that are identical to the same
 * number of pivot item indexes in the beginning of the rgisxvi array of the
 * previous SXLIItem structure in the rgsxli array of the preceding SXLI record.
 * The value MUST be greater than or equal to 0 and less than the isxviMac
 * field. If the fGrand field equals 1, then this value MUST be 0.
 *
 *
 * itmType (15 bits): An unsigned integer that specifies the type of this pivot
 * line. MUST be a value from the table:
 *
 *
 * ITMTYPEDATA 0x0000 A value in the data ITMTYPEDEFAULT 0x0001 Automatic
 * subtotal selection ITMTYPESUM 0x0002 Sum of values in the data ITMTYPECOUNTA
 * 0x0003 Count of values in the data ITMTYPECOUNT 0x0004 Count of numbers in
 * the data ITMTYPEAVERAGE 0x0005 Average of values in the data ITMTYPEMAX
 * 0x0006 Maximum value in the data ITMTYPEMIN 0x0007 Minimum value in the data
 * ITMTYPEPRODUCT 0x0008 Product of values in the data ITMTYPESTDEV 0x0009
 * Statistical standard deviation (estimate) ITMTYPESTDEVP 0x000A Statistical
 * standard deviation (entire population) ITMTYPEVAR 0x000B Statistical variance
 * (estimate) ITMTYPEVARP 0x000C Statistical variance (entire population)
 * ITMTYPEGRAND 0x000D Grand total ITMTYPEBLANK 0x000E Blank line
 *
 *
 * A - reserved1 (1 bit): MUST be zero and MUST be ignored.
 *
 *
 * isxviMac (2 bytes): A signed integer that specifies the number of elements in
 * the rgisxvi array that are displayed in this pivot line. MUST be greater than
 * or equal to 0. If the fGrand field equals 1, then the value of this field
 * MUST be 1. If the fGrand field equals zero and the preceding SXLI record
 * contains row area pivot items, then this value MUST be less than or equal to
 * the cDimRw field of the preceding SxView. If the fGrand field equals zero and
 * the preceding SXLI record contains column area pivot items, then this value
 * MUST be less than or equal to the cDimCol field of the preceding SxView.
 *
 *
 * B - fMultiDataName (1 bit): A bit that specifies whether the data field name
 * is used for the total or the subtotal. MUST be a value from the following
 * table: Value Meaning 0 The data field name is used for the total. 1 The data
 * field name is used for the subtotal.
 *
 *
 * If the fGrand field equals 1 or the fBlock field equals 1, then this value
 * MUST equal the value in the fMultiDataOnAxis field. If the fGrand and fBlock
 * fields equal zero, the fSbt and fMultiDataOnAxis fields equal 1, and the cSic
 * field is less than iposData, then this value MUST equal 1. Otherwise, this
 * value MUST be zero.
 *
 *
 * iposData is specified as follows: If the preceding SXLI record contains row
 * area pivot items, iposData equals the index of the SxIvdRw record in the
 * rgSxivd array of the SxIvd containing SxIvdRw records where the rw field
 * equals -2. If there is not an SxIvdRw record with the rw field equal to -2,
 * iposData equals zero.
 *
 *
 * If the preceding SXLI record contains column area pivot items, iposData
 * equals the index of the SxIvdCol record in the rgSxivd array of the SxIvd
 * containing SxIvdCol records where the col field equals -2. If there is not an
 * SxIvdCol record with the col field equal to -2, iposData equals zero.
 *
 *
 * iData (8 bits): An unsigned integer that specifies a data item index as
 * specified in Data Items, for an SXDI record specifying a data item used for a
 * subtotal. This field MUST be 0 if the cDimData field of the preceding SxView
 * record is 0 or if the fGrand field equals 1. If the cDimData field of the
 * preceding SxView is greater than 0, then this value MUST be greater than or
 * equal to 0 and less than the cDimData field of the preceding SxView record.
 * If the fMultiDataOnAxis field equals 1 and the itmType field does not equal
 * ITMTYPEBLANK and the isxviMac field is greater than iposData as specified in
 * fMultiDataName, then the value of this field MUST equal the value of the
 * element of the rgisxvi array in the position equal to iposData as specified
 * in fMultiDataName.
 *
 *
 * C - fSbt (1 bit): A bit that specifies whether this pivot line is a subtotal.
 * This value MUST equal 1 if the itmType field is greater than or equal to
 * ITMTYPEDEFAULT and the itmType field is less than or equal to ITMTYPEGRAND
 * and the fBlock field equals 0. Otherwise, this value MUST be 0.
 *
 *
 * D - fBlock (1 bit): A bit that specifies whether this pivot line is a block
 * total. A block total is a total of a group of pivot items. For more details
 * see Grouping. If the fGrand field equals 0 and the fBlock field in the
 * previous SXLIItem record equals 1, this value MUST be 1.
 *
 *
 * E - fGrand (1 bit): A bit that specifies whether this pivot line is a grand
 * total. If the fGrand field in the previous SXLIItem record is 1, then this
 * value MUST be 1. Otherwise, if the itmType field equals ITMTYPEGRAND this
 * field MUST equal 1 and if the itmType field does not equal ITMTYPEGRAND this
 * field MUST equal 0.
 *
 *
 * F - fMultiDataOnAxis (1 bit): A bit that specifies whether a pivot line entry
 * in this pivot line is a data item index.
 *
 *
 * If the preceding SXLI record contains row area pivot items, the cDimData
 * field of the preceding SxView record is greater than 1, the
 * sxaxis4Data.sxaxisRw field of the preceding SxView equals 1 and itmType is
 * not equal to ITMTYPEBLANK, then this value MUST be 1. Otherwise, this value
 * MUST be 0.
 *
 *
 * If the preceding SXLI record contains column area pivot items, the cDimData
 * field of the preceding SxView record is greater than 1, the
 * sxaxis4Data.sxaxisCol field of the preceding SxView equals 1 and itmType is
 * not equal to ITMTYPEBLANK, then this value MUST be 1. Otherwise, this value
 * MUST be 0.
 *
 *
 * G - unused1 (1 bit): Undefined, and MUST be ignored.
 *
 *
 * H - unused2 (1 bit): Undefined, and MUST be ignored.
 *
 *
 * I - reserved2 (1 bit): MUST be zero and MUST be ignored.
 *
 *
 * rgisxvi (variable): An array of 2-byte signed integers that specifies a pivot
 * line entry. Each element of this array is either a pivot item index or a data
 * item index. If fGrand is 1 or itmType is ITMTYPEBLANK then all elements of
 * this field are undefined and MUST be ignored. Otherwise each element MUST be
 * a value from the following table: Value Meaning 0x0000 to 0x7EF4 This value
 * specifies a data item index or pivot item index in the associated pivot field
 * as specified in Pivot Items. 0x7FFF This value specifies that there is no
 * pivot item and that the cell in the pivot line is blank.
 *
 *
 *
 *
 *
 *
 * "http://www.extentech.com">Extentech Inc.
 */
class Sxli : XLSRecord() {
    internal var items: ArrayList<SXLI_Item>
    internal var nItemsPerLine: Int = 0

    override fun init() {
        super.init()
        val sx = wkbook!!.allPivotTableViews[wkbook!!.nPivotTableViews - 1]
        nItemsPerLine = (if (sx.hasRowPivotItemsRecord()) sx.cDimCol else sx.cDimRw).toInt()
        // total # items will be sx.cRw or sx.cCol
        // per each pivot line, # items [see SXLI_Item.rgisxvi] will be sx.cDimRw or sx.cDimCol
        items = SXLI_Item.parse(this.getData()!!, nItemsPerLine)
        if (DEBUGLEVEL > 3)
            Logger.logInfo("SXLI: rgsxli:$items")
    }

    override fun toString(): String {
        return "SXLI: rgsxli:$items"
    }

    /**
     * adds a field to the end of this field list
     */
    fun addField(repeat: Int, nLines: Int, type: Int, indexes: ShortArray) {
        val sxitem = SXLI_Item(repeat, nLines, type, indexes, nItemsPerLine)
        items.add(sxitem)
        updateRecord(true)
    }

    internal fun updateRecord(appendLast: Boolean) {
        if (!appendLast) { // add all
            this.data = ByteArray(0)
            for (sxitem in items) {
                data = ByteTools.append(sxitem.data, data)
            }
        } else
            data = ByteTools.append(items[items.size - 1].data, data)

    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4157827774990504633L

        /**
         * return a new, blank SxLi
         *
         * @return
         */
        fun getPrototype(wkbook: WorkBook): XLSRecord {
            val li = Sxli()
            li.workBook = wkbook
            li.opcode = XLSConstants.SXLI
            li.setData(byteArrayOf())
            li.init()
            return li
        }
    }
}

/**
 * The SXLIItem structure specifies a pivot line in the row area or column area
 * of a PivotTable view.
 */
internal class SXLI_Item {

    var cSic: Short = 0
    /* 0x0000 specifies that no pivot items in the rgisxvi array are identical to the first pivot items in the previous pivot line item in this record.
     */
    var itmType: Short = 0
    /* see ITEMTYPE enum */
    var isxviMac: Short = 0 /* number of pivot items on the pivot line */
    var iData: Byte = 0 /* specifies a data item index as specified in Data Items, for an SXDI record specifying a data item used for a subtotal.  This field MUST be 0 if the cDimData field of the preceding SxView record is 0 or if the fGrand field equals */
    var fMultiDataName: Boolean = false
    /* specifies whether the data field name is used for the total or the subtotal */
    var fSbt: Boolean = false
    /* Specifies whether this pivot line is a subtotal. */
    var fBlock: Boolean = false
    /* Specifies whether this pivot line is a block total */
    var fGrand: Boolean = false
    /* Specifies whether this pivot line is a grand total. */
    var fMultiDataOnAxis: Boolean = false /* Specifies whether a pivot line entry in this pivot line is a data item index. */
    var rgisxvi: ByteArray /* An array of 2-byte signed integers that specifies a pivot line entry.
						Each element of this array is either a pivot item index or a data item index.
						0x7FFF means blank, no pivot item
					 */

    /**
     * package SXLI_ITEM into byte array
     *
     * @return byte[]
     */
    val data: ByteArray
        get() {
            var data = ByteArray(8)
            var b = ByteTools.shortToLEBytes(cSic)
            data[0] = b[0]
            data[1] = b[1]
            b = ByteTools.shortToLEBytes(itmType)
            data[2] = b[0]
            data[3] = b[1]
            b = ByteTools.shortToLEBytes(isxviMac)
            data[4] = b[0]
            data[5] = b[1]
            var tmp = (if (fMultiDataName) 1 else 0).toShort()
            tmp = (tmp or (iData shl 1)).toShort()
            if (fSbt)
                tmp = (tmp or 0x200).toShort()
            if (fBlock)
                tmp = (tmp or 0x400).toShort()
            if (fGrand)
                tmp = (tmp or 0x800).toShort()
            if (fMultiDataOnAxis)
                tmp = (tmp or 0x1000).toShort()
            b = ByteTools.shortToLEBytes(tmp)
            data[6] = b[0]
            data[7] = b[1]
            data = ByteTools.append(rgisxvi, data)
            return data
        }

    enum class ITEMTYPE private constructor(/* Blank line */
            private val itm: String) {
        ITMTYPEDATA("data"), /* A value in the data */
        ITMTYPEDEFAULT("default"), /* Automatic subtotal selection */
        ITMTYPESUM("sum"), /* Sum of values in the data */
        ITMTYPECOUNTA("countA"), /* Count of values in the data */
        ITMTYPECOUNT("count"), /* Count of numbers in the data */
        ITMTYPEAVERAGE("avg"), /* Average of values in the data */
        ITMTYPEMAX("max"), /* Maximum value in the data */
        ITMTYPEMIN("min"), /* Minimum value in the data */
        ITMTYPEPRODUCT("product"), /* Product of values in the data */
        ITMTYPESTDEV("stdDev"), /* Statistical standard deviation (estimate) */
        ITMTYPESTDEVP("stdDevP"), /*
         * Statistical standard deviation (entire
         * population)
         */
        ITMTYPEVAR("var"), /* Statistical variance (estimate) */
        ITMTYPEVARP("varP"), /* Statistical variance (entire population) */
        ITMTYPEGRAND("grand"), /* Grand total */
        ITMTYPEBLANK("blank");


        companion object {

            operator fun get(s: String): Int {
                for (c in values()) {
                    if (c.itm == s)
                        return c.ordinal
                }
                return 0
            }

            operator fun get(id: Int): ITEMTYPE? {
                for (c in values()) {
                    if (c.ordinal == id)
                        return c
                }
                return null
            }
        }

    }

    constructor(data: ByteArray, nItemsPerLine: Int) {
        try {
            cSic = ByteTools.readShort(data[pos + 0].toInt(), data[pos + 1].toInt())
            itmType = (ByteTools.readShort(data[pos + 2].toInt(), data[pos + 3].toInt()) and 0x7FFF).toShort() // 1st // 15 bits
            isxviMac = ByteTools.readShort(data[pos + 4].toInt(), data[pos + 5].toInt())
            var tmp = ByteTools.readShort(data[pos + 6].toInt(), data[pos + 7].toInt())
            fMultiDataName = tmp and 0x8000 == 0x8000
            iData = (tmp shr 7 and 0x80).toByte()
            tmp = (tmp shr 9).toByte().toShort()
            fSbt = tmp and 0x1 == 0x1
            fBlock = tmp and 0x2 == 0x2
            fGrand = tmp and 0x4 == 0x4
            fMultiDataOnAxis = tmp and 0x8 == 0x8
            pos += 8
            // not in 1.5 rgisxvi = Arrays.copyOfRange(data, pos, pos + (nItemsPerLine * 2));
            rgisxvi = ByteArray(nItemsPerLine * 2)
            System.arraycopy(data, pos, rgisxvi, 0, rgisxvi.size)
            pos += rgisxvi.size
        } catch (e: Exception) {
        }

    }

    /**
     * create a new SXLI item from speicifications for particular row or column line(s)
     *
     * @param repeat        if a field is repeated NOT IMPLEMENTED YET
     * @param nLines        # lines for
     * @param type
     * @param indexes
     * @param nItemsPerLine
     */
    constructor(repeat: Int, nLines: Int, type: Int, indexes: ShortArray, nItemsPerLine: Int) {
        cSic = repeat.toShort()
        isxviMac = nLines.toShort()
        val t = ITEMTYPE[type]
        itmType = t!!.ordinal.toShort()
        // fMultiDataName ???
        // iData ???
        // fBlock ???
        // fMultiDataOnAxis ???
        when (t) {
            SXLI_Item.ITEMTYPE.ITMTYPEDATA, SXLI_Item.ITEMTYPE.ITMTYPEBLANK -> {
            }
            SXLI_Item.ITEMTYPE.ITMTYPEGRAND -> {
                fGrand = true
                fSbt = true
            }
            else -> fSbt = true
        }/*			case ITMTYPEDEFAULT:
			case ITMTYPEAVERAGE:
			case ITMTYPECOUNT:
			case ITMTYPECOUNTA:
			case ITMTYPEMAX:
			case ITMTYPEMIN:
			case ITMTYPEPRODUCT:
			case ITMTYPESTDEV:
			case ITMTYPESTDEVP:
			case ITMTYPESUM:
			case ITMTYPEVAR:
			case ITMTYPEVARP:*/
        rgisxvi = ByteArray(indexes.size * 2)
        var i = 0
        while (i < indexes.size * 2) {
            val b = ByteTools.shortToLEBytes(indexes[i / 2])
            rgisxvi[i] = b[0]
            rgisxvi[i + 1] = b[1]
            i += 2
        }
    }

    override fun toString(): String {
        return String.format("[rep %d typ %d imax %d iData %d fSbt %b fBlock %b fGrand %b rgi: %s",
                cSic, itmType, isxviMac, iData, fSbt, fBlock, fGrand, Arrays.toString(rgisxvi))
    }

    companion object {

        var pos = 0

        fun parse(data: ByteArray, nItemsPerLine: Int): ArrayList<SXLI_Item> {
            pos = 0
            val items = ArrayList<SXLI_Item>()

            while (pos < data.size - 7) {
                items.add(SXLI_Item(data, nItemsPerLine))
            }
            return items
        }
    }
}