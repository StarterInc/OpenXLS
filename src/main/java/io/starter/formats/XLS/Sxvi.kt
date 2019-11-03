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

import java.io.UnsupportedEncodingException


/**
 * SXVI B2h: This record stores view information about an Item.
 * itmType (2 bytes): A signed integer that specifies the pivot item type.
 * The value MUST be one of the following values:
 * Value		Name	    	Meaning
 * 0x0000		itmtypeData	    A data value
 * 0x0001	    itmtypeDEFAULT	Default subtotal for the pivot field
 * 0x0002	    itmtypeSUM	    Sum of values in the pivot field
 * 0x0003		itmtypeCOUNTA	Count of values in the pivot field
 * 0x0004	    itmtypeAVERAGE	Average of values in the pivot field
 * 0x0005	    itmtypeMAX	    Max of values in the pivot field
 * 0x0006	    itmtypeMIN	    Min of values in the pivot field
 * 0x0007	    itmtypePRODUCT  Product of values in the pivot field
 * 0x0008	    itmtypeCOUNT	Count of numbers in the pivot field
 * 0x0009	    itmtypeSTDEV	Statistical standard deviation (estimate) of the pivot field
 * 0x000A	    itmtypeSTDEVP	Statistical standard deviation (entire population) of the pivot field
 * 0x000B		itmtypeVAR	    Statistical variance (estimate) of the pivot field
 * 0x000C	    itmtypeVARP	    Statistical variance (entire population) of the pivot field
 *
 *
 * A - fHidden (1 bit): A bit that specifies whether this pivot item is hidden.
 * MUST be zero if itmType is not itmtypeData. MUST be zero for OLAP PivotTable view.
 *
 *
 * B - fHideDetail (1 bit): A bit that specifies whether the pivot item detail is collapsed.
 * MUST be zero for OLAP PivotTable view.
 *
 *
 * C - reserved1 (1 bit): MUST be zero, and MUST be ignored.
 *
 *
 * D - fFormula (1 bit): A bit that specifies whether this pivot item is a calculated item.
 * This field MUST be zero if any of the following apply:
 * itmType is not zero.
 * This item is in an OLAP PivotTable view.
 * The sxaxisPage field of sxaxis in the Sxvd record of the pivot field equals 1 (the associated Sxvd is the last Sxvd record before this record in the stream (1)).
 * The fCalculatedField field in the SXVDEx record of the pivot field equals 1.
 * There is not an associated SXFDB record in the associated PivotCache.
 * The fRangeGroup field of the SXFDB record, of the associated cache field of the pivot field, equals 1.
 * The fCalculatedField field of the SXFDB record, of the associated cache field of the pivot field, equals 1.
 *
 *
 * E - fMissing (1 bit): A bit that specifies if this pivot item does not exist in the data source (1).
 * MUST be zero if itmType is not zero. MUST be zero for OLAP PivotTable view.
 *
 *
 * reserved2 (11 bits): MUST be zero, and MUST be ignored.
 *
 *
 * iCache (2 bytes): A signed integer that specifies a reference to a cache item.
 * MUST be a value from the following table:
 * Value			Meaning
 * -1			    No cache item is referenced.
 * 0+			    A cache item index in the cache field associated with the pivot field, as specified by Cache Items.
 * If itmType is not zero, a reference to a cache item is not specified and this value MUST be -1.
 * Otherwise, this value MUST be greater than or equal to 0.
 *
 *
 * cchName (2 bytes): An unsigned integer that specifies the length of the stName string.
 * If the value is 0xFFFF then stName is NULL.
 * Otherwise, the value MUST be less than or equal to 254.
 *
 *
 * stName (variable): An XLUnicodeStringNoCch structure that specifies the name of this pivot item.
 * If not NULL, this is used as the caption of the pivot item instead of the value in the
 * cache item specified by iCache. The length of this field is specified in cchName.
 * This field exists only if cchName is not 0xFFFF. If this is in a non-OLAP PivotTable
 * view and this string is not NULL, it MUST be unique within all SXVI records in associated
 * with the pivot field.
 */

class Sxvi : XLSRecord() {
    internal var data: ByteArray? = null
    /**
     * returns the type of this pivot item
     * <br></br>one of:
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
     *
     * @return
     */
    var itemType: Short = -1
        internal set
    internal var cchName: Short = -1
    internal var name: String? = null
    internal var iCache: Short = -1
    internal var fHidden: Boolean = false
    internal var fHideDetail: Boolean = false
    internal var fFormula: Boolean = false
    internal var fMissing: Boolean = false

    /**
     * returns true if this pivot item is hidden
     *
     * @return
     */
    /**
     * sets the hidden state for this pivot item
     *
     * @param b
     */
    var isHidden: Boolean
        get() = fHidden
        set(b) {
            fHidden = b
            val by = this.getByteAt(2)
            if (fHidden)
                this.getData()[2] = (by and 0x1).toByte()
            else
                this.getData()[2] = (by xor 0x1).toByte()
        }

    /**
     * specifies whether the pivot item detail is collapsed.
     *
     * @return
     */
    /**
     * specifies whether the pivot item detail is collapsed.
     *
     * @return
     */
    var isCollapsed: Boolean
        get() = fHideDetail
        set(b) {
            fHideDetail = b
            val by = this.getByteAt(2)
            if (fHideDetail)
                this.getData()[2] = (by and 0x2).toByte()
            else
                this.getData()[2] = (by xor 0x2).toByte()
        }
    // TODO: fFormula, fMissing


    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, /* itmtype */
            0, 0, /* flags */
            0, 0, /* icache */
            -1, -1)/* cchName */


    override fun init() {
        super.init()
        itemType = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        val b = this.getByteAt(2)
        fHidden = b and 0x1 == 0x1
        fHideDetail = b and 0x2 == 0x2
        // bit 3- reserved
        fFormula = b and 0x8 == 0x8
        fMissing = b and 0x10 == 0x10
        iCache = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        cchName = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        if (cchName.toInt() != -1) {
            val encoding = this.getByteAt(10)
            val tmp = this.getBytesAt(11, cchName * (encoding + 1))
            try {
                if (encoding.toInt() == 0)
                    name = String(tmp!!, XLSConstants.DEFAULTENCODING)
                else
                    name = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding PivotTable caption name in Sxvd: $e")
            }

        }
        if (DEBUGLEVEL > 3) Logger.logInfo("SXVI - itemtype:$itemType iCache: $iCache name:$name")
    }

    override fun toString(): String {
        return "SXVI - itemtype:$itemType iCache: $iCache name:$name"
    }

    /**
     * sets the pivot item type:
     * <br></br>one of:
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
     *
     * @return
     */
    fun setItemType(type: Int) {
        itemType = type.toShort()
        val b = ByteTools.shortToLEBytes(itemType)
        this.getData()[0] = b[0]
        this.getData()[1] = b[1]
    }

    /**
     * reference to a cache item :
     * <br></br>-1			    No cache item is referenced.
     * <br></br>0+			    A cache item index in the cache field associated with the pivot field, as specified by Cache Items.
     *
     * @param icache
     */
    fun setCacheItem(icache: Int) {
        this.iCache = icache.toShort()
        val b = ByteTools.shortToLEBytes(this.iCache)
        this.getData()[4] = b[0]
        this.getData()[5] = b[1]
    }

    /**
     * returns the name of this pivot item; if not null, is the caption
     * for this pivot item
     *
     * @return
     */
    fun getName(): String? {
        return name
    }

    /**
     * returns the name of this pivot item; if not null, is the caption
     * for this pivot item
     */
    fun setName(name: String?) {
        this.name = name
        var data = ByteArray(8)
        System.arraycopy(this.getData()!!, 0, data, 0, 7)
        if (name != null) {
            var strbytes: ByteArray? = null
            try {
                strbytes = this.name!!.toByteArray(charset(XLSConstants.DEFAULTENCODING))
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding pivot table name in SXVI: $e")
            }

            //update the lengths:
            cchName = strbytes!!.size.toShort()
            val nm = ByteTools.shortToLEBytes(cchName)
            data[6] = nm[0]
            data[7] = nm[1]

            // now append variable-length string data
            val newrgch = ByteArray(cchName + 1)    // account for encoding bytes
            System.arraycopy(strbytes, 0, newrgch, 1, cchName.toInt())

            data = ByteTools.append(newrgch, data)
        } else {
            data[6] = -1
            data[7] = -1
        }
        this.setData(data)
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6399665481118265257L

        val itmtypeData: Short = 0x0000        //A data value
        val itmtypeDEFAULT: Short = 0x0001    //Default subtotal for the pivot field
        val itmtypeSUM: Short = 0x0002        //Sum of values in the pivot field
        val itmtypeCOUNTA: Short = 0x0003    //Count of values in the pivot field
        val itmtypeAVERAGE: Short = 0x0004    // Average of values in the pivot field
        val itmtypeMAX: Short = 0x0005        // Max of values in the pivot field
        val itmtypeMIN: Short = 0x0006        //Min of values in the pivot field
        val itmtypePRODUCT: Short = 0x0007    //Product of values in the pivot field
        val itmtypeCOUNT: Short = 0x0008    //Count of numbers in the pivot field
        val itmtypeSTDEV: Short = 0x0009    //Statistical standard deviation (estimate) of the pivot field
        val itmtypeSTDEVP: Short = 0x000A    //Statistical standard deviation (entire population) of the pivot field
        val itmtypeVAR: Short = 0x000B        //Statistical variance (estimate) of the pivot field
        val itmtypeVARP: Short = 0x000C        //Statistical variance (entire population) of the pivot field

        val prototype: XLSRecord?
            get() {
                val si = Sxvi()
                si.opcode = XLSConstants.SXVI
                si.setData(si.PROTOTYPE_BYTES)
                si.init()
                return si
            }
    }
}