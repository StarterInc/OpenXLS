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

import io.starter.OpenXLS.DateConverter
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException
import java.util.Arrays
import java.util.Locale

/**
 * The SXAddl record specifies additional information for a PivotTable view,
 * PivotCache, or query table. The current class and full type of this record
 * are specified by the hdr field which determines the contents of the data
 * field.
 *
 *
 * hdr (6 bytes): An SXAddlHdr structure that specifies header information for
 * an SXAddl record. data (variable): A variable-size field that contains data
 * specific to the full record type of the SXAddl record.
 *
 *
 *
 *
 * The SXAddlHdr structure specifies header information for an SXAddl record.
 * frtHeaderOld (4 bytes): An FrtHeaderOld. The frtHeaderOld.rt field MUST be
 * 0x0864. sxc (1 byte): An unsigned integer that specifies the current class.
 * See class for details. sxd (1 byte): An unsigned integer that specifies the
 * type of record contained in the data field of the containing SXAddl record.
 * See class for details
 *
 *
 *
 *
 * "http://www.extentech.com">Extentech Inc.
 */
class SxAddl : XLSRecord(), XLSConstants {
    private var sxc: Short = 0
    private var sxd: Short = 0

    /**
     * returns the class of this SXADDL_ record
     *
     * @return ADDL_CLASSES instance
     */
    val adDlClass: ADDL_CLASSES?
        get() = ADDL_CLASSES[sxc.toInt()]

    /**
     * returns the class which matches the class/record id of this SXADDL_
     * record
     *
     * @return
     */
    /* 0 */ val recordId: Any?
        get() {
            when (ADDL_CLASSES[sxc.toInt()]) {
                SxAddl.ADDL_CLASSES.sxcView -> return SxcView.lookup(sxd.toInt())
                SxAddl.ADDL_CLASSES.sxcCache -> return SxcCache.lookup(sxd.toInt())
                SxAddl.ADDL_CLASSES.sxcField12 -> return SxcField12.lookup(sxd.toInt())
            }
            return null
        }

    internal enum class ADDL_CLASSES private constructor(cls: Int) {
        sxcView(0), sxcField(1), sxcHierarchy(2), sxcCache(3), /* 3 */
        sxcCacheField(4), sxcQsi(5), sxcQuery(6), sxcGrpLevel(7), sxcGroup(8), sxcCacheItem(
                9), /* 9 */
        sxcSxrule(0xC), sxcSxfilt(0xD), sxcSxdh(0x10), sxcAutoSort(0x12), sxcSxmgs(
                0x13),
        sxcSxmg(0x14), sxcField12(0x17), sxcSxcondfmts(0x1A), sxcSxcondfmt(
                0x1B),
        sxcSxfilters12(0x1C), sxcSxfilter12(0x1D);

        private val cls: Short

        init {
            this.cls = cls.toShort()
        }

        fun sxd(): Short {
            return cls
        }

        companion object {

            operator fun get(cls: Int): ADDL_CLASSES? {
                for (c in values()) {
                    if (c.cls.toInt() == cls)
                        return c
                }
                return null
            }
        }
    }

    /**
     * sxc= 3
     */
    internal enum class SxcCache private constructor(sxd: Int) {
        SxdId(0), // SXAddl_SXCCache_SXDId
        SxdVerUpdInv(1), // SXAddl_SXCCache_SXDVerUpdInv
        SxdVer10Info(2), // SXAddl_SXCCache_SXDVer10Info
        SxdVerSxMacro(0x18), // SXAddl_SXCCache_SXDVerSXMacro 0x18 (24)
        SxdInvRefreshReal(0x34), // SXAddl_SXCCache_SXDInvRefreshReal 0x34
        SxdInfo12(0x41), // SXAddl_SXCCache_SXDInfo12 0x41
        SxdEnd(-1);

        // SXAddl_SXCCache_SXDEnd 0xFF
        private val sxd: Short

        init {
            this.sxd = sxd.toShort()
        }

        fun sxd(): Short {
            return sxd
        }

        companion object {

            fun lookup(record: Int): SxcCache? {
                for (c in values()) {
                    if (c.sxd.toInt() == record)
                        return c
                }
                return null
            }
        }
    }

    /**
     * sxc= 0
     */
    internal enum class SxcView private constructor(sxd: Int) {
        sxdId(0), // SXAddl_SXCView_SXDId
        sxdVerUpdInv(1), // SXAddl_SXCView_SXDVerUpdInv
        sxdVer10Info(2), // SXAddl_SXCView_SXDVer10Info
        sxdCalcMember(3), // SXAddl_SXCView_SXDCalcMember
        sxdCalcMemString(0xA), // SXAddl_SXCView_SXDCalcMemString 0xA
        sxdVer12Info(0x19), // SXAddl_SXCView_SXDVer12Info 0x19
        sxdTableStyleClient(0x1E), // SXAddl_SXCView_SXDTableStyleClient 0x1E
        // (30)
        sxdCompactRwHdr(0x21), // SXAddl_SXCView_SXDCompactRwHdr 0x21
        sxdCompactColHdr(0x22), // SXAddl_SXCView_SXDCompactColHdr 0x22
        sxdSxpiIvmb(0x26), // SXAddl_SXCView_SXDSXPIIvmb 0x26
        sxdEnd(-1);

        // SXAddl_SXCView_SXDEnd 0xFF
        private val sxd: Short

        init {
            this.sxd = sxd.toShort()
        }

        fun sxd(): Short {
            return sxd
        }

        companion object {

            fun lookup(record: Int): SxcView? {
                for (c in values()) {
                    if (c.sxd.toInt() == record)
                        return c
                }
                return null
            }
        }
    }

    /**
     * sxd= 0x17
     */
    internal enum class SxcField12 private constructor(sxd: Int) {
        sxdId(0), sxdVerUpdInv(1), sxdMemberCaption(0x11), sxdVer12Info(0x19), sxdIsxth(
                0x1C),
        sxdAutoshow(0x37), sxdEnd(-1);

        private val sxd: Short

        init {
            this.sxd = sxd.toShort()
        }

        fun sxd(): Short {
            return sxd
        }

        companion object {

            fun lookup(record: Int): SxcField12? {
                for (c in values()) {
                    if (c.sxd.toInt() == record)
                        return c
                }
                return null
            }
        }
    }

    override fun init() {
        super.init()
        sxc = this.getData()!![4].toShort() // class: see addlclass
        sxd = this.getData()!![5].toShort()
        val len = this.getData()!!.size
        /*
         * notes: If the value of the hdr.sxc field of SXAddl is 0x09 and the
         * value of the hdr.sxd field of SXAddl is 0xFF, then the current class
         * is specified by SxcCacheField class and the full record type is
         * SXAddl_SXCCacheItem_SXDEnd. Classes can be nested inside other
         * classes in a hierarchical manner
         */
        when (ADDL_CLASSES[sxc.toInt()]) {
            SxAddl.ADDL_CLASSES.sxcView /* 0 */ -> {
                val record = SxcView.lookup(sxd.toInt())
                when (record) {
                    SxAddl.SxcView.sxdId ->
                        // An SXAddl_SXString structure that specifies the PivotTable
                        // view that this SxcView class applies to.
                        // The corresponding SxView record of this PivotTable view is
                        // the SxView record, in this Worksheet substream,
                        // with its stTable field equal to the value of this field.
                        // SXADDL_sxcView: record=sxdId data:[11, 0, 0, 0, 0, 0, 11, 0,
                        // 0, 80, 105, 118, 111, 116, 84, 97, 98, 108, 101, 50]
                        // cchTotal 4 bytes -- if multiple segments (for strings > 255)
                        // will be 0
                        // reserved 2 bytes
                        // String-- cch-2 bytes, encoding-1 byte
                        if (len > 6) {
                            var cch = ByteTools.readShort(this.getData()!![6].toInt(),
                                    this.getData()!![7].toInt())
                            if (cch > 0) { // otherwise it's a multiple segment
                                cch = ByteTools.readShort(this.getData()!![12].toInt(),
                                        this.getData()!![13].toInt())
                                val encoding = this.getData()!![14].toShort()
                                val tmp = this
                                        .getBytesAt(15, cch * (encoding + 1))
                                var name: String? = null
                                try {
                                    if (encoding.toInt() == 0)
                                        name = String(tmp!!, XLSConstants.DEFAULTENCODING)
                                    else
                                        name = String(tmp!!, XLSConstants.UNICODEENCODING)
                                } catch (e: UnsupportedEncodingException) {
                                    Logger.logInfo("encoding PivotTable caption name in Sxvd: $e")
                                }

                                if (DEBUGLEVEL > 3)
                                    Logger.logInfo("SXADDL_sxcView: record=" + record
                                            + " name: " + name)
                            } else if (DEBUGLEVEL > 3)
                                Logger.logInfo("SXADDL_sxcView: record=" + record
                                        + " name: MULTIPLESEGMENTS")
                        } else if (DEBUGLEVEL > 3)
                            Logger.logInfo("SXADDL_sxcView: record=" + record
                                    + " name: null")
                    SxAddl.SxcView.sxdVer10Info, SxAddl.SxcView.sxdTableStyleClient, SxAddl.SxcView.sxdVerUpdInv -> if (DEBUGLEVEL > 3)
                        Logger.logInfo("SXADDL_sxcView: record=" + record
                                + " data:"
                                + Arrays.toString(this.getBytesAt(6, len - 6)))
                }
            }
            SxAddl.ADDL_CLASSES.sxcCache /* 3 */ -> {
                val crec = SxcCache.lookup(sxd.toInt())
                when (crec) {
                    SxAddl.SxcCache.SxdVer10Info -> {
                        val verLastRefresh = this.getByteAt(16)
                        val verRefreshMin = this.getByteAt(17)
                        val lastdate = ByteTools.eightBytetoLEDouble(this.getBytesAt(18, 8)!!)
                        val ld = DateConverter.getDateFromNumber(lastdate)
                        if (DEBUGLEVEL > 3) {
                            val dateFormatter = java.text.DateFormat.getDateInstance(java.text.DateFormat.DEFAULT, Locale.getDefault())
                            Logger.logInfo("SXADDL_sxcCache: record=" + crec +
                                    " lastDate:" + dateFormatter.format(ld) + " verLast:" + verLastRefresh
                                    + " verMin:" + verRefreshMin)
                        }
                    }
                    else -> if (DEBUGLEVEL > 3)
                        Logger.logInfo("SXADDL_sxcCache: record=" + crec + " data:"
                                + Arrays.toString(this.getBytesAt(6, len - 6)))
                }
            }
            SxAddl.ADDL_CLASSES.sxcField12 -> {
                val srec = SxcField12.lookup(sxd.toInt())
                if (DEBUGLEVEL > 3)
                    Logger.logInfo("SXADDL_sxcField12: record=" + srec + " data:"
                            + Arrays.toString(this.getBytesAt(6, len - 6)))
            }
            SxAddl.ADDL_CLASSES.sxcField, SxAddl.ADDL_CLASSES.sxcHierarchy, SxAddl.ADDL_CLASSES.sxcCacheField, SxAddl.ADDL_CLASSES.sxcQsi, SxAddl.ADDL_CLASSES.sxcQuery, SxAddl.ADDL_CLASSES.sxcGrpLevel, SxAddl.ADDL_CLASSES.sxcGroup, SxAddl.ADDL_CLASSES.sxcCacheItem, SxAddl.ADDL_CLASSES.sxcSxrule, SxAddl.ADDL_CLASSES.sxcSxfilt, SxAddl.ADDL_CLASSES.sxcSxdh, SxAddl.ADDL_CLASSES.sxcAutoSort, SxAddl.ADDL_CLASSES.sxcSxmgs, SxAddl.ADDL_CLASSES.sxcSxmg, SxAddl.ADDL_CLASSES.sxcSxcondfmts, SxAddl.ADDL_CLASSES.sxcSxcondfmt, SxAddl.ADDL_CLASSES.sxcSxfilters12, SxAddl.ADDL_CLASSES.sxcSxfilter12 -> if (DEBUGLEVEL > 3)
                Logger.logInfo("SXADDL: hdr: " + " sxc:" + sxc + " sxd:" + sxd
                        + " data:"
                        + Arrays.toString(this.getBytesAt(6, len - 6)))
        }
    }

    /**
     * for SXADDL_SxView_SxDID record this sets the view name (matches table
     * name in Sxview)
     *
     * @param viewName
     */
    fun setViewName(viewName: String) {
        if (sxc.toInt() != 0 && sxd.toInt() != 0)
            Logger.logErr("Incorrect SXADDL_ record for view name")

        var data = ByteArray(14)
        System.arraycopy(this.getData()!!, 0, data, 0, 5)
        var strbytes: ByteArray? = null
        try {
            strbytes = viewName.toByteArray(charset(XLSConstants.DEFAULTENCODING))
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("encoding pivot view name in SXADDL: $e")
        }

        // update the lengths:
        val cch = strbytes!!.size.toShort()
        val nm = ByteTools.shortToLEBytes(cch)
        data[6] = nm[0]
        data[7] = nm[1]
        data[12] = nm[0]
        data[13] = nm[1]

        // now append variable-length string data
        val newrgch = ByteArray(cch + 1) // account for encoding bytes
        System.arraycopy(strbytes, 0, newrgch, 1, cch.toInt())

        data = ByteTools.append(newrgch, data)
        this.setData(data)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2639291289806138985L

        /**
         * creates a SxAddl record for the desired class and record id
         *
         * @param cls      int class one of ADDL_CLASSES enum
         * @param recordid desired record in class
         * @param dara     if not null,specifies the data for the class. if null, the
         * default data will be used
         * @return SxAddl record
         */
        fun getDefaultAddlRecord(cls: SxAddl.ADDL_CLASSES,
                                 recordid: Int, data: ByteArray?): SxAddl {
            var data = data
            val sxa = SxAddl()
            sxa.opcode = XLSConstants.SXADDL
            var newData = ByteArray(6)
            newData[0] = 100
            newData[1] = 8
            newData[4] = cls.ordinal.toByte()
            newData[5] = recordid.toByte()

            if (data == null) { // if !null, use passed in data for record creation
                // and return; otherwise create default data for
                // record
                when (cls) {
                    SxAddl.ADDL_CLASSES.sxcView /* 0 */ -> {
                        val record = SxcView.lookup(recordid)
                        when (record) {
                            SxAddl.SxcView.sxdId, SxAddl.SxcView.sxdTableStyleClient, SxAddl.SxcView.sxdVerUpdInv -> {
                            }
                            SxAddl.SxcView.sxdVer10Info -> data = byteArrayOf(1, 0x41, 0, 0, 0, 0) // common flags
                            SxAddl.SxcView.sxdEnd -> data = byteArrayOf(0, 0, 0, 0, 0, 0)
                        }
                    }
                    SxAddl.ADDL_CLASSES.sxcCache /* 3 */ -> {
                        val crec = SxcCache.lookup(recordid)
                        when (crec) {
                            SxAddl.SxcCache.SxdId // pivot cache stream id
                            -> data = byteArrayOf(1, 0, 0, 0, 0, 0)
                            SxAddl.SxcCache.SxdVer10Info -> {
                                data = byteArrayOf(0, 0, 0, 0, 0, 0, -1, -1, -1, -1, /* reserved, citmGhostMax */
                                        3, /* ver last saved -- 0 or 3 */
                                        0)    /* ver min */
                                //date last refreshed: 8 bytes
                                //reserved: 0, 0 };/* reserved 2 bytes*/
                                val d = DateConverter.getXLSDateVal(java.util.Date())
                                val dates = ByteTools.doubleToLEByteArray(d)
                                data = ByteTools.append(dates, data)
                                data = ByteTools.append(byteArrayOf(0, 0), data)
                            }
                            SxAddl.SxcCache.SxdVerSxMacro // DataFunctionalityLevel
                            -> data = byteArrayOf(1, 0, 0, 0, 0, 0)
                            SxAddl.SxcCache.SxdEnd -> data = byteArrayOf(0, 0, 0, 0, 0, 0)
                        }
                    }
                }
            }
            newData = ByteTools.append(data, newData)
            sxa.setData(newData)
            sxa.init()
            return sxa
        }
    }
}
