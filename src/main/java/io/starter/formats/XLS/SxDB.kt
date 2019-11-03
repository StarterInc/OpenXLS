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
import java.util.Arrays

/**
 * SXDB		0xC6
 *
 *
 * The SXDB record specifies PivotCache properties.
 *
 *
 * crdbdb (4 bytes): A signed integer that specifies the number of cache records for this PivotCache.
 * MUST be greater than or equal to 0. MUST be 0 for OLAP PivotCaches. MUST be ignored if fSaveData is 0.
 *
 *
 * idstm (2 bytes): An unsigned integer that specifies the stream that contains the data for this PivotCache.
 * MUST be equal to the value of the idstm field of the SXStreamID record that specifies the PivotCache stream that contains this record.
 *
 *
 * A - fSaveData (1 bit): A bit that specifies whether cache records exist. MUST be 0 for OLAP PivotCaches.
 *
 *
 * B - fInvalid (1 bit): A bit that specifies whether the cache records are in the not-valid state.
 * MUST be equal to 1 if the PivotCache functionality level is greater than or equal to 3.
 * MUST be equal to 1 for OLAP PivotCaches. See cache records for more information.
 *
 *
 * C - fRefreshOnLoad (1 bit): A bit that specifies whether the PivotCache is refreshed on load.
 *
 *
 * D - fOptimizeCache (1 bit): A bit that specifies whether optimization is applied to the PivotCache to reduce memory usage.
 * MUST be 0 and MUST be ignored for a non-ODBC PivotCache.
 *
 *
 * E - fBackgroundQuery (1 bit): A bit that specifies whether the query used to refresh the PivotCache is executed asynchronously.
 * MUST be ignored if vsType not equals 0x0002.
 *
 *
 * F - fEnableRefresh (1 bit): A bit that specifies whether refresh of the PivotCache is enabled.
 * MUST be equal to 0 if the PivotCache functionality level is greater than or equal to 3.
 * MUST be equal to 0 for OLAP PivotCaches.
 *
 *
 * unused1 (10 bits): Undefined and MUST be ignored.
 *
 *
 * unused2 (2 bytes): Undefined and MUST be ignored.
 *
 *
 * cfdbdb (2 bytes): A signed integer that specifies the number of cache fields that corresponds to the source data.
 * MUST be greater than or equal to 0.
 *
 *
 * cfdbTot (2 bytes): A signed integer that specifies the number of cache fields in the PivotCache.
 * MUST be greater than or equal to 0.
 *
 *
 * crdbUsed (2 bytes): An unsigned integer that specifies the number of records used to calculate the PivotTable report.
 * Records excluded by PivotTable view filtering are not included in this value. MUST be 0 for OLAP PivotCaches.
 *
 *
 * vsType (2 bytes): An unsigned integer that specifies the type of source data.
 * MUST be equal to the value of the sxvs field of the SXVS record that follows the SXStreamID record that
 * specifies the PivotCache stream that contains this record.
 *
 *
 * cchWho (2 bytes): An unsigned integer that specifies the number of characters in rgb.
 * MUST be equal to 0xFFFF, or MUST be greater than or equal to 1 and less than or equal to 0x00FF.
 *
 *
 * rgb (variable): An optional XLUnicodeStringNoCch structure that specifies the name of the user who last refreshed the PivotCache.
 * MUST exist if and only if the value of cchWho is not equal to 0xFFFF.  If this field exists, the length MUST equal cchWho.
 * The length of this value MUST be less than 256 characters. The name is an application-specific setting that is not necessarily
 * related to the User Names StreamABNF.
 */

class SxDB : XLSRecord(), XLSConstants, PivotCacheRecord {
    private var crdbdb: Int = 0
    /**
     * returns the streamId -- index linked back to SxStreamID
     *
     * @return
     */
    var streamID: Short = 0
        private set
    private var grbit: Short = 0
    private var cfdbdb: Short = 0
    private var cfdbTot: Short = 0
    private var crdbUsed: Short = 0
    private var vsType: Short = 0
    private var cchWho: Short = 0
    private val fInvalid: Boolean = false
    private val fRefreshOnLoad: Boolean = false
    private val fEnableRefresh: Boolean = false    // significant bit fields
    private var rgb: String? = null


    private val PROTOTYPE_BYTES = byteArrayOf(1, 0, 0, 0, /* n cache records - minimum=1 */
            0, 0, /* stream id */
            33, 0, /* flags */
            -1, 31, /* unused */
            1, 0, /* cfdbdb */
            1, 0, /* cfdbTot */
            0, 0, /* crdbUsed */
            1, 0, /* vsType */
            -1, -1)    // cch

    /**
     * return the bytes describing this record, including the header
     *
     * @return
     */
    override val record: ByteArray
        get() {
            val b = ByteArray(4)
            System.arraycopy(ByteTools.shortToLEBytes(this.opcode), 0, b, 0, 2)
            System.arraycopy(ByteTools.shortToLEBytes(this.getData()!!.size.toShort()), 0, b, 2, 2)
            return ByteTools.append(this.getData(), b)
        }

    /**
     * returns the number of cache records (= number of non-header rows)
     *
     * @return
     */
    /**
     * sets the number of cache records (= number of non-header rows)
     *
     * @param n
     */
    //		crdbUsed= (short)n;		// TODO: filtering affects this!
    //		System.arraycopy( ByteTools.shortToLEBytes(crdbUsed), 0, this.getData(), 14, 2);
    var nCacheRecords: Int
        get() = crdbdb
        set(n) {
            crdbdb = n.toShort().toInt()
            System.arraycopy(ByteTools.cLongToLEBytes(crdbdb), 0, this.getData()!!, 0, 4)
        }

    /**
     * returns the number of cache fields (==columns) for the pivot cache
     */
    /**
     * sets the number of cache fields (==columns) for the pivot cache
     *
     * @param n
     */
    // TODO: doesn't cfdbTot==cfdbdb always????????????
    var nCacheFields: Int
        get() = cfdbdb.toInt()
        set(n) {
            cfdbdb = n.toShort()
            var b = ByteTools.shortToLEBytes(cfdbdb)
            this.getData()[10] = b[0]
            this.getData()[11] = b[1]
            cfdbTot = n.toShort()
            b = ByteTools.shortToLEBytes(cfdbTot)
            this.getData()[12] = b[0]
            this.getData()[13] = b[1]
        }

    // TODO: doesnt cfdbdb==cfdbTot always??
    // TODO: cfdbUsed == filtering ****
    // TODO: flags
    // TODO: cfWho
    override fun init() {
        super.init()
        if (DEBUGLEVEL > 3)
            Logger.logInfo("SXDB -" + Arrays.toString(this.getData()))
        crdbdb = ByteTools.readInt(this.getBytesAt(0, 4)!!)                    // # cache records
        streamID = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())    // streamid
        grbit = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())    //
        // 8,9 = unused
        cfdbdb = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())    // # cache fields
        cfdbTot = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        crdbUsed = ByteTools.readShort(this.getByteAt(14).toInt(), this.getByteAt(15).toInt())    // # used - filtering
        // KSC: TESTING
        //if (cfdbdb!=cfdbTot)
        //Logger.logWarn("SXDB: all cache items are not being used");
        vsType = ByteTools.readShort(this.getByteAt(16).toInt(), this.getByteAt(17).toInt())
        cchWho = ByteTools.readShort(this.getByteAt(18).toInt(), this.getByteAt(19).toInt())
        if (cchWho > 0) {
            val encoding = this.getByteAt(20)

            val tmp = this.getBytesAt(21, cchWho * (encoding + 1))
            try {
                if (encoding.toInt() == 0)
                    rgb = String(tmp!!, XLSConstants.DEFAULTENCODING)
                else
                    rgb = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("SxDB.init: $e")
            }

        }
    }

    override fun toString(): String {
        return "SXDB: nCacheRecords/rows:" + crdbdb +
                " nCacheFields:" + cfdbdb +
                " cfdbTot:" + cfdbTot +
                " sid:" + streamID +
                " vsType: " + vsType +
                Arrays.toString(this.record)
    }

    /**
     * sets the streamId -- index linked back to SxStreamID
     *
     * @param sid
     */
    fun setStreamID(sid: Int) {
        streamID = sid.toShort()
        val b = ByteTools.shortToLEBytes(streamID)
        this.getData()[4] = b[0]
        this.getData()[5] = b[1]
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 9027599480633995587L

        /**
         * create a new minimum SXDB
         *
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val sxdb = SxDB()
                sxdb.opcode = XLSConstants.SXDB
                sxdb.setData(sxdb.PROTOTYPE_BYTES)
                sxdb.init()
                return sxdb
            }
    }
}