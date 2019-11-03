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

import java.util.Arrays

/**
 * The SXDBB record specifies the values of all the cache fields that have a
 * fAllAtoms field of the SXFDB record equal to 1 and that correspond to source
 * data entities, as specified by cache fields, for a single cache record.
 *
 *
 * blob (var) blob (variable): An array of 1-byte and 2-byte unsigned integers
 * that specifies indexes to cache items of cache fields that correspond to
 * source data entities, as specified by cache fields, that have an fAllAtoms
 * field of the SXFDB record equal to 1. The order of the indexes specified in
 * the array corresponds to the order of the cache fields as they appear in the
 * PivotCache. Each unsigned integer specifies a zero-based index of a record in
 * the sequence of records that conforms to the SRCSXOPER rule of the associated
 * cache field. The referenced record from the SRCSXOPER rule specifies a cache
 * item that specifies a value for the associated cache field. If the
 * fShortIitms field of an SXFDB record of the cache field equals 1, the index
 * value for this cache field is stored in this field in two bytes; otherwise,
 * the index value is stored in this field in a single byte.
 */

class SxDBB : XLSRecord(), XLSConstants, PivotCacheRecord {
    /**
     * retrieve the cache item indexes for this cache record (or row)
     *
     * @return
     */
    var cacheItemIndexes: ShortArray
        internal set

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

    // TODO: handle > 255 cache items (see SxFDB fShortItms)
    override fun init() {
        super.init()
        val data = this.getData()
        cacheItemIndexes = ShortArray(data!!.size)
        for (i in data.indices) {
            cacheItemIndexes[i] = data[i].toShort()        // TODO: may also be two bytes
        }
        if (DEBUGLEVEL > 3)
            Logger.logInfo("SXDBB -" + Arrays.toString(cacheItemIndexes))
    }

    override fun toString(): String {
        return "SXDBB: " + Arrays.toString(cacheItemIndexes) +
                Arrays.toString(this.record)
    }

    /**
     * sets the cache item indexes for this cache record (or row)
     *
     * @param cacheitems
     */
    fun setCacheItemIndexes(cacheitems: ByteArray) {
        /*If the fShortIitms field of an SXFDB record of the cache field equals 1, the index
         * value for this cache field is stored in this field in two bytes; otherwise,
         * the index value is stored in this field in a single byte.*/
        // fShortItems means that  > 255 cache items -- assume
        this.setData(cacheitems)
        this.init()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 9027599480633995587L

        /**
         * create a new minimum SXDBB
         *
         * @return
         */
        // minimum (??)
        val prototype: XLSRecord?
            get() {
                val sxdbb = SxDBB()
                sxdbb.opcode = XLSConstants.SXDBB
                sxdbb.setData(byteArrayOf(0))
                sxdbb.init()
                return sxdbb
            }
    }
}
