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
package io.starter.formats.XLS.charts

import io.starter.formats.XLS.WorkBookFactory
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException


/**
 * **SeriesText: Chart Legend/Category/Value Text Definition (100Dh)**<br></br>
 *
 *
 * This record defines the SeriesText data of a chart.
 *
 *
 * sdtX and sdtY fields determine data type (numeric and text)
 *
 *
 * cValx and cValy fields determine number of cells in series
 *
 *
 * Offset           Name    Size    Contents
 * --
 * 4               id    	2       Text identifier (should be zero)
 * 8               cch     2       length of String text
 * 10              rgch    2       String text
 *
 *
 *
 * @see Chart
 */

class SeriesText : GenericChartObject(), ChartObject {
    protected var id = -1
    protected var cch = -1
    private var text = ""

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 7, 1, 74, 0, 97, 0, 110, 0, 117, 0, 97, 0, 114, 0, 121, 0)

    fun setText(t: String) {
        // create a new SeriesText value from the passed-in String
        var strbytes: ByteArray? = null
        var uni: Byte = 0x0
        var lent = 0
        try {
            strbytes = t.toByteArray(charset(WorkBookFactory.UNICODEENCODING))
            uni = 0x1
            lent = strbytes!!.size / 2
        } catch (e: Exception) {
            strbytes = t.toByteArray()
            lent = strbytes!!.size
        }

        val newbytes = ByteArray(strbytes!!.size + 4)
        //		byte[] lenbytes = ByteTools.shortToLEBytes((short)strbytes.length);
        newbytes[0] = 0x0
        newbytes[1] = 0x0
        newbytes[2] = lent.toByte()
        newbytes[3] = uni
        System.arraycopy(strbytes, 0, newbytes, 4, strbytes.size)
        this.data = newbytes
        this.text = t
    }

    override fun init() {
        super.init()
        //byte[] data = this.getData();
        var multi = 2
        if (this.getByteAt(3).toInt() == 0x0) multi = 1
        cch = this.getByteAt(2).toInt() * multi
        if (cch < 0) cch *= -1 // strangely it can be negative...
        try {
            val namebytes = this.getBytesAt(4, cch)
            try {
                text = String(namebytes!!, WorkBookFactory.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("Unsupported Encoding error in SeriesText: $e")
            }

            if (DEBUGLEVEL > 10) Logger.logInfo("Series Text Value: $text")

        } catch (ex: Exception) {
            Logger.logWarn("SeriesText.init failed: $ex")
        }

        //Logger.logInfo("Initialized SeriesText: "+ text);
    }

    override fun toString(): String {
        return this.text
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3794355940075116165L

        fun getPrototype(text: String): SeriesText {
            val st = SeriesText()
            st.opcode = XLSConstants.SERIESTEXT
            st.data = st.PROTOTYPE_BYTES
            st.init()
            st.setText(text)
            return st
        }
    }
}
