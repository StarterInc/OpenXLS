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


/**
 * **Footerrec: Print Footer on Each Page (15h)**<br></br>
 *
 *
 * Footerrec describes the header printed on every page
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       cch             1       Length of the Footer String
 * 5       rgch            var     Footer String
 *
</pre> *
 */
class Footerrec : io.starter.formats.XLS.XLSRecord() {
    internal var cch = -1
    internal var rgch: String? = ""
    internal var DEBUG = false

    /**
     * get the footer text
     */
    /**
     * set the footer text
     * The key here is that we need to construct an excel formatted unicode string, so there will be 2 differing length fields.
     *
     *
     * Yes, this will probably be an issue in Japan some day....
     */
    var footerText: String?
        get() = rgch
        set(t) {
            try {
                if (ByteTools.isUnicode(t)) {
                    val bts = t.toByteArray(charset(XLSConstants.UNICODEENCODING))
                    cch = bts.size / 2
                    val newbytes = ByteArray(cch + 3)
                    val cchx = io.starter.toolkit.ByteTools.shortToLEBytes(cch.toShort())
                    newbytes[0] = cchx[0]
                    newbytes[1] = cchx[1]
                    newbytes[2] = 0x1
                    System.arraycopy(bts, 0, newbytes, 3, bts.size)
                    this.setData(newbytes)
                } else {
                    val bts = t.toByteArray(charset(XLSConstants.DEFAULTENCODING))
                    cch = bts.size
                    val newbytes = ByteArray(cch + 3)
                    val cchx = io.starter.toolkit.ByteTools.shortToLEBytes(cch.toShort())
                    newbytes[0] = cchx[0]
                    newbytes[1] = cchx[1]
                    newbytes[2] = 0x0
                    System.arraycopy(bts, 0, newbytes, 3, bts.size)
                    this.setData(newbytes)
                }
            } catch (e: Exception) {
                Logger.logInfo("setting Footer text failed: $e")
            }

            this.rgch = t
        }


    override fun setSheet(bs: Sheet?) {
        super.setSheet(bs)
        bs!!.setFooter(this)
    }

    override fun init() {
        super.init()
        if (this.length > 4) {
            val cch = this.getByteAt(0).toInt()
            val namebytes = this.getBytesAt(0, this.length - 4)
            val fstr = Unicodestring()
            fstr.init(namebytes!!, false)
            rgch = fstr.toString()
            if (DEBUGLEVEL > XLSConstants.DEBUG_LOW)
                Logger.logInfo("Footer text: " + this.footerText!!)
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 227652250172483965L
    }

}