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

import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException


/**
 * **WRITEACCESS 0x5C: Contains name of Excel installed user.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       stName      112     User Name as unformatted Unicodestring
 *
</pre> *
 *
 * @see WorkBook
 */

class Writeaccess : io.starter.formats.XLS.XLSRecord() {
    private var strname: Unicodestring? = null

    /**
     * set the Writeaccess username
     */
    var name: String?
        get() = strname!!.toString()
        set(str) = try {
            val nameb = str.toByteArray(charset(XLSConstants.DEFAULTENCODING))
            val newb = ByteArray(112)
            val diff = 112 - nameb.size
            if (diff < 0)
                System.arraycopy(nameb, 0, newb, 0, 112)
            else
                System.arraycopy(nameb, 0, newb, 0, nameb.size)
            strname!!.init(newb, false)
            this.setData(newb)
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("setting name in Writeaccess record failed: $e")
        }

    override fun init() {
        super.init()
        strname = Unicodestring()
        strname!!.init(bytes!!, false)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8868603864018600260L
    }
}