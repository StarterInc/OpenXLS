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
 * '
 * the CODENAME record stores tha name for a worksheet object.  It is not necessarily the same name as you see
 * in the worksheet tab, rather it is the VB identifier name!
 */
class Codename : io.starter.formats.XLS.XLSRecord() {

    private var stCodeName: String? = null

    internal var cch: Byte = 0
    internal var grbitChr: Byte = 0

    override fun init() {
        super.init()
        val wtf = this.getBytesAt(0, this.length)
        cch = this.getByteAt(0)
        grbitChr = this.getByteAt(2)

        val namebytes = this.getBytesAt(3, cch.toInt())

        try {
            if (grbitChr.toInt() == 0x1) {
                stCodeName = String(namebytes!!, WorkBookFactory.UNICODEENCODING)
            } else {
                stCodeName = String(namebytes!!, WorkBookFactory.DEFAULTENCODING)
            }
        } catch (e: UnsupportedEncodingException) {
            Logger.logWarn("UnsupportedEncodingException in setting codename: $e")
        }

    }

    fun setName(newname: String) {
        var modnamelen = 0
        var oldnamelen = 0
        if (grbitChr.toInt() == 0x0) {
            oldnamelen = stCodeName!!.length
        } else {
            oldnamelen = stCodeName!!.length * 2
        }

        cch = newname.length.toByte()
        var namebytes = newname.toByteArray()
        // if (!ByteTools.isUnicode(namebytes)){
        if (!ByteTools.isUnicode(newname)) {
            grbitChr = 0x0
            modnamelen = newname.length
        } else {
            grbitChr = 0x1
            modnamelen = newname.length * 2
        }
        val newdata = ByteArray(this.getData()!!.size - oldnamelen + modnamelen)
        try {
            if (grbitChr.toInt() == 0x1) {
                namebytes = newname.toByteArray(charset(WorkBookFactory.UNICODEENCODING))
            } else {
                namebytes = newname.toByteArray(charset(WorkBookFactory.DEFAULTENCODING))
            }
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("UnsupportedEncodingException in setting sheet name: $e")
        }

        System.arraycopy(namebytes, 0, newdata, 3, namebytes.size)
        newdata[0] = cch
        newdata[2] = grbitChr
        this.setData(newdata)
        this.init()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8327865068784623792L
    }

}
