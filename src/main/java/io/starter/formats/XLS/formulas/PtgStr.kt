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
package io.starter.formats.XLS.formulas

import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException


/**
 * PTG that stores a unicode string
 *
 *
 * Offset  Name       Size     Contents
 * ------------------------------------
 * 0       cch          1      Length of the string
 * 1       rgch         var    The string
 *
 *
 * *  I think the string includes a grbit itself, see UnicodeString.  Internationalization issues
 * may exist here!!!
 *
 *
 * -- Yes, it did include grbit, all handled now.
 *
 * @see Ptg
 *
 * @see Formula
 */
class PtgStr : GenericPtg, Ptg {

    override val isOperand: Boolean
        get() = true

    internal var cch: Short = 0
    internal var grbit: Byte = 0
    internal var negativeCch = false

    override// hits on Japanese strings in formulas
    val string: String
        get() {
            var strVal: String? = null
            try {
                if (grbit and 0x1 == 0x1) {
                    val barr = ByteArray(cch * 2)
                    System.arraycopy(record, 3, barr, 0, cch * 2)
                    strVal = String(barr, XLSConstants.UNICODEENCODING)
                } else {
                    val barr = ByteArray(cch)
                    System.arraycopy(record, 3, barr, 0, cch.toInt())
                    strVal = String(barr, XLSConstants.DEFAULTENCODING)
                }
            } catch (e: Exception) {
                val barr = ByteArray(cch)
                System.arraycopy(record, 3, barr, 0, cch.toInt())
                strVal = String(barr)
            }

            return strVal
        }

    /**
     * return the human-readable String representation of
     */
    override val textString: String
        get() {
            try {
                val d = Double(string)
            } catch (e: NumberFormatException) {
            }

            return "\"" + string + "\""
        }

    override val value: Any?
        get() = string

    var `val`: String
        get() = string
        set(s) {
            tempstr = s
            this.updateRecord()
        }

    private var tempstr: String? = null

    override val length: Int
        get() = record.size

    override fun toString(): String {
        return string
    }

    constructor() {
        // default constructor
    }

    constructor(s: String) {
        opcode = 0x17
        `val` = s
    }

    override fun init(b: ByteArray) {
        grbit = b[2]
        cch = (b[1] and 0xff).toShort() // this is the cch
        opcode = b[0]
        record = b
        this.populateVals()
    }

    /**
     * Constructer to create these on the fly, this is needed
     * for value storage in calculations of formulas.
     */
    private fun populateVals() {
        // no longer does anything, no String value stored
    }

    override fun updateRecord() {
        val ts = tempstr ?: return

        if (ByteTools.isUnicode(ts)) {
            grbit = (grbit or 0x1).toByte()
        }
        try {
            var strbytes: ByteArray? = null
            if (grbit and 0x1 == 0x1) {
                strbytes = ts.toByteArray(charset(XLSConstants.UNICODEENCODING))
            } else {
                strbytes = ts.toByteArray(charset(XLSConstants.DEFAULTENCODING))
            }

            val strbytelen = strbytes!!.size.toShort()
            cch = strbytelen
            if (grbit and 0x1 == 0x1) cch = (strbytelen / 2).toShort()
            //cch = (short)( getString().length() + 3);
            record = ByteArray(strbytelen + 3)
            //record = new byte[(cch*times) + 3];
            record[0] = 0x17
            record[1] = cch.toByte()
            record[2] = grbit
            System.arraycopy(strbytes, 0, record, 3, strbytelen.toInt())

        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("decoding formula string failed: $e")
        }

    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1427051673654768400L
    }


}