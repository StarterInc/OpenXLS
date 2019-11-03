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


/**
 * **String: String Value of a Formula (207h)**<br></br>
 * When a formula evaluates to a string, a STRING record occurs after
 * the FORMULA record.  If the formula is part of an ARRAY, the STRING
 * occurs after the ARRAY record.
 *
 * <pre>
 * Offset          Name      Size      Contents
 * 4               rgch       var       String
 *
 * The String conforms to the information on Unicode Strings on page 264 of
 * the hard-copy of the Developer's Kit. Within the string you will see
 * this structure:
 *
 * Offset          Name      Size     Contents
 *
 * 0               cch       2         Count of characters in the string
 * 2               grbit     1         Option Flags
 * 3               rgb       var       Array of string characters and
 * formatting runs - if applicable
</pre> *
 *
 * @see Formula
 */
class StringRec : io.starter.formats.XLS.XLSRecord {
    internal var cch: Int = 0
    internal var rgch: ByteArray? = null
    internal var us: Unicodestring? = null

    override// why is this null???
    var stringVal: String?
        get() = if (us == null) null else us!!.toString()
        set(s) {
            us!!.updateUnicodeString(s)
            this.setData(us!!.read())
        }

    override val `val`: Any?
        get() = us!!.toString()

    override fun init() {
        super.init()
        us = Unicodestring()
        us!!.init(bytes!!, false)
        //this.setIsValueForCell(true);
        // will get "picked up" by formula due to WorkBook.addRecord code
        // isValueForCell will DELETE THE CELL if remove/change NOT WHAT WE WANT HERE
        this.isString = true
    }

    /**
     * Nullary constructor for init from bytes.
     * If this constructor is used [.init] must be called.
     */
    constructor() {}

    /**
     * Create a string record from scratch
     *
     * @param newValue the string this StringRec should represent
     */
    constructor(newValue: String) {
        this.opcode = XLSConstants.STRINGREC
        us = Sst.createUnicodeString(newValue, null, WorkBook.STRING_ENCODING_AUTO)
        this.setData(us!!.read())
        //this.setValueForCell(true); why ??? not true
    }

    override fun toString(): String? {
        return us!!.toString()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3377086771218110838L
    }
}