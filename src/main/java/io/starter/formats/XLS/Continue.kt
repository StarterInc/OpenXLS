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


/**
 * **Continue: Continues Long Records (3Ch)**<br></br>
 * Records longer than 8228 must be split into several records.
 *
 *
 * These records contain only data.
 *
 *<pre>
 *
 * offset  name        size    contents
 * ---
 * 4       grBit       1       (SST ONLY) Whether the first UNICODESTRING
 * segment is compressed or uncompressed data
 * 5                   var     Continuation of record data.
 *
</pre> *
 *
 * @see Sst
 *
 * @see Txo
 */
/* more info:
 * When a SST record contains a string that is continued in a CONTINUE record,
 * the description of the CONTINUE record for a BIFF 8 (Excel 97, Excel 2000, and Excel 2002 Workbook,)
 * states that the record data continues at offset 4.
 * This omits any comment to the effect that at offset 4 there is a grbit field
 * holding a flag that describes the UNICODE state - compressed or uncompressed -
 * of the portion of the string that is continued beginning at offset 5.
 *
 * Where any character in the data segment requires Unicode high-order byte information,
 * the grbit flag will be 01h, and all characters in the string segment will be two-byte,
 * uncompressed Unicode.
 */
class Continue : io.starter.formats.XLS.XLSRecord() {
    /**
     * The SST record is the Shared String Table record,
     * and will contain the strings of text from cells of the worksheet.
     * For Excel 97, Excel 2000, Excel 2002 the size of this record is
     * limited to 8224 bytes of data, including the formatting runs
     * and string-length information. Shared string data that exceeds
     * this limit will be stored in a CONTINUE record. When the last
     * string in the SST must be broken into two segments, and the
     * last segment stored as the first data in the CONTINUE record,
     * that segment may be stored as either compressed or uncompressed
     * Unicode. Consequently, the first byte of the data will contain 00h or 01h.
     * This is a one-byte field called a grbit field. It is not part of the string segment.
     *
     *
     * The grbit flag value 00h says no bytes of the data need
     * Unicode high-order byte data, so all are stored as compressed Unicode
     * (all the high-order bytes of the Unicode representation of the data
     * characters have been stripped. They all contained 00h, so Excel manages
     * the logic of restoring that high-order information when it loads the record.)
     *
     *
     * Where any character in the data segment requires Unicode
     * high-order byte information, the grbit flag will be 01h,
     * and all characters in the string segment will be two-byte,
     * uncompressed Unicode.
     */

    internal var hasgrbit: Boolean? = null
    internal var maskedMso: MSODrawing? = null    // Continue's can mask Mso's-this links to Masked Mso rec created from this data; see ContinueHandler.addRec;

    var grbit: Byte = 0x0

    internal val isBigRecContinue: Boolean
        get() = if (this.predecessor == null) true else this.predecessor!!.length >= XLSConstants.MAXRECLEN

    /**
     * Set that this continue has a grbit.  I find this method very odd in that
     * when one is in a continue record, it backs off the amount of the continue offset to the beginning of the
     * sst record that is being handled.  In cases of encrypted workbooks this causes a failure.  I am not exactly enamored
     * of this fix, but the whole code section seems off.
     *
     * @param b
     */
    internal// ie Bit is set
    var hasGrbit: Boolean
        get() {
            if (this.hasgrbit != null) return hasgrbit!!.booleanValue()

            if (DEBUGLEVEL > 1) {
                Logger.logInfo("Grbit pos0: " + (getGrbit() and 0))
                Logger.logInfo("Grbit pos2: " + (getGrbit() and 4))
                Logger.logInfo("Grbit pos3: " + (getGrbit() and 8))
            }

            return if (getGrbit() and 8 != 0) true else getGrbit() < 0x2 && getGrbit() >= 0x0

        }
        set(b) {
            if (b && this.encryptedByteReader === this.byteReader)
                grbit = this.getByteAt(0)
            hasgrbit = java.lang.Boolean.valueOf(b)
        }

    private var mygr: Byte = 0

    internal var deldata: ByteArray? = null

    /**
     * @return
     */
    var predecessor: BiffRec? = null
        internal set

    var grbitoff = 0
    /**
     * @return
     */
    /**
     * @param i
     */
    var continueOffset = -1


    /**
     * Override this to not return the grbit as part of Continue data
     */
    override var data: ByteArray?
        get() {
            if (hasGrbit && !streaming) {
                super.getData()
                return this.getBytesAt(0, this.length - 4)
            }
            streaming = false
            return super.getData()
        }
        set

    internal var streaming = false

    internal fun getGrbit(): Byte {
        return if (this.data != null) this.data!![0] else super.getByteAt(0)
    }

    override fun init() {
        super.init()
        streaming = false
        mygr = super.getByteAt(0)
        if (DEBUGLEVEL > 2) Logger.logInfo(" init() GRBIT: $mygr")
    }


    override fun getByteAt(off: Int): Byte {
        var rpos = off + this.grbitoff
        if (rpos < 0) {
            if (DEBUGLEVEL > 5) Logger.logWarn("Continue pointer is: $rpos")
            rpos = 0
        }
        rpos -= continueOffset
        return super.getByteAt(rpos)

    }

    /**
     * set the streaming flag so we get the grbit in output
     */
    override fun preStream() {
        streaming = true
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -6303887828816619839L


        /**
         * creates and returns a basic Continues record which defines a text string
         *
         * @param txt - String txt to represent
         * @return new Continue record
         */
        fun getTextContinues(txt: String): Continue {
            val c = Continue()
            c.opcode = XLSConstants.CONTINUE
            val data = ByteArray(txt.toByteArray().size + 1)
            System.arraycopy(txt.toByteArray(), 0, data, 1, data.size - 1)
            c.setData(data)
            c.init()
            return c
        }

        /**
         * create and return the absolute minimum Continue record defining a formatting run
         *
         * @return new Continue record
         */
        /**
         * 0 2 First formatted character (zero-based)
         * 2 2 Index to FONT record (➜5.45)
         */// meaning: from character 0 to the end, use default font 0
        val basicFormattingRunContinues: Continue
            get() {
                val c = Continue()
                c.opcode = XLSConstants.CONTINUE
                c.setData(ByteArray(4))
                c.init()
                return c
            }
    }
}
