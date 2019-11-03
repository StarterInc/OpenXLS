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

// TODO: Handle other types of External Names besides Add-Ins

/**
 * **Externname: External Name Record (17h)**<br></br>
 *
 *
 * <pre>
 * for add-in formulas (1 for each add-in formula)
 *
 * offset  name        size    contents
 * ---
 * 0       Op Flags 		2       Always 00 for Add-ins
 * 2		Not used		4
 * 6		Name			var.	Unicode formula name (1st 2 bytes are length)
 * var.	#REF Err Code	4		always 02 00 1C 17  (2 0 28 23)
 *
 * for external names
 * offset  name        size    contents
 * ---
 * 0       Op Flags 		2       Always 00 for Add-ins
 * 2		INDEX			2		One-based index to EXTERNSHEET
 * 4		Not used		2
 * 6		Name			var.	External name
 * var.	Formula data	var.	RPN Token Array
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Externsheet
 */
class Externname : XLSRecord() {

    override var workBook: WorkBook?
        get
        set(bk) {
            super.workBook = bk
        }

    val externalNames: Array<String>
        get() = workBook!!.externalNames

    override fun init() {
        super.init()
        val externnametype = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        if (externnametype.toInt() == 0) {// add-in
            // read in length of external name
            val len = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
            val b = this.getBytesAt(8, len.toInt())
            val s = String(b!!)
            workBook!!.addExternalName(s)    // store external names in workbook
        }
    }

    fun getExternalName(t: Int): String {
        return workBook!!.getExternalName(t)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7153354861666069899L

        internal val ENDOFRECORDBYTES = byteArrayOf(0x2, 0x0, 0x1C, 0x17)
        internal val STRINGLENPOS = 6
        internal val STRINGPOS = 8
        internal val STATICPORTIONSIZE = 12 // header=6, endofrecord=4, strlength=2

        // TODO: Finish for other types of external names (???)
        fun getPrototype(s: String): XLSRecord {
            val x = Externname()
            val len = s.length
            x.length = (STATICPORTIONSIZE + len).toShort()
            x.opcode = XLSConstants.EXTERNNAME
            val dta = ByteArray(STATICPORTIONSIZE + len)
            // write string
            try {
                // write length
                val slen = ByteTools.shortToLEBytes(len.toShort())
                dta[STRINGLENPOS] = slen[0]
                dta[STRINGLENPOS + 1] = slen[1]
                // write string
                val bts = s.toByteArray()
                System.arraycopy(bts, 0, dta, STRINGPOS, bts.size)
                // write end of record
                System.arraycopy(ENDOFRECORDBYTES, 0, dta, STRINGPOS + len, ENDOFRECORDBYTES.size)

                x.setData(dta)
                x.originalsize = STATICPORTIONSIZE + len
            } catch (e: Exception) {
                Logger.logWarn("Exception excountered writing Externname: $e")
            }

            return x
            //            cch = bts.length/2;
            //           byte[] newbytes = new byte[cch+3];
            //            byte[] cchx = io.starter.toolkit.ByteTools.shortToLEBytes((short)cch);
            //           newbytes[0] = cchx[0];
            //          newbytes[1] = cchx[1];
            //         newbytes[2] = 0x1;
            //        System.arraycopy(bts,0,newbytes,3, bts.length);
            //       this.setData(newbytes);
        }
    }
}
