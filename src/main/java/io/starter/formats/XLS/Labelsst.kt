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

import io.starter.OpenXLS.DateConverter
import io.starter.OpenXLS.WorkBookHandle
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.Calendar
import java.util.GregorianCalendar


/**
 * **Labelsst: BiffRec Value, String Constant/Sst 0xFD**<br></br>
 * The Labelsst record contains a string constant
 * from the Shared String Table (Sst).
 * The isst field contains a zero-based index into the shared string table
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number
 * 8       ixfe        2       Index to XF format record
 * 10      isst        4       Index into the Sst record
</pre> *
 *
 * @see Sst
 *
 * @see Labelsst
 *
 * @see Extsst
 */
class Labelsst : XLSCellRecord() {
    internal var isst: Int = 0


    private var unsharedstr: Unicodestring? = null

    val unsharedString: Unicodestring?
        get() {
            if (unsharedstr == null)
                this.initUnsharedString()
            return unsharedstr
        }

    /**
     * Returns the value of the Unicodestring
     * int the Shared String Table pointed to by this
     * LABELSst record.
     */
    /**
     * set a new value for the string
     */
    override// reset unsharedstr (see getStringVal) specifically to fix OOXML t="s" setStringVal
    //ensure reclen and datalen are maintained correctly:
    var stringVal: String?
        get() = if (unsharedstr != null)
            unsharedstr!!.toString()
        else
            this.workBook!!.sharedStringTable!!.getUStringAt(isst).toCachingString()
        set(v) {
            val ov = this.stringVal
            if (v == ov) return
            if (this.sheet!!.workBook!!.sharedStringTable!!.isSharedString(isst)) {
                isst = this.sheet!!.workBook!!.sharedStringTable!!.insertUnicodestring(v)
                System.arraycopy(ByteTools.cLongToLEBytes(isst), 0, getData()!!, 6, 4)
                init()
                val str = this.sheet!!.workBook!!.sharedStringTable!!.getUStringAt(isst)
                this.unsharedstr = str
            } else {
                val str = this.sheet!!.workBook!!.sharedStringTable!!.getUStringAt(isst)
                val origLen = str.length
                str.updateUnicodeString(v)
                val delta = str.length - origLen
                this.sheet!!.workBook!!.sharedStringTable!!.adjustSstLength(delta)
                this.unsharedstr = str
            }
        }

    /**
     * try to convert the String Value of this Labelsst record to an int
     * If it cannot be converted, returns NaN.
     */
    override var intVal: Int
        get() {
            val s = stringVal
            try {
                val i = Integer.valueOf(s!!)
                return i.toInt()
            } catch (n: NumberFormatException) {
                return java.lang.Float.NaN.toInt()
            }

        }
        set

    /**
     * try to convert the String Value of this Labelsst record to a double
     * If it cannot be converted,return NaN.
     */
    override// use it
    // fall through
    val dblVal: Double
        get() {
            val s = stringVal
            try {
                return Double(s!!)
            } catch (n: NumberFormatException) {
                this.xfRec
                if (myxf!!.isDatePattern) {
                    try {
                        val format = myxf!!.formatPattern
                        WorkBookHandle.simpledateformat.applyPattern(format!!)
                        val d = WorkBookHandle.simpledateformat.parse(s)
                        val c = GregorianCalendar()
                        c.time = d
                        return if (c == null) java.lang.Double.NaN else DateConverter.getXLSDateVal(c)
                    } catch (e: Exception) {
                    }

                }
                val c = DateConverter.convertStringToCalendar(s) ?: return java.lang.Double.NaN
                return DateConverter.getXLSDateVal(c)
            }

        }

    internal fun setIsst(i: Int) {
        isst = i
        System.arraycopy(ByteTools.cLongToLEBytes(isst), 0, this.getData()!!, 6, 4)
        try {
            this.workBook!!.sharedStringTable!!.initSharingOnStrings(isst)
        } catch (e: NullPointerException) {
        }

    }

    override fun init() {
        super.init()

        // this.initCacheBytes(0,10);
        // get the row, col and ixfe information
        super.initRowCol()
        val s = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        ixfe = s.toInt()
        // get the length of the string
        isst = ByteTools.readInt(
                this.getByteAt(6),
                this.getByteAt(7),
                this.getByteAt(8),
                this.getByteAt(9))
        this.isValueForCell = true
        this.isString = true
        this.resetCacheBytes()
        // init shared string info.
        if (isst != -1) {// not initialized - OOXML use - MUST be set later using setIsst
            try {
                this.workBook!!.sharedStringTable!!.initSharingOnStrings(isst)
            } catch (e: NullPointerException) {
                // nothing.  When adding new strings we have access issues, but it doesn't matter, we just care on book initialization for this..
            }

        }
    }


    internal fun initUnsharedString() {
        unsharedstr = this.workBook!!.sharedStringTable!!.getUStringAt(isst)
    }

    /**
     * Adds the LabelSST's string to the sst.
     *
     *
     * This is used when a worksheet is transferred over to a book that
     * does not contain it's entry in the sst.
     */
    internal fun insertUnsharedString(sst: Sst): Boolean {
        if (unsharedstr == null) {
            return false
        }
        this.isst = sst.insertUnicodestring(unsharedstr!!.toString())
        this.setIsst(isst)
        return true
    }

    /**
     * set this Label cell to a new Unicode string
     * Rich Unicode strings include formatting information
     */
    // 20090520 KSC: for OOXML, must use entire Unicode string so retain formatting info
    fun setStringVal(v: Unicodestring) {
        if (v == this.unsharedString)
            return
        isst = this.sheet!!.workBook!!.sharedStringTable!!.find(v)    // find this particular unicode string (including formatting)
        if (isst == -1)
            isst = this.sheet!!.workBook!!.sharedStringTable!!.insertUnicodestring(v)
        System.arraycopy(ByteTools.cLongToLEBytes(isst), 0, getData()!!, 6, 4)
        init()
        this.unsharedstr = v
    }


    /**
     * return string representation
     */
    override fun toString(): String {
        try {
            return "LABELSST:" + this.cellAddress + ":" + stringVal
        } catch (e: Exception) {
            Logger.logErr("Labelsst toString failed.", e)
            return "#ERR!"
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 467127849827595055L

        /**
         * Constructor which takes a number value
         * an Sst to store its Unicodestring in,
         * and returns an int offset to the string
         * in the Sst.
         */
        fun getPrototype(`val`: String?, bk: WorkBook): Labelsst {
            val retlab = Labelsst()
            // associate with the Sst
            retlab.originalsize = 10
            retlab.opcode = XLSConstants.LABELSST
            retlab.length = 10.toShort()
            retlab.setData(ByteArray(retlab.originalsize))
            //retlab.setDataContainsHeader(true);
            if (`val` != null) { // for XLSX handling ... label is linked to sst later
                // get the high Sst index, insert the new Unicodestring
                val sst = bk.sharedStringTable
                retlab.isst = sst!!.insertUnicodestring(`val`)
                System.arraycopy(ByteTools.cLongToLEBytes(retlab.isst), 0, retlab.getData()!!, 6, 4)
            } else
                retlab.isst = -1    // flag it's not set - MUST be set later
            retlab.getData()[4] = 0x0f
            retlab.workBook = bk
            retlab.init()
            return retlab
        }
    }
}
