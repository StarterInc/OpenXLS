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


/**
 * **Blank: a blank cell value 0x201**<br></br>
 *
 *
 * Blank records define a blank cell.  The rw field defines row number (0 based)
 * and the col field defines column number.
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2           Row
 * 6       col         2           Column
 * 8       ixfe        2           Index to the XF record
</pre> *
 */
class Blank
/**
 * Provide constructor which automatically
 * sets the body data and header info.  This
 * is needed by Mulblank which creates the Blanks without
 * the benefit of WorkBookFactory.parseRecord().
 */
@JvmOverloads internal constructor(b: ByteArray = byteArrayOf(0, 0, 0, 0, 0, 0)) : XLSCellRecord() {

    // return a blank string val
    override var stringVal: String
        get() = ""
        set


    init {
        setData(b)
        opcode = XLSConstants.BLANK
        length = 6.toShort()
        this.init()
    }


    override fun init() {
        super.init()
        var pos = 4
        super.initRowCol()
        ixfe = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()).toInt()
        this.isValueForCell = true
        this.isBlank = true
    }

    fun setCol(i: Int) {
        if (this.isValueForCell) {
            this.getData()
            if (data == null) setData(byteArrayOf(0, 0, 0, 0, 0, 0))
            val c = ByteTools.shortToLEBytes(i.toShort())
            System.arraycopy(c, 0, this.getData()!!, 2, 2)
        }
        colNumber = i.toShort()
    }

    /**
     * set the row
     */
    fun setRow(i: Int) {
        if (this.isValueForCell) {
            this.getData()
            if (data == null) setData(byteArrayOf(0, 0, 0, 0, 0, 0))
            val r = ByteTools.shortToLEBytes(i.toShort())
            System.arraycopy(r, 0, this.getData()!!, 0, 2)
        }
        rw = i
    }

    companion object {


        private val serialVersionUID = -3847009755105117050L

        //protected byte[] BLANK_CELL_BYTES = { 0, 0, 0, 0, 0, 0};

        val prototype: XLSRecord?
            get() {
                val bl = Blank()
                bl.setData(byteArrayOf(0, 0, 0, 0, 0, 0))
                return bl
            }
    }

}/*    setData(BLANK_CELL_BYTES);
        setOpcode(BLANK);
        setLength((short)6);
        this.init();
*/