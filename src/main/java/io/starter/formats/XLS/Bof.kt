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

import io.starter.formats.LEO.InvalidFileException
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger


/**
 * <pre>**Bof: Beginning of File Stream 0x809**<br></br>
 *
 * Marks the beginning of an XLS file Stream including Boundsheets
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 0       vers        1       version:
 * 1       bof         1       0x09
 * ...
 *
</pre></pre> *
 */
class Bof : UnencryptedXLSRecord() {
    internal var grbit: Short = 0
    internal var xlsVersionString = ""
    internal var oldlen = -1

    /**
     * Set the offset for this BOF
     */
    override var offset: Int
        get
        set(s) {
            super.setOffset(s)
            if (worksheet != null)
                if (isSheetBof || isVBModuleBof || worksheet!!.myBof == this)
                    worksheet!!.lbPlyPos = this.lbPlyPos
        }

    internal val isValidBIFF8: Boolean
        get() = oldlen == 20

    /**
     * this is equal to the lbPlyPos stored in
     * the Boundsheet associated with this Bof
     */
    internal val lbPlyPos: Long
        get() = if (!isValidBIFF8) (offset + 8).toLong() else offset.toLong()

    /**
     * @return Returns the sheetbof.
     */
    val isSheetBof: Boolean
        get() = grbit and 0x10 == 0x10

    val isVBModuleBof: Boolean
        get() = grbit and 0x06 == 0x06

    /**
     * @return Returns the sheetbof.
     */
    val isChartBof: Boolean
        get() = grbit and 0x20 == 0x20

    override fun toString(): String {
        return super.toString() + " lbplypos: " + this.lbPlyPos
    }

    /**
     * Initialize the BOF record
     */
    override fun init() {
        super.init()
        this.getData()
        grbit = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        val compat = ByteTools.readInt(this.getByteAt(12), this.getByteAt(13), this.getByteAt(14), this.getByteAt(15)) // 1996
        xlsVersionString = compat.toString() + ""
        oldlen = this.length
        if (oldlen < 16) {
            Logger.logErr("Not Excel '97 (BIFF8) or later version.  Unsupported file format.")
            throw InvalidFileException("InvalidFileException: Not Excel '97 (BIFF8) or later version.  Unsupported file format.")
        }
    }

    /**
     * @param sheetbof The sheetbof to set.
     */
    fun setSheetBof() {
        grbit = 0x10
    }

    companion object {


        private val serialVersionUID = 3005631881544437570L
    }
}