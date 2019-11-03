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
 * **GUTS: Size of Row and Column Gutters**<br></br>
 *
 *
 * This record stores information about gutter settings
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4		dxRwGut			2		Size of the row gutter
 * 6		dyColGut		2		Size of the Col gutter
 * 8		iLevelRwMac		2		Maximum outline level for row
 * 10		iLevelColMac	2		Maximum outline level for col
 *
</pre> *
 */
class Guts : io.starter.formats.XLS.XLSRecord() {
    private var dxRwGut: Short = 0
    private var dyColGut: Short = 0
    private var iLevelRwMac: Short = 0
    private var iLevelColMac: Short = 0

    var rowGutterSize: Int
        get() = dxRwGut.toInt()
        set(i) {
            dxRwGut = i.toShort()
            updateRecBody()
        }

    var colGutterSize: Int
        get() = dyColGut.toInt()
        set(i) {
            dyColGut = i.toShort()
            updateRecBody()
        }

    var maxRowLevel: Int
        get() = iLevelRwMac.toInt()
        set(i) {
            iLevelRwMac = i.toShort()
            updateRecBody()
        }

    var maxColLevel: Int
        get() = iLevelColMac.toInt()
        set(i) {
            iLevelColMac = i.toShort()
            updateRecBody()
        }

    override fun init() {
        super.init()
        dxRwGut = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        dyColGut = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        iLevelRwMac = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        iLevelColMac = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())

        if (DEBUGLEVEL > 5)
            Logger.logInfo(
                    "INFO: Guts settings: dxRwGut:"
                            + dxRwGut
                            + " dyColGut:"
                            + dyColGut
                            + " iLevelRwMac:"
                            + iLevelRwMac
                            + " iLevelColMac:"
                            + iLevelColMac)
    }

    private fun updateRecBody() {
        var newbytes = ByteTools.shortToLEBytes(dxRwGut)
        newbytes = ByteTools.append(ByteTools.shortToLEBytes(dyColGut), newbytes)
        newbytes = ByteTools.append(ByteTools.shortToLEBytes(iLevelRwMac), newbytes)
        newbytes = ByteTools.append(ByteTools.shortToLEBytes(iLevelColMac), newbytes)
        this.setData(newbytes)

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2815489536116897500L
    }
}
