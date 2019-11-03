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
/**
 *
 */
package io.starter.formats.XLS

import io.starter.toolkit.ByteTools

/**
 * **XCT  CRN Count (0059h)**<br></br>
 *
 *
 * <pre>
 * This record stores the number of immediately following Crn records.
 * These records are used to store the cell contents of external references.
 *
 * offset  size 	contents
 * ---
 * 0 		2 		Number of following CRN records
 * 2 		2 		Index into sheet table of the involved SUPBOOK record
</pre> *
 *
 */

class Xct : XLSRecord() {
    private var nCRNs: Int = 0    // number of External Cell References CRN record, similar to EXTERNNAME
    private var supBookIndex: Int = 0

    override fun init() {
        super.init()
        nCRNs = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        supBookIndex = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
    }

    override fun toString(): String {
        return "XTC: nCRNS=$nCRNs SupBook Index: $supBookIndex"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -5112701341711255711L
    }
}
