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
 * Record specifying the right margin of the sheet for printing.
 */
class RightMargin : XLSRecord() {

    internal var margin: Double = 0.toDouble()

    override fun init() {
        super.init()

        margin = ByteTools.eightBytetoLEDouble(getBytesAt(0, 8)!!)
    }

    init {
        this.opcode = XLSConstants.RIGHTMARGIN
        margin = 0.75    // default
        setData(ByteTools.doubleToLEByteArray(margin))
    }

    override fun setSheet(sheet: Sheet?) {
        super.setSheet(sheet)
        (sheet as Boundsheet).addPrintRec(this)
    }

    fun getMargin(): Double {
        return margin
    }

    fun setMargin(value: Double) {
        margin = value
        setData(ByteTools.doubleToLEByteArray(value))
    }

    companion object {
        private val serialVersionUID = -3649192673573344145L
    }
}
