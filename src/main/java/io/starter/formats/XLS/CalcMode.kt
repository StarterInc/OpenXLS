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
 * **CALCMODE: (OxD)**<br></br>
 *
 *
 * It specifies whether to calculate formulas manually,
 * automatically or automatically except for multiple table operations.
 *
 * <pre>
 * Offset Size Contents
 * 0 		2 	FFFFH = automatically except for multiple table operations
 * 0000H = manually
 * 0001H = automatically (default)
</pre> *
 */

class CalcMode : io.starter.formats.XLS.XLSRecord() {
    internal var calcmode: Short = 0

    /**
     * returns the recalculation mode:
     * 0= Manual, 1= Automatic, 2= Automatic except for Multiple Table Operations
     *
     * @return int recalculation mode
     */
    val recalcuationMode: Int
        get() = if (calcmode < 0) 2 else calcmode.toInt()

    override fun init() {
        super.init()
        /*  FFFFH = automatically except for multiple table operations
			0000H = manually
			0001H = automatically (default)
         */
        calcmode = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    /**
     * Sets the recalculation mode for the Workbook:
     * <br></br>0= Manual
     * <br></br>1= Automatic
     * <br></br>2= Automatic except for multiple table operations
     */
    fun setRecalculationMode(mode: Int) {
        var mode = mode
        if (mode >= 0 && mode <= 2) {
            if (mode == 2) mode = -1
            calcmode = mode.toShort()
            val b = ByteTools.shortToLEBytes(calcmode)
            this.getData()[0] = b[0]
            this.getData()[1] = b[1]
        }
    }

    companion object {

        private val serialVersionUID = -4544323710670598072L
    }

}