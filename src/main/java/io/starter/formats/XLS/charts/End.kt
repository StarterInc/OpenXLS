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

package io.starter.formats.XLS.charts

import io.starter.formats.XLS.XLSRecord

/**
 * ** End:  End of chart substream (0x1034)**
 *
 *
 * End is an identifier record for the chart record type.  There is no data to a end record,
 * and every end record must have a corrosponding begin record
 */
class End : GenericChartObject(), ChartObject {

    override fun init() {
        super.init()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 9022736093645720842L

        val prototype: XLSRecord?
            get() {
                val bl = End()
                bl.opcode = XLSConstants.END
                bl.data = byteArrayOf()
                return bl
            }
    }
}
