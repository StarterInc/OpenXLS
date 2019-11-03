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
 * **PIVOTCHARTLINK: Pivot Chart Link (861h)**
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 * This record stores the link to a PivotTable for a PivotChart. Similar in function
 * to SXVIEWLINK but used only during copy & paste of a chart via BIFF.
 *
 *
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0861h
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		ai			var		same as AI record
 */
class PivotChartLink : GenericChartObject(), ChartObject {

    private val PROTOTYPE_BYTES = byteArrayOf()

    override fun init() {
        super.init()
    }    // TODO: Prototype Bytes

    private fun updateRecord() {}

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -6202325538826559210L

        val prototype: XLSRecord?
            get() {
                val pcl = PivotChartLink()
                pcl.opcode = XLSConstants.PIVOTCHARTLINK
                pcl.data = pcl.PROTOTYPE_BYTES
                pcl.init()
                return pcl
            }
    }

}
