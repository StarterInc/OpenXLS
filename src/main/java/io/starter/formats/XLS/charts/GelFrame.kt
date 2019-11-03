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

import io.starter.OpenXLS.FormatHandle
import io.starter.formats.XLS.MSODrawingConstants
import io.starter.formats.escher.MsofbtOPT

/**
 * **GelFrame: Fill Data(0x1066)**
 * The GelFrame record specifies the properties of a fill pattern for parts of a chart.
 */
class GelFrame : GenericChartObject(), ChartObject {
    private var fillColor: java.awt.Color? = null
    private val fillType = 0    // default= solid

    override fun init() {
        super.init()
        // try to interpret
        val optrec = MsofbtOPT(MSODrawingConstants.MSOFBTOPT, 0, 3)    //version is always 3, inst is current count of properties.
        optrec.setData(this.data)    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
        fillColor = optrec.fillColor
    }

    override fun toString(): String {
        return "GelFrame: fillType=" + fillType + " fillColor:" + fillColor!!.toString()
    }

    /**
     * return the fill color for this frame
     *
     * @return Color Hex String
     */
    fun getFillColor(): String? {
        return if (fillColor == null) null else FormatHandle.colorToHexString(fillColor!!)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 581278144607124129L
    }
}
