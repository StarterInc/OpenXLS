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

import io.starter.toolkit.ByteTools

/**
 * **SeriesList: Specifies the Series in an Overlay Chart (0x1016)**
 *
 *
 * bytes - 2	- nseries following
 * 2 * nseries = An array of 2-byte unsigned integers,
 * each of which specifies a one-based index of a Series record
 * in the collection of Series records in the current chart sheet substream
 */
class SeriesList : GenericChartObject(), ChartObject {
    internal var seriesmap: IntArray? = null

    /**
     * return the series mappings for the associated overlay chart
     * <br></br>series mappings links the overlay chart to the absolute series number
     * (determined by the actual order of the series in the chart array structure)
     *
     * @return
     */
    /**
     * set the series mappings for the associated overlay chart
     * <br></br>series mappings links the overlay chart to the absolute series number
     * (determined by the actual order of the series in the chart array structure)
     *
     * @param seriesmap
     */
    var seriesMappings: IntArray?
        get() = seriesmap
        set(smap) {
            val nseries = smap.size.toShort()
            seriesmap = IntArray(nseries)
            val data = ByteArray((nseries + 1) * 2)
            var b = ByteTools.shortToLEBytes(nseries)
            data[0] = b[0]
            data[1] = b[1]
            for (i in 0 until nseries) {
                val idx = (i + 1) * 2
                seriesmap[i] = smap[i]
                b = ByteTools.shortToLEBytes(smap[i].toShort())
                data[idx] = b[0]
                data[idx + 1] = b[1]
            }
            setData(data)
        }

    override fun init() {
        super.init()
        val nseries = ByteTools.readShort(this.data!![0].toInt(), this.data!![1].toInt()).toInt()
        seriesmap = IntArray(nseries)
        for (i in 0 until nseries) {
            val idx = (i + 1) * 2
            seriesmap[i] = ByteTools.readShort(this.data!![idx].toInt(), this.data!![idx + 1].toInt()).toInt()
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7852050067799624402L
    }
}
