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
import io.starter.toolkit.ByteTools

/**
 * **ObjectLink: Attaches Text to Chart or to Chart Item (0x1027)**
 *
 *
 * 4	wLinkObj	2		Object text is linked to (1= chart title, 2= Veritcal (y) axis title, 3= Category (x) axis title, 4= data series points, 7=Series Axis 12= Display Units
 * 6	wLinkVar1	2		0-based series number	(only if wLinkObj=4,  otherwise 0)
 * 8	wLinkVar2	2		0-based category number within the series specified by wLinkVar1.  (only if wLinkObj=4,  otherwise 0).  If attached to entire series rather
 * than a single data point, = 0xFFFF.
 */
class ObjectLink : GenericChartObject(), ChartObject {
    internal var wLinkObj: Short = 0

    /**
     * Does this object link refer to the chart title?
     *
     * @return
     */
    internal val isChartTitle: Boolean
        get() = wLinkObj.toInt() == TYPE_TITLE

    /**
     * Does this object link refer to the XAxis Label?
     *
     * @return
     */
    internal val isXAxisLabel: Boolean
        get() = wLinkObj.toInt() == TYPE_XAXIS


    /**
     * Does this object link refer to the YAxis Label?
     *
     * @return
     */
    internal val isYAxisLabel: Boolean
        get() = wLinkObj.toInt() == TYPE_YAXIS

    var type: Int
        get() = wLinkObj.toInt()
        set(type) {
            wLinkObj = type.toShort()
            this.data[0] = wLinkObj.toByte()
        }

    private val PROTOTYPE_BYTES = byteArrayOf(1, 0, 0, 0, 0, 0)

    override fun init() {
        super.init()
        val rkdata = this.data
        wLinkObj = ByteTools.readShort(rkdata!![0].toInt(), rkdata[1].toInt())
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -301929936750246017L
        val TYPE_TITLE = 1
        val TYPE_XAXIS = 3
        val TYPE_YAXIS = 2
        val TYPE_DATAPOINTS = 4
        val TYPE_ZAXIS = 7
        /**
         * An axis-formatting option that determines how numeric units are displayed on a value axis.
         */
        val TYPE_DISPLAYUNITS = 0xC

        fun getPrototype(type: Int): XLSRecord {
            val o = ObjectLink()
            o.opcode = XLSConstants.OBJECTLINK
            o.data = o.PROTOTYPE_BYTES
            o.type = type
            return o
        }
    }
}
