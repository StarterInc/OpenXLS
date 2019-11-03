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
 * **DefColWidth: default column width for WorkBook (55h)**<br></br>
 *
 *
 * offset  name        size    contents
 * ---
 * 6       miyRw       2       Default Column Width
 *
 *
 *
 *
 * @see DefaultRowHeight
 */
class DefColWidth : io.starter.formats.XLS.XLSRecord() {
    var defaultWidth: Short = 0
        internal set

    override fun init() {
        val mydata = this.getData()
        defaultWidth = ByteTools.readShort(mydata!![0].toInt(), mydata[1].toInt())
    }


    fun setDefaultColWidth(t: Int) {
        this.defaultWidth = t.toShort()
        val mydata = this.getData()
        val heightbytes = ByteTools.shortToLEBytes(t.toShort())
        mydata[0] = heightbytes[0]
        mydata[1] = heightbytes[1]
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8726286841723548636L
    }
}