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
 * **DefaultRowHeight: default row height for WorkBook (225h)**<br></br>
 *
 *
 * offset  name        size    contents
 * ---
 * 4       options     2       comments
 * 6       miyRw       2       Default height for unused rows, in twips = 1/20 of a point
 *
 *
 * ﻿	0 2 	2 2 	Option flags:
 * Bit	Mask Contents
 * 0 0001H 1 = Row height and default font height do not match 1
 * 0002H 1 = Row is hidden 2
 * 0004H 1 = Additional space above the row 3
 * 0008H 1 = Additional space below the row
 *
 *
 *
 *
 *
 * @see DefColWidth
 */
class DefaultRowHeight : io.starter.formats.XLS.XLSRecord() {
    internal var rwh: Short = 0

    override var workBook: WorkBook?
        get
        set(b) {
            super.workBook = b
            b!!.setDefaultRowHeightRec(this)
        }

    /**
     * init: save the default row height
     */
    override fun init() {
        super.init()
        rwh = ByteTools.readShort(this.getData()!![2].toInt(), this.getData()!![3].toInt())
    }

    /**
     * set the default row height in twips (=1/20 of a point)
     * <br></br>Twips are 20*Excel units
     * <br></br>e.g. default row height in Excel units=12.75
     * <br></br>20*12.75= 256 (approx) twips
     *
     * @param t - desired default row height in twips
     */
    fun setDefaultRowHeight(t: Int) {
        this.rwh = t.toShort()
        val mydata = this.getData()
        val heightbytes = ByteTools.shortToLEBytes(this.rwh)
        mydata[2] = heightbytes[0]
        mydata[3] = heightbytes[1]
    }

    /**
     * set the sheet's default row height in Excel units or twips
     */
    override fun setSheet(bs: Sheet?) {
        this.worksheet = bs
        (bs as Boundsheet).defaultRowHeight = this.rwh / 20.0
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -930064032441287284L
    }

}