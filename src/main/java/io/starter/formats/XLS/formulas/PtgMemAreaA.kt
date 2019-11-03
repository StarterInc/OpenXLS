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
package io.starter.formats.XLS.formulas

import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools

/**
 * PtgMemArea is an optimization of referenced areas.  Sweet!
 *
 *
 * Like most optimizations it really sucks.  It is also one of the few Ptg's that
 * has a variable length.
 *
 *
 * Format of length section
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           (reserved)     4       Whatever it may be
 * 2           cce			   2	   length of the reference subexpression
 *
 * Format of reference Subexpression
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0			cref		2			The number of rectangles to follow
 * 2			rgref		var			An Array of rectangles
 *
 * Format of Rectangles
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rwFirst     2       The First row of the reference
 * 2           rwLast     2       The Last row of the reference
 * 4           ColFirst    1       (see following table)
 * 6           ColLast    1       (see following table)
</pre> *
 *
 * @see Ptg
 *
 * @see Formula
 */
open class PtgMemAreaA : PtgMemArea() {
    internal var cce = 0
    internal var cref = 0
    internal var areas: Array<MemArea>

    override val length: Int
        get() = -1

    /*
     *  return a string representation of all of the ranges, seperated by comma.
     */
    override val value: Any?
        get() {
            var res = ""
            for (i in areas.indices) {
                res += areas[i].string
                if (i != areas.size - 1) res += ","
            }
            return res
        }

    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    internal override fun populateVals() {
        cce = ByteTools.readInt(record[6].toInt(), record[5].toInt())
        cref = ByteTools.readInt(record[8].toInt(), record[7].toInt())
        areas = arrayOfNulls(cref)
        var holder = 9
        for (i in 0 until cref) {
            val arr = ByteArray(6)
            System.arraycopy(record, holder, arr, 0, 6)
            areas[i].init(arr)
            holder += 6
        }

    }

    /*
     * Describes a representation of an excel Reference.  This is not an actual
     * reference to the cells, just the description!
     *
     */
    private inner class MemArea {
        internal var rwFirst: Int = 0
        internal var rwLast: Int = 0
        internal var colFirst: Int = 0
        internal var colLast: Int = 0

        /*
         * returns a string representation of the area, or cell if only one.
         */
        internal// it is a single cell amoeba
        val string: String
            get() {
                if (rwFirst == rwLast && colFirst == colLast) {
                    var retstr = ExcelTools.getAlphaVal(colLast)
                    retstr = retstr + (rwLast + 1)
                    return retstr
                }
                val arr = intArrayOf(rwFirst, colFirst, rwLast, colLast)
                return ExcelTools.formatRangeRowCol(arr)
            }

        internal fun init(b: ByteArray) {
            rwFirst = ByteTools.readInt(b[0].toInt(), b[1].toInt())
            rwLast = ByteTools.readInt(b[2].toInt(), b[3].toInt())
            colFirst = b[4].toInt()
            colLast = b[5].toInt()
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5528547215693511069L
    }

}