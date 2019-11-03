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

import java.io.Serializable

interface ColumnRange {
    /**
     * Gets the first column in the range.
     */
    val colFirst: Int

    /**
     * Gets the last column in the range.
     */
    val colLast: Int

    /**
     * determines if its a single column
     */
    val isSingleCol: Boolean

    /**
     * A lightweight immutable `ColumnRange` useful as a reference
     * element for searching collections.
     */
    class Reference(override val colFirst: Int, override val colLast: Int) : ColumnRange, Serializable {

        override val isSingleCol: Boolean
            get() = colFirst == colLast

        companion object {
            private const val serialVersionUID = -2240322394559418980L
        }
    }

    class Comparator : java.util.Comparator<ColumnRange>, Serializable {

        /**
         * This comparator will return equal if one of the column
         * ranges passed in is a single column reference and lies within
         * the bounds of the second column reference.  This is required for equals
         * within the cell collection.  If it turns out to be an issue (colinfos?) we should
         * separate this out.
         */
        override fun compare(r1: ColumnRange, r2: ColumnRange): Int {
            val single1 = r1.isSingleCol
            val single2 = r2.isSingleCol

            // if we're comparing a single column to a range
            if ((single1 || single2) && single1 != single2) { // XOR
                val range = if (single1) r2 else r1
                val single = if (single1) r1 else r2

                // and the single column falls within the range
                if (range.colFirst <= single.colFirst && range.colLast >= single.colLast) {

                    // if it's a range boundary, it chooses what it is
                    if (single is CellAddressible.RangeBoundary) {
                        val value = single.comareToRange()

                        // it needs to be reversed if the range was first
                        return if (single2) -value else value
                    }

                    // otherwise it's equal
                    return 0
                }
            }

            val first = r1.colFirst - r2.colFirst
            return if (0 != first) first else r2.colLast - r1.colLast
        }

        companion object {
            private const val serialVersionUID = -4506187924019516336L
        }
    }
}
