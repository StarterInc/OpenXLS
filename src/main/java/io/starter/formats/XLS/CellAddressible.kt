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

/**
 * Something with a cell address.
 */
interface CellAddressible : ColumnRange {

    /**
     * Returns the row the cell resides in.
     *
     * @return the zero-based index of the cell's parent row
     */
    val rowNumber: Int

    /**
     * A simple immutable cell reference useful as a map key.
     */
    open class Reference(override val rowNumber: Int, override val colFirst: Int, override val colLast: Int) : CellAddressible, Serializable {

        override val isSingleCol: Boolean
            get() = this.colFirst == this.colLast

        constructor(row: Int, col: Int) : this(row, col, col) {}

        companion object {
            private const val serialVersionUID = -9071483662123798966L
        }
    }

    /**
     * Cell reference for use as the boundary of a column range.
     */
    class RangeBoundary
    /**
     * @param before whether the boundary should sort before (true) or
     * after (false) a range which includes it
     */
    (row: Int, col: Int, private val before: Boolean) : Reference(row, col, col) {

        fun comareToRange(): Int {
            return if (before) -1 else 1
        }

        companion object {
            private val serialVersionUID = -1357617242449928095L
        }
    }

    /**
     * [Comparator] that sorts cells in ascending row major order.
     */
    class RowMajorComparator : java.util.Comparator<CellAddressible>, Serializable {

        override fun compare(cell1: CellAddressible, cell2: CellAddressible): Int {
            val rows = cell1.rowNumber - cell2.rowNumber
            return if (0 != rows) rows else colComp.compare(cell1, cell2)

        }

        companion object {
            private const val serialVersionUID = 5477030152120715766L
            private val colComp = ColumnRange.Comparator()
        }
    }

    /**
     * [Comparator] that sorts cells in ascending column major order.
     */
    class ColumnMajorComparator : java.util.Comparator<CellAddressible>, Serializable {

        override fun compare(cell1: CellAddressible, cell2: CellAddressible): Int {
            val cols = colComp.compare(cell1, cell2)
            return if (0 != cols) cols else cell1.rowNumber - cell2.rowNumber

        }

        companion object {
            private const val serialVersionUID = -1193867650674693873L
            private val colComp = ColumnRange.Comparator()
        }
    }
}
