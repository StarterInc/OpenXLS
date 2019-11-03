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
import java.util.Comparator

/**
 * A `Comparator` that sorts cell records by address.
 * As `XLSRecord` currently uses default equality the
 * imposed ordering is inconsistent with equals.
 */
class CellAddressComparator : Comparator<*>, Serializable {

    /**
     * Compares its two arguments for order.
     * The arguments must be `XLSRecord`s that represent cells
     * (`isValueForCell()` must return true).
     *
     * @param o1 the first cell record to be compared
     * @param o2 the second cell record to be compared
     * @throws ClassCastException if either argument is not a cell record
     */
    override fun compare(o1: Any?, o2: Any?): Int {
        if (o1 == null || o1 !is XLSRecord
                || o2 == null || o2 !is XLSRecord)
            throw ClassCastException()

        val rec1 = o1 as XLSRecord?
        val rec2 = o2 as XLSRecord?

        if (!rec1!!.isValueForCell || !rec2!!.isValueForCell)
            throw ClassCastException()

        val diff = rec1.rowNumber - rec2.rowNumber
        return if (diff != 0) diff else rec1.colNumber - rec2.colNumber

    }

    /**
     * Indicates whether another `Comparator` imposes the same
     * ordering as this one. As all instances of this class impose the same
     * order, this method returns true if and only if the given object is
     * an instance of this class.
     */
    override fun equals(obj: Any?): Boolean {
        return obj != null && this.javaClass == obj.javaClass
    }

    /**
     * Returns a hash code value for this object.
     * As all instances of this class are equivalent this returns a
     * constant value.
     */
    override fun hashCode(): Int {
        // This is just a random number
        return 973216835
    }

    companion object {
        private const val serialVersionUID = 686639297047358268L
    }
}