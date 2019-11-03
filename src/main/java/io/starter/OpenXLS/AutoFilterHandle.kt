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
package io.starter.OpenXLS

import io.starter.formats.XLS.AutoFilter

/**
 * AutoFilterHandle allows for manipulation of the AutoFilters within the
 * Spreadsheet
 *
 * @author John McMahon
 */
class AutoFilterHandle
/**
 * For internal use only. Creates an AutoFilter Handle based on the AutoFilter
 * passed in
 *
 * @param AutoFilter af
 */
(af: AutoFilter) : Handle {
    private val af: AutoFilter? = null

    /**
     * returns the column this AutoFilter is applied to <br></br>
     * NOTE: this may not be 100% exact
     *
     * @return in column number
     */
    val col: Int
        get() = if (this.af != null) af.col else -1

    /**
     * returns the String representation of the comparison value for this
     * AutoFilter, if any
     *
     * <br></br>
     * This will return the comparison value of the first condition for those
     * AutoFilters containing two conditions
     *
     * @return String comparison value for the second condition or null if none
     * exists
     * @see getVal2
     */
    val `val`: String?
        get() = if (af != null) af.`val` as String else null

    /**
     * returns the String representation of the second comparison value for this
     * AutoFilter, if any
     *
     * <br></br>
     * This will return the comparison value of the second condition for those
     * AutoFilters containing two conditions
     *
     * @return String comparison value for the second condition or null if none
     * exists
     * @see getVal
     */
    val val2: String?
        get() = if (af != null) af.val2 as String else null

    /**
     * get the operator associated with this AutoFilter <br></br>
     * NOTE: this will return the operator in the first condition if this AutoFilter
     * contains two conditions <br></br>
     * Use getOp2 to retrieve the second condition operator
     *
     * @return String operator
     */
    val op: String?
        get() = af?.op

    /**
     * returns true if this AutoFilter is set to Top-10 <br></br>
     * Top-n filters only show the Top n values or percent in the column
     *
     * @return
     */
    val isTop10: Boolean
        get() = af?.isTop10 ?: false

    /**
     * returns true if this AutoFitler is set to filter all blank rows
     *
     * @return true if filter blanks, false otherwise
     */
    val isFilterBlanks: Boolean
        get() = af?.isFilterBlanks ?: false

    /**
     * returns true if this AutoFitler is set to filter all non-blank rows
     *
     * @return true if filter non-blanks, false otherwise
     */
    val isFilterNonBlanks: Boolean
        get() = af?.isFilterNonBlanks ?: false

    init {
        this.af = af
    }

    /**
     * returns the string representation of this AutoFilter
     *
     * @return the string representation of this AutoFilter
     */
    override fun toString(): String {
        return if (this.af != null) af.toString() else "No AutoFilter"
    }

    /**
     * Sets the custom comparison of this AutoFilter via a String operator and an
     * Object value <br></br>
     * Only those rows that meet the equation: (column value) OP value will be shown
     * <br></br>
     * e.g show all rows where column value >= 2.99
     *
     *
     * Object value can be of type
     *
     *
     * String <br></br>
     * Boolean <br></br>
     * Error <br></br>
     * a Number type object
     *
     *
     * String operator may be one of: "=", ">", ">=", "<>", "<", "<="
     *
     * @param Object val - value to set
     * @param String op - operator
     * @see setVal2
     */
    fun setVal(`val`: Any, op: String) {
        af?.setVal(`val`, op)
    }

    /**
     * Sets the custom comparison of the second condition of this AutoFilter via a
     * String operator and an Object value <br></br>
     * This method sets the second condition of a two-condition filter
     *
     *
     * Only those rows that meet the equation: <br></br>
     * first condition AND/OR (column value) OP value will be shown <br></br>
     * e.g show all rows where (column value) <= 1.99 AND (column value) >= 2.99
     *
     *
     * Object value can be of type
     *
     *
     * String <br></br>
     * Boolean <br></br>
     * Error <br></br>
     * a Number type object
     *
     *
     * String operator may be one of: "=", ">", ">=", "<>", "<", "<="
     *
     * @param Object  val - value to set
     * @param String  op - operator
     * @param boolean AND - true if two conditions should be AND'ed, false if OR'd
     * @see setVal2
     */
    fun setVal2(`val`: Any, op: String, AND: Boolean) {
        af?.setVal2(`val`, op, AND)
    }

    /**
     * Sets this AutoFilter to be a Top-n or Bottom-n type of filter <br></br>
     * Top-n filters only show the Top n values or percent in the column <br></br>
     * Bottom-n filters only show the bottom n values or percent in the column <br></br>
     * n can be from 1-500, or 0 to turn off Top 10 filtering
     *
     * @param int     n - 0-500
     * @param boolean percent - true if show Top-n percent; false to show Top-n items
     * @param boolean top10 - true if show Top-n (items or percent), false to show
     * Bottom-n (items or percent)
     */
    fun setTop10(n: Int, percent: Boolean, top10: Boolean) {
        af?.setTop10(n, percent, top10)
    }

    /**
     * sets this AutoFilter to filter all blank rows
     */
    fun setFilterBlanks() {
        af?.setFilterBlanks()
    }

    /**
     * sets this AutoFilter to filter all non-blank rows
     */
    fun setFilterNonBlanks() {
        af?.setFilterNonBlanks()
    }
}
