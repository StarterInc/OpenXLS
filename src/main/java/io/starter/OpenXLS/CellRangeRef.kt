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

import io.starter.formats.XLS.WorkSheetNotFoundException
import io.starter.formats.XLS.XLSRecord

/**
 * Represents a reference to a 3D range of cells.
 *
 *
 * This class is not currently part of the public API.
 *
 *
 * No input validation whatsoever is performed. Any combination of values may be
 * set, whether it makes any sense or not.
 *
 * [Starter Inc.](http://starter.io)
 *
 * @see DataBoundCellRange
 *
 * @see XLSRecord
 */
class CellRangeRef : Cloneable {
    /**
     * Returns the lowest-indexed column in this range.
     *
     * @return the column index or null if this is a row range
     */
    /**
     * Sets the first column in this range.
     *
     * @param value the column index to set
     */
    var firstColumn: Int = 0
    /**
     * Returns the lowest-indexed row in this range.
     *
     * @return the row index or null if this is a column range
     */
    /**
     * Sets the first row in this range.
     *
     * @param value the row index to set
     */
    var firstRow: Int = 0
    /**
     * Returns the highest-indexed column in this range.
     *
     * @return the column index or null if this is a row range
     */
    /**
     * Sets the last column in this range.
     *
     * @param value the column index to set
     */
    var lastColumn: Int = 0
    /**
     * Returns the highest-indexed row in this range.
     *
     * @return the row index or null if this is a column range
     */
    /**
     * Sets the last row in this range.
     *
     * @param value the row index to set
     */
    var lastRow: Int = 0
    private var first_sheet_name: String? = null
    private var last_sheet_name: String? = null
    private var first_sheet: WorkSheetHandle? = null
    private var last_sheet: WorkSheetHandle? = null

    /**
     * Returns the name of the first sheet in this range.
     *
     * @return the name of the sheet or null if this range is not qualified with a
     * sheet
     */
    val firstSheetName: String?
        get() = if (first_sheet != null)
            first_sheet!!.sheetName
        else
            first_sheet_name

    /**
     * Returns the first sheet in this range.
     *
     * @return the WorkSheetHandle or null if this range is not qualified with a
     * sheet or the sheet names have not been resolved
     */
    /**
     * Sets the first sheet in this range.
     */
    var firstSheet: WorkSheetHandle?
        get() = first_sheet
        set(sheet) {
            first_sheet_name = null
            first_sheet = sheet
        }

    /**
     * Returns the name of the last sheet in this range.
     *
     * @return the name of the sheet or null if this range is not qualified with a
     * sheet
     */
    val lastSheetName: String?
        get() = if (last_sheet != null)
            last_sheet!!.sheetName
        else
            last_sheet_name

    /**
     * Returns the last sheet in this range.
     *
     * @return the WorkSheetHandle or null if this range is not qualified with a
     * sheet or the sheet names have not been resolved
     */
    /**
     * Sets the last sheet in this range.
     */
    var lastSheet: WorkSheetHandle?
        get() = last_sheet
        set(sheet) {
            last_sheet_name = null
            last_sheet = sheet
        }

    /**
     * Determines whether this range spans multiple sheets.
     */
    val isMultiSheet: Boolean
        get() = first_sheet != null && last_sheet != null && first_sheet !== last_sheet || first_sheet_name != null && last_sheet_name != null && first_sheet_name !== last_sheet_name

    /**
     * Private nullary constructor for use by static pseudo-constructors.
     */
    private constructor() {}

    constructor(first_row: Int, first_col: Int, last_row: Int, last_col: Int) {
        this.firstRow = first_row
        this.firstColumn = first_col
        this.lastRow = last_row
        this.lastColumn = last_col
    }

    /**
     * return the number of cells in this rangeref
     *
     * @return number of cells in ref
     */
    fun numCells(): Int {
        var ret = -1
        var numrows = this.lastRow - this.firstRow
        numrows++
        var numcols = this.lastColumn - this.firstColumn
        numcols++
        ret = numrows * numcols
        return ret
    }

    constructor(first_row: Int, first_col: Int, last_row: Int, last_col: Int, first_sheet: String,
                last_sheet: String) : this(first_row, first_col, last_row, last_col) {
        this.first_sheet_name = first_sheet
        this.last_sheet_name = last_sheet
    }

    constructor(first_row: Int, first_col: Int, last_row: Int, last_col: Int, first_sheet: WorkSheetHandle,
                last_sheet: WorkSheetHandle) : this(first_row, first_col, last_row, last_col) {
        this.first_sheet = first_sheet
        this.last_sheet = last_sheet
    }

    /**
     * Resolves sheet names into sheet handles against the given book.
     *
     * @param book the book against which the sheet names should be resolved
     * @throws WorkSheetNotFoundException if either of the sheets does not exist in the given book
     */
    @Throws(WorkSheetNotFoundException::class)
    fun resolve(book: WorkBookHandle) {
        if (first_sheet_name != null)
            first_sheet = book.getWorkSheet(first_sheet_name)
        if (last_sheet_name != null)
            last_sheet = book.getWorkSheet(last_sheet_name)
    }

    /**
     * Determines whether this range is qualified with a sheet.
     */
    fun hasSheet(): Boolean {
        return first_sheet != null || first_sheet_name != null
    }

    /**
     * Returns whether this range entirely contains the given range. This ignores
     * the sheets, if any, and compares only the cell ranges.
     */
    operator fun contains(range: CellRangeRef): Boolean {
        return (this.firstRow <= range.firstRow && this.lastRow >= range.lastRow && this.firstColumn <= range.firstColumn
                && this.lastColumn >= range.lastColumn)
    }

    /**
     * Compares this range to the specified object. The result is `true`
     * if and only if the argument is not `null` and is a
     * `CellRangeRef` object that represents the same range as this
     * object.
     */
    override fun equals(other: Any?): Boolean {
        // if it's null or not a CellRangeRef it can't be equal
        return if (other == null || other !is CellRangeRef) false else this.toString() == other.toString()

    }

    /**
     * Creates and returns a copy of this range.
     */
    public override fun clone(): Any {
        try {
            return super.clone()
        } catch (e: CloneNotSupportedException) {
            // This can't happen (we're Cloneable) but we have to catch it
            throw Error("Object.clone() threw CNSE but we're Cloneable")
        }

    }

    /**
     * Gets this range in A1 notation.
     */
    override fun toString(): String {
        val sheet1 = firstSheetName
        val sheet2 = lastSheetName
        return (if (sheet1 != null) sheet1 + (if (sheet2 != null && sheet2 !== sheet1) ":$sheet2" else "") + "!" else "") + ExcelTools.formatRange(intArrayOf(firstColumn, firstRow, lastColumn, lastRow))
    }

    companion object {

        /**
         * Parses a range in A1 notation and returns the equivalent CellRangeRef.
         */
        fun fromA1(reference: String): CellRangeRef {
            val ret = CellRangeRef()
            var range: String?

            run {
                val parts = ExcelTools.stripSheetNameFromRange(reference)
                ret.first_sheet_name = parts[0]
                range = parts[1]
                ret.last_sheet_name = parts[2]
            }

            if (range == null)
                throw IllegalArgumentException("missing range component")

            run {
                val parts = ExcelTools.getRangeRowCol(range!!)
                ret.firstRow = parts[0]
                ret.firstColumn = parts[1]
                ret.lastRow = parts[2]
                ret.lastColumn = parts[3]
            }

            return ret
        }

        /**
         * Convenience method combining [.fromA1] and
         * [.resolve].
         */
        @Throws(WorkSheetNotFoundException::class)
        fun fromA1(reference: String, book: WorkBookHandle): CellRangeRef {
            val ret = fromA1(reference)
            ret.resolve(book)
            return ret
        }
    }
}
