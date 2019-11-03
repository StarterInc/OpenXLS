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

/**
 * A lightweight subset of Cell methods allowing for low memory overhead
 * streaming implementations
 *
 * @author John McMahon
 */
interface Cell {

    val isDate: Boolean

    val cellType: Int

    /**
     * Returns the Formatting record ID (FormatId) for this Cell <br></br>
     * This can be used with 'setFormatId(int i)' to copy the formatting from one
     * Cell to another (e.g. a template cell to a new cell)
     *
     * @return int the FormatId for this Cell
     */
    val formatId: Int

    /**
     * Returns the value of this Cell in the native underlying data type.
     *
     *
     * Formula cells will return the calculated value of the formula in the
     * calculated data type.
     *
     *
     * Use 'getStringVal()' to return a String regardless of underlying value type.
     *
     * @return Object value for this Cell
     */
    val `val`: Any

    /**
     * Returns the value of the Cell as a String with formatting pattern applied..
     * <br></br>
     * see: [http://java.sun.com/docs/books/tutorial/i18n/format/decimalFormat.html](tutorial)
     * <br></br>
     * boolean Cell types will return "true" or "false" <br></br>
     * Negative numbers that are formatted in excel to show as red values rather
     * than using a "-" will return with a minus symbol.
     *
     * @return String the formatted value of the Cell
     */
    val formattedStringVal: String

    /**
     * Returns the column number of this Cell.
     *
     * @return int the Column Number of the Cell
     */
    val colNum: Int

    /**
     * Returns the row number of this Cell.
     *
     *
     * NOTE: This is the 1-based row number such as you will see in a spreadsheet
     * UI.
     *
     *
     * ie: A1 = row 1
     *
     * @return int the ONE-based Row Number of the Cell
     */
    val rowNum: Int

    /**
     * Returns the Address of this Cell as a String.
     *
     * @return String the address of this Cell in the WorkSheet
     */
    val cellAddress: String

    /**
     * Returns the name of this Cell's WorkSheet as a String.
     *
     * @return String the name this Cell's WorkSheet
     */
    val workSheetName: String

    companion object {

        /**
         * Cell types
         */
        val TYPE_BLANK = -1
        val TYPE_STRING = 0
        val TYPE_FP = 1
        val TYPE_INT = 2
        val TYPE_FORMULA = 3
        val TYPE_BOOLEAN = 4
        val TYPE_DOUBLE = 5
    }

}
