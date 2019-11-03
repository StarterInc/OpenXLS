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

import io.starter.formats.XLS.*

/**
 * An interface representing an OpenXLS WorkSheet
 *
 * @see WorkSheetHandle
 */
interface WorkSheet : Handle {

    val workBook: WorkBook

    val mergedCellsRec: Mergedcells

    /**
     * Get the first row on the Worksheet
     *
     * @return the Minimum Row Number on the Worksheet
     */
    val firstRow: Int

    /**
     * Get the first column on the Worksheet
     *
     * @return the Minimum Column Number on the Worksheet
     */
    val firstCol: Int

    /**
     * Get the last row on the Worksheet
     *
     * @return the Maximum Row Number on the Worksheet
     */
    val lastRow: Int

    /**
     * Get the last column on the Worksheet
     *
     * @return the Maximum Column Number on the Worksheet
     */
    val lastCol: Int

    /**
     * get whether this sheet is selected upon opening the file.
     *
     * @return boolean b selected state
     */
    /**
     * set whether this sheet is selected upon
     * opening the file.
     *
     * @param boolean b selected value
     */
    var selected: Boolean

    /**
     * get whether this sheet is hidden from the user opening the file.
     *
     * @return boolean b hidden state
     */
    /**
     * set whether this sheet is hidden from the user opening the file.
     *
     *
     * if the sheet is selected, the API will set the first
     * visible sheet to selected as you cannot have your selected sheet
     * be hidden.
     *
     *
     * to override this behavior, set your desired sheet to selected
     * after calling this method.
     *
     * @param boolean b hidden state
     */
    var hidden: Boolean

    /**
     * get the tab display order of this Worksheet
     *
     *
     * this is a zero based index with zero representing
     * the left-most WorkSheet tab.
     *
     * @return int idx the index of the sheet tab
     */
    /**
     * set the tab display order of this Worksheet
     *
     *
     * this is a zero based index with zero representing
     * the left-most WorkSheet tab.
     *
     * @param int idx the new index of the sheet tab
     */
    var tabIndex: Int

    /**
     * returns all of the Columns in this WorkSheet
     *
     * @return ColHandle[] Columns
     */
    val columns: Array<ColHandle>

    /**
     * returns a Vector of Column names
     *
     * @return CompatibleVector column names
     */
    val colNames: List<*>

    /**
     * returns a Vector of Row numbers
     *
     * @return CompatibleVector row numbers
     */
    val rowNums: List<*>

    /**
     * get an array of all RowHandles for
     * this WorkSheet
     *
     * @return RowHandle[] all Rows on this WorkSheet
     */
    val rows: Array<RowHandle>

    /**
     * Returns the number of rows in this WorkSheet
     *
     * @return int Number of Rows on this WorkSheet
     */
    val numRows: Int

    /**
     * Returns the number of Columns in this WorkSheet
     *
     * @return int Number of Cols on this WorkSheet
     */
    val numCols: Int

    /**
     * Returns the low-level OpenXLS Sheet.
     *
     * @return io.starter.formats.XLS.Sheet
     */
    val sheet: Sheet

    /**
     * Returns the name of the Sheet.
     *
     * @return String Sheet Name
     */
    /**
     * Set the name of the Worksheet.  This method will change
     * the name on the Worksheet's tab as displayed in the
     * WorkBook, as well as all programmatic and internal
     * references to the name.
     *
     *
     * This change takes effect immediately, so all attempts
     * to reference the Worksheet by its previous name will
     * fail.
     *
     * @param String the new name for the Worksheet
     */
    var sheetName: String

    /**
     * Returns the Serialized bytes for this WorkSheet.
     *
     *
     * The output of this method can be used to insert
     * a copy of this WorkSheet into another WorkBook
     * using the WorkBook.addWorkSheet(byte[] serialsheet, String NewSheetName)
     * method.
     *
     * @return byte[] the WorkSheet's Serialized bytes
     * @see WorkBook.addWorkSheet
     */
    val serialBytes: ByteArray

    /**
     * Returns all CellHandles defined on this WorkSheet.
     *
     * @return Cell[] - the array of Cells in the Sheet
     */
    val cells: Array<CellHandle>

    /**
     * Get the text for the Footer printed at the bottom
     * of the Worksheet
     *
     * @return String footer text
     */
    /**
     * Set the text for the Footer printed at the bottom
     * of the Worksheet
     *
     * @param String footer text
     */
    var footerText: String

    /**
     * Get the text for the Header printed at the top
     * of the Worksheet
     *
     * @return String header text
     */
    /**
     * Set the text for the Header printed at the top
     * of the Worksheet
     *
     * @param String header text
     */
    var headerText: String

    fun addChart(serialchart: ByteArray, name: String)

    /* set protection on sheet

        @param boolean whether to protect the sheet
    */
    fun setProtected(b: Boolean)

    /**
     * set whether this sheet is VERY hidden opening the file.
     *
     *
     * VERY hidden means users will not be able to unhide
     * the sheet without using VB code.
     *
     * @param boolean b hidden state
     */
    fun setVeryHidden(b: Boolean)

    /**
     * set this WorkSheet as the first visible tab on the left
     */
    fun setFirstVisibleTab()

    /**
     * returns the Column at the index position
     *
     * @return ColHandle the Column
     */
    @Throws(ColumnNotFoundException::class)
    fun getCol(clnum: Int): ColHandle

    /**
     * returns the Column at the named position
     *
     * @return ColHandle the Column
     */
    @Throws(ColumnNotFoundException::class)
    fun getCol(name: String): ColHandle

    /**
     * get an a RowHandle for
     * this WorkSheet by number
     *
     * @param int row number to return
     * @return RowHandle a Row on this WorkSheet
     */
    @Throws(RowNotFoundException::class)
    fun getRow(t: Int): RowHandle

    /**
     * Remove a Cell from this WorkSheet.
     *
     * @param CellHandle to remove
     */
    fun removeCell(celldel: CellHandle)

    /**
     * Remove a Cell from this WorkSheet.
     *
     * @param String celladdr - the Address of the Cell to remove
     */
    fun removeCell(celladdr: String)

    /**
     * Remove a Row and all associated Cells from
     * this WorkSheet.
     *
     * @param int rownum - the number of the row to remove
     */
    @Throws(RowNotFoundException::class)
    fun removeRow(rownum: Int)

    /**
     * Remove a Row and all associated Cells from
     * this WorkSheet. Optionally shift all rows below
     * target row up one.
     *
     * @param int     rownum - the number of the row to remove
     * @param boolean shiftrows - true will shift all lower rows up one.
     */
    @Throws(RowNotFoundException::class)
    fun removeRow(rownum: Int, shiftrows: Boolean)

    /**
     * Remove a Column and all associated Cells from
     * this WorkSheet.
     *
     * @param String colstr - the name of the column to remove
     */
    @Throws(ColumnNotFoundException::class)
    fun removeCol(colstr: String)

    /**
     * Remove a Column and all associated Cells from
     * this WorkSheet.
     *
     *
     * Optionally shift all cols to the right
     * of this column to the left by one.
     *
     * @param String  colstr - the name of the column to remove
     * @param boolean shiftcols - true will shift following cols
     */
    @Throws(ColumnNotFoundException::class)
    fun removeCol(colstr: String, shiftcols: Boolean)

    /**
     * Set the Object value of the Cell at the given address.
     *
     * @param String Cell Address (ie: "D14")
     * @param Object new Cell Object value
     * @throws CellNotFoundException is thrown if there is
     * no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, `val`: Any)

    /**
     * Set the double value of the Cell at the given address
     *
     * @param String  Cell Address (ie: "D14")
     * @param double  new Cell double value
     * @param address
     * @param d
     * @throws io.starter.formats.XLS.CellNotFoundException is
     * thrown if there is no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, d: Double)

    /**
     * Set the String value of the Cell at the given address
     *
     * @param String Cell Address (ie: "D14")
     * @param String new Cell String value
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, s: String)

    /**
     * Set the int value of the Cell at the given address
     *
     * @param String Cell Address (ie: "D14")
     * @param int    new Cell int value
     * @throws io.starter.formats.XLS.CellNotFoundException is thrown if there is no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, i: Int)

    /**
     * Get the Object value of a Cell.
     *
     *
     * Numeric Cell values will return as type Long, Integer, or Double.
     * String Cell values will return as type String.
     *
     * @return the value of the Cell as an Object.
     * @throws CellNotFoundException is thrown if there is
     * no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class)
    fun getVal(address: String): Any

    /**
     * Insert a blank row into the worksheet.
     * Shift all rows below the cell down one.
     *
     *
     * Adding new cells to non-existent rows will
     * automatically create new rows in the file,
     * This method is only necessary to "move" existing cells
     * by inserting empty rows.
     *
     * @param rownum the rownumber to insert
     */
    fun insertRow(rownum: Int)

    /**
     * Insert a blank row into the worksheet.
     * Shift all rows below the cell down one.
     *
     *
     * Adding new cells to non-existent rows will
     * automatically create new rows in the file,
     *
     *
     * After calling this method, setVal() can be used on the
     * newly created cells to update with new values.
     *
     * @param rownum  the rownumber to insert
     * @param whether to shift down existing Cells
     */
    fun insertRow(rownum: Int, shiftrows: Boolean)

    /**
     * Insert a blank row into the worksheet.
     * Shift all rows below the cell down one.
     *
     *
     * Adding new cells to non-existent rows will
     * automatically create new rows in the file,
     *
     *
     * After calling this method, setVal() can be used on the
     * newly created cells to update with new values.
     *
     * @param rownum  the rownumber to insert
     * @param whether to shift down existing Cells
     */
    fun insertRow(
            rownum: Int,
            copyrow: Int,
            flag: Int,
            shiftrows: Boolean)

    /**
     * Insert a blank column into the worksheet.
     * Shift all columns to the right of the cell over one.
     *
     *
     * Adding new cells to non-existent columns will
     * automatically create new Columns in the file,
     * This method is only necessary to "move" existing cells
     * by inserting empty columns.
     *
     * @param colstr the Column string to insert
     */
    fun insertCol(colnum: String)

    /**
     * When adding a new Cell to the sheet, OpenXLS can
     * automatically copy the formatting from the Cell directly
     * above the inserted Cell.
     *
     *
     * ie: if set to true, newly added Cell D19 would take its formatting
     * from Cell D18.
     *
     *
     * Default is false
     *
     *
     *
     *
     *
     *
     * boolean copy the formats from the prior Cell
     */
    fun setCopyFormatsFromPriorWhenAdding(f: Boolean)

    /**
     * Add a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on
     * type-compatibility of the value.
     *
     *
     * In other words, if the Object cannot be converted
     * safely to a Numeric Object type, then it is treated
     * as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     * @param obj the value of the new Cell
     * @param int row the row of the new Cell
     * @param int col the column of the new Cell
     */
    fun add(obj: Any, row: Int, col: Int): CellHandle

    /**
     * Add a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on
     * type-compatibility of the value.
     *
     *
     * In other words, if the Object cannot be converted
     * safely to a Numeric Object type, then it is treated
     * as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     * @param obj     the value of the new Cell
     * @param address the address of the new Cell
     */
    fun add(obj: Any, address: String): CellHandle

    /**
     * Add a java.sql.Date Cell to a WorkSheet.
     *
     *
     * You must specify a formatting pattern for the
     * new date, or null for the default ("m/d/yy h:mm".)
     *
     *
     * valid date format patterns
     * "m/d/y"
     * "d-mmm-yy"
     * "d-mmm"
     * "mmm-yy"
     * "h:mm AM/PM"
     * "h:mm:ss AM/PM"
     * "h:mm"
     * "h:mm:ss"
     * "m/d/yy h:mm"
     * "mm:ss"
     * "[h]:mm:ss"
     * "mm:ss.0"
     *
     * @param dt         the value of the new java.sql.Date Cell
     * @param address    the address of the new java.sql.Date Cell
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    fun add(
            dt: java.util.Date,
            address: String,
            fmt: String): CellHandle

    /**
     * Remove this WorkSheet from the WorkBook
     */
    fun remove()

    /**
     * Returns a FormulaHandle for working with
     * the ranges of a formula on a WorkSheet.
     *
     * @param addr the address of the Cell
     * @throws FormulaNotFoundException is thrown if there is no existing formula at the specified address.
     */
    @Throws(FormulaNotFoundException::class, CellNotFoundException::class)
    fun getFormula(addr: String): FormulaHandle

    /**
     * Returns a CellHandle for working with
     * the value of a Cell on a WorkSheet.
     *
     * @param addr the address of the Cell
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class)
    fun getCell(addr: String): CellHandle

    /**
     * Returns a CellHandle for working with
     * the value of a Cell on a WorkSheet.
     *
     * @param int Row the integer row of the Cell
     * @param int Col the integer col of the Cell
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
     */
    @Throws(CellNotFoundException::class)
    fun getCell(row: Int, col: Int): CellHandle

    /**
     * Move a cell on this WorkSheet.
     *
     * @param CellHandle c - the cell to be moved
     * @param String     celladdr - the destination address of the cell
     */
    @Throws(CellPositionConflictException::class)
    fun moveCell(c: CellHandle, addr: String)

    /**
     * Returns the name of this Sheet.
     *
     * @see java.lang.Object.toString
     */
    override fun toString(): String

    companion object {

        val ROW_KEEP = 0

        val ROW_INSERT = 1
    }
}