/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.OpenXLS

import io.starter.formats.XLS.*

import java.io.OutputStream

/**
 * An interface representing an OpenXLS WorkBook.
 */
interface WorkBook : Handle, Document {

    /**
     * Explicit calcing of formulas
     *
     * @param mode
     */
    var formulaCalculationMode: Int

    /**
     * The Session for the WorkBook instance
     *
     * @return public BookSession getSession();
     */

    val colorTable: Array<java.awt.Color>

    /**
     * Returns all of the Cells contained in the WorkBook
     */
    val cells: Array<CellHandle>

    /**
     * Returns an Array of all the FormatHandles present in the workbook
     *
     * @return all existing FormatHandles in the workbook
     */
    val formats: Array<FormatHandle>

    /**
     * Gets the date format used by this book.
     */
    val dateFormat: DateConverter.DateFormat

    /**
     * get an array of handles to all PivotTables in the WorkBook
     *
     * @return PivotTable[] all of the WorkBooks PivotTables
     */
    val pivotTables: Array<PivotTableHandle>

    /**
     * Returns all Chart Handles contained in the WorkBook
     *
     * @return ChartHandle[] an array of all Charts in the WorkBook
     */
    val charts: Array<ChartHandle>

    /**
     * Returns all Named Range Handles
     *
     * @return NameHandle[] all of the Named ranges in the WorkBook
     */
    val namedRanges: Array<NameHandle>

    /**
     * Returns the name of this WorkBook
     *
     * @return String name of WorkBook
     */
    /**
     * Sets the internal name of this WorkBookHandle.
     *
     *
     * Overrides the default for 'getName()' which returns
     * the file name source of this WorkBook by default.
     *
     * @param WorkBook Name
     */
    override var name: String

    /**
     * Returns the number of Cells in this WorkBook
     *
     * @return int number of Cells
     */
    val numCells: Int

    /**
     * Returns a byte Array containing the
     * valid file containing this WorkBook
     * and associated Storages (such as VB files
     * and PivotTables.)
     *
     *
     * This is the actual file data and that can be
     * read from and written to FileOutputStreams and
     * ServletOutputStreams.
     *
     * @return byte[] the XLS File's bytes
     */
    val bytes: ByteArray

    /** Initialize a Vector of the CellRanges existing in this WorkBook
     * specifically the Ranges referenced in Formulas, Charts, and
     * Named Ranges.
     *
     * This is necessary to allow for automatic updating of references
     * when adding/removing/moving Cells within these ranges, as well
     * as shifting references to Cells in Formulas when Formula records
     * are moved.
     *
     */
    // public abstract void initCellRanges(boolean
    // createblanks);

    /**
     * Returns an array of handles to all
     * of the WorkSheets in the Workbook.
     *
     * @return WorkSheetHandle[] Array of all WorkSheets in WorkBook
     */
    val workSheets: Array<WorkSheetHandle>

    /**
     * Returns a low-level WorkBook.
     *
     *
     * NOTE: The WorkBook class is NOT a part of the
     * published API.  Any of the methods and/or
     * variables on a WorkBook object are subject
     * to change without notice in new versions of OpenXLS.
     */
    val workBook: io.starter.formats.XLS.WorkBook

    /**
     * Returns a WorkBookHandle containing an empty
     * version of this WorkBook.
     *
     *
     * Use in conjunction with  addSheetFromWorkBook() to create
     * new output WorkBooks containing various sheets from a master
     * template.
     *
     *
     * ie:
     * WorkBookHandle emptytemplate = this.getNoSheetWorkBook();
     * emptytemplate.addSheetFromWorkBook(this, "Sheet1", "TargetSheet");
     *
     * @return WorkBookHandle - the empty WorkBookHandle duplicate
     * @see addSheetFromWorkBook
     */
    val noSheetWorkBook: WorkBookHandle

    /**
     * get a non-Excel property
     *
     * @return Returns the properties.
     */
    override fun getProperty(name: String): Any

    /**
     * add non-Excel property
     *
     * @param properties The properties to set.
     */
    override fun addProperty(name: String, `val`: Any)

    /**
     * Returns the Cell at the specified Location
     *
     * @param address
     * @return
     */
    @Throws(CellNotFoundException::class, WorkSheetNotFoundException::class)
    fun getCell(address: String): CellHandle

    /**
     * get a handle to a PivotTable in the WorkBook
     *
     * @param String name of the PivotTable
     * @return PivotTable the PivotTable
     */
    @Throws(PivotTableNotFoundException::class)
    fun getPivotTable(ptname: String): PivotTableHandle

    /**
     * set the workbook to protected mode
     *
     *
     * Note: the password cannot be decrypted or changed
     * in Excel -- protection can only be set/removed using
     * OpenXLS
     *
     * @param int Default Column width
     */
    fun setProtected(b: Boolean)

    /**
     * set Default row height
     *
     *
     * Note: only affects undefined Rows containing Cells
     *
     * @param int Default Row Height
     */
    fun setDefaultRowHeight(t: Int)

    /**
     * set Default col width
     *
     *
     * Note: only affects undefined Columns containing Cells
     *
     * @param int Default Column width
     */
    fun setDefaultColWidth(t: Int)

    /**
     * Set the Debugging level.  Higher values output more
     * debugging info during execution.
     *
     * @parameter int Debug level.  higher=more verbose
     */
    override fun setDebugLevel(l: Int)

    /**
     * Returns a Named Range Handle
     *
     * @return NameHandle a Named range in the WorkBook
     */
    @Throws(CellNotFoundException::class)
    fun getNamedRange(rangename: String): NameHandle

    /**
     * Returns a Chart Handle
     *
     * @return ChartHandle a Chart in the WorkBook
     */
    @Throws(ChartNotFoundException::class)
    fun getChart(chartname: String): ChartHandle

    /**
     * returns the handle to a WorkSheet by name.
     *
     * @param index of worksheet (ie: 0)
     * @return WorkSheetHandle the WorkSheet
     * @throws WorkSheetNotFoundException if the specified WorkSheet is
     * not found in the WorkBook.
     */
    @Throws(WorkSheetNotFoundException::class)
    fun getWorkSheet(i: Int): WorkSheetHandle

    /**
     * returns the handle to a WorkSheet by name.
     *
     * @param String name of worksheet (ie: "Sheet1")
     * @return WorkSheetHandle the WorkSheet
     * @throws WorkSheetNotFoundException if the specified WorkSheet is
     * not found in the WorkBook.
     */
    @Throws(WorkSheetNotFoundException::class)
    fun getWorkSheet(handstr: String): WorkSheetHandle

    /**
     * Clears all values in a template WorkBook.
     *
     *
     * Use this method to 'reset' the values of your
     * WorkBook in memory to defaults.
     *
     *
     * For example, if you load a Servlet with a
     * single WorkBookHandle instance, then modify
     * values and stream to a Client system, yo
     * should call 'clearAll()' when the request
     * is completed to remove the modified values
     * and set them back to a default.
     */
    override fun reset()

    /**
     * Set Encoding mode of new Strings added to file.
     *
     *
     * OpenXLS has 3 modes for handling the internal encoding of
     * String data that is added to the file.
     *
     *
     * OpenXLS can save space in the file if it knows that all characters
     * in your String data can be represented with a single byte (Compressed.)
     *
     *
     * If your String contains characters which need 2 bytes to represent (such
     * as Eastern-language characters) then it needs to be stored in an uncompressed
     * Unicode format.
     *
     *
     * OpenXLS can either automatically detect the mode for each String, or you
     * can set it explicitly.  The auto mode is the most flexible but requires processing
     * overhead.
     *
     *
     * Default mode is WorkBookHandle.STRING_ENCODING_AUTO.
     *
     *
     * Valid Modes Are:
     *
     *
     * WorkBookHandle.STRING_ENCODING_AUTO          Use if you are adding mixed Unicode and non-unicode
     * Strings and can accept the performance hit
     * -slowest String adds
     * -optimal file size for mixed Strings
     * WorkBookHandle.STRING_ENCODING_UNICODE       Use if all of your new Strings are Unicode - faster than AUTO
     * -faster than AUTO
     * -largest file size
     * WorkBookHandle.STRING_ENCODING_COMPRESSED    Use if all of your new Strings are non-Unicode and can have high-bytes compressed
     * -faster than AUTO
     * -smallest file size
     *
     * @param int String Encoding Mode
     */
    fun setStringEncodingMode(mode: Int)

    /**
     * Set Duplicate String Handling Mode.
     *
     *
     * The Duplicate String Mode determines the behavior of
     * the String table when inserting new Strings.
     *
     *
     * The String table shares a single entry for multiple
     * Cells containing the same string.  When multiple Cells
     * have the same value, they share the same underlying string.
     *
     *
     * Changing the value of any one of the Cells will change
     * the value for any Cells sharing that reference.
     *
     *
     * For this reason, you need to determine
     * the handling of new strings added to the sheet that
     * are duplicates of strings already in the table.
     *
     *
     * If you will be changing the values of these
     * new Cells, you will need to set the Duplicate
     * String Mode to ALLOWDUPES.  If the string table
     * encounters a duplicate entry being added, it
     * will insert a duplicate that can then be subsequently
     * changed without affecting the other duplicate Cells.
     *
     *
     * Valid Modes Are:
     *
     *
     * WorkBookHandle.ALLOWDUPES - faster, smaller file sizes, dupe Cells share changes
     * WorkBookHandle.SHAREDUPES - slower inserts, changing Cells has no effect on dupe Cells
     *
     * @param int Duplicate String Handling Mode
     */
    fun setDupeStringMode(mode: Int)

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chartname
     * @param sheetname
     */
    @Throws(ChartNotFoundException::class, WorkSheetNotFoundException::class)
    fun copyChartToSheet(chartname: String, sheetname: String)

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chart
     * @param sheet
     */
    @Throws(ChartNotFoundException::class, WorkSheetNotFoundException::class)
    fun copyChartToSheet(chart: ChartHandle, sheet: WorkSheetHandle)

    /**
     * Copy (duplicate) a worksheet in the workbook and add it to the end of the workbook with a new name
     *
     * @param String the Name of the source worksheet;
     * @param String the Name of the new (destination) worksheet;
     * @return the new WorkSheetHandle
     */
    @Throws(WorkSheetNotFoundException::class)
    fun copyWorkSheet(SourceSheetName: String, NewSheetName: String): WorkSheetHandle

    /**
     * Iterate through the formulas in this WorkBook and call the
     * calculate method on each.
     *
     *
     * May be more expensive than calling update on individual FormulaHandles
     * depending on extent of data changes to your WorkBook, or calling update
     * on only the 'top-level' formula in a calculation.  When a formula references
     * a Cell containing another formula, it will recursively calculate until it
     * reaches non-formula Cells.  Thus, calling this method may calculate formula
     * Cells in a heirarchy more than once.
     *
     * @throws FunctionNotSupportedException
     */
    @Throws(FunctionNotSupportedException::class)
    fun calculateFormulas()

    /**
     * Removes all of the WorkSheets from this WorkBook.
     *
     *
     * Bytes streamed from this WorkBook will create invalid
     * Spreadsheet files unless a WorkSheet(s) are added to it.
     */
    fun removeAllWorkSheets()

    /**
     * Inserts a worksheet from a Source WorkBook.
     *
     * @param sourceBook      - the WorkBook containing the sheet to copy
     * @param sourceSheetName - the name of the sheet to copy
     * @param destSheetName   - the name of the new sheet in this workbook
     * @throws WorkSheetNotFoundException
     */
    @Throws(WorkSheetNotFoundException::class)
    fun addSheetFromWorkBook(sourceBook: WorkBookHandle, sourceSheetName: String, destSheetName: String): Boolean

    /**
     * Inserts a new worksheet and places it at the end of the workbook
     *
     * @param WorkSheetHandle the source WorkSheetHandle;
     * @param String          the Name of the new (destination) worksheet;
     */
    fun addWorkSheet(sht: WorkSheetHandle, NewSheetName: String): WorkSheetHandle

    /**
     * Creates a new worksheet and places it at the end of the workbook
     *
     * @param String the Name of the newly created worksheet
     * @return the new WorkSheetHandle
     */
    fun createWorkSheet(name: String): WorkSheetHandle

    /**
     * Returns the name of this Sheet.
     *
     * @see java.lang.Object.toString
     */
    override fun toString(): String

    fun writeBytes(bbout: OutputStream): StringBuffer

    companion object {

        val CALCULATE_ALWAYS = XLSConstants.CALCULATE_ALWAYS
        val CALCULATE_EXPLICIT = XLSConstants.CALCULATE_EXPLICIT
        val CALCULATE_AUTO = XLSConstants.CALCULATE_AUTO
        val CALC_MODE_PROP = XLSConstants.CALC_MODE_PROP
        val REFTRACK_PROP = XLSConstants.REFTRACK_PROP
        val USETEMPFILE_PROP = XLSConstants.USETEMPFILE_PROP
        val DEFAULTENCODING = XLSConstants.DEFAULTENCODING
        val UNICODEENCODING = XLSConstants.UNICODEENCODING
        val VALIDATEWORKBOOK = XLSConstants.VALIDATEWORKBOOK

        // public static final int FORMULA_CALC_AUTO = 0;

        val STRING_ENCODING_AUTO = XLSConstants.STRING_ENCODING_AUTO
        val STRING_ENCODING_UNICODE = XLSConstants.STRING_ENCODING_UNICODE
        val STRING_ENCODING_COMPRESSED = XLSConstants.STRING_ENCODING_COMPRESSED
        val ALLOWDUPES = XLSConstants.ALLOWDUPES
        val SHAREDUPES = XLSConstants.SHAREDUPES
    }
}