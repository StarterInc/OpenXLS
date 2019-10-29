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
package io.starter.OpenXLS;

import io.starter.formats.XLS.*;

import java.io.OutputStream;

/**
 * An interface representing an OpenXLS WorkBook.
 */
public interface WorkBook extends Handle, Document {

    int CALCULATE_ALWAYS = XLSConstants.CALCULATE_ALWAYS;
    int CALCULATE_EXPLICIT = XLSConstants.CALCULATE_EXPLICIT;
    int CALCULATE_AUTO = XLSConstants.CALCULATE_AUTO;
    String CALC_MODE_PROP = XLSConstants.CALC_MODE_PROP;
    String REFTRACK_PROP = XLSConstants.REFTRACK_PROP;
    String USETEMPFILE_PROP = XLSConstants.USETEMPFILE_PROP;
    String DEFAULTENCODING = XLSConstants.DEFAULTENCODING;
    String UNICODEENCODING = XLSConstants.UNICODEENCODING;
    String VALIDATEWORKBOOK = XLSConstants.VALIDATEWORKBOOK;

    // public static final int FORMULA_CALC_AUTO = 0;

    int STRING_ENCODING_AUTO = XLSConstants.STRING_ENCODING_AUTO;
    int STRING_ENCODING_UNICODE = XLSConstants.STRING_ENCODING_UNICODE;
    int STRING_ENCODING_COMPRESSED = XLSConstants.STRING_ENCODING_COMPRESSED;
    int ALLOWDUPES = XLSConstants.ALLOWDUPES;
    int SHAREDUPES = XLSConstants.SHAREDUPES;

    /**
     * Explicit calcing of formulas
     *
     * @param mode
     */
    void setFormulaCalculationMode(int mode);

    int getFormulaCalculationMode();

    /**
     * get a non-Excel property
     *
     * @return Returns the properties.
     */
    @Override
    Object getProperty(String name);

    /**
     * add non-Excel property
     *
     * @param properties The properties to set.
     */
    @Override
    void addProperty(String name, Object val);

    /**
     * The Session for the WorkBook instance
     *
     * @return public BookSession getSession();
     */

    java.awt.Color[] getColorTable();

    /**
     * Returns all of the Cells contained in the WorkBook
     */
    CellHandle[] getCells();

    /**
     * Returns the Cell at the specified Location
     *
     * @param address
     * @return
     */
    CellHandle getCell(String address) throws CellNotFoundException, WorkSheetNotFoundException;

    /**
     * Returns an Array of all the FormatHandles present in the workbook
     *
     * @return all existing FormatHandles in the workbook
     */
    FormatHandle[] getFormats();

    /**
     * Gets the date format used by this book.
     */
    DateConverter.DateFormat getDateFormat();

    /**
     * get a handle to a PivotTable in the WorkBook
     *
     * @param String name of the PivotTable
     * @return PivotTable the PivotTable
     */
    PivotTableHandle getPivotTable(String ptname) throws PivotTableNotFoundException;

    /**
     * get an array of handles to all PivotTables in the WorkBook
     *
     * @return PivotTable[] all of the WorkBooks PivotTables
     */
    PivotTableHandle[] getPivotTables() throws PivotTableNotFoundException;

    /**
     * set the workbook to protected mode
     * <p>
     * Note: the password cannot be decrypted or changed
     * in Excel -- protection can only be set/removed using
     * OpenXLS
     *
     * @param int Default Column width
     */
    void setProtected(boolean b);

    /**
     * set Default row height
     * <p>
     * Note: only affects undefined Rows containing Cells
     *
     * @param int Default Row Height
     */
    void setDefaultRowHeight(int t);

    /**
     * set Default col width
     * <p>
     * Note: only affects undefined Columns containing Cells
     *
     * @param int Default Column width
     */
    void setDefaultColWidth(int t);

    /**
     * Sets the internal name of this WorkBookHandle.
     * <p>
     * Overrides the default for 'getName()' which returns
     * the file name source of this WorkBook by default.
     *
     * @param WorkBook Name
     */
    @Override
    void setName(String nm);

    /**
     * Set the Debugging level.  Higher values output more
     * debugging info during execution.
     *
     * @parameter int Debug level.  higher=more verbose
     */
    @Override
    void setDebugLevel(int l);

    /**
     * Returns a Named Range Handle
     *
     * @return NameHandle a Named range in the WorkBook
     */
    NameHandle getNamedRange(String rangename) throws CellNotFoundException;

    /**
     * Returns a Chart Handle
     *
     * @return ChartHandle a Chart in the WorkBook
     */
    ChartHandle getChart(String chartname) throws ChartNotFoundException;

    /**
     * Returns all Chart Handles contained in the WorkBook
     *
     * @return ChartHandle[] an array of all Charts in the WorkBook
     */
    ChartHandle[] getCharts();

    /**
     * Returns all Named Range Handles
     *
     * @return NameHandle[] all of the Named ranges in the WorkBook
     */
    NameHandle[] getNamedRanges();

    /**
     * Returns the name of this WorkBook
     *
     * @return String name of WorkBook
     */
    @Override
    String getName();

    /**
     * Returns the number of Cells in this WorkBook
     *
     * @return int number of Cells
     */
    int getNumCells();

    /**
     * Returns a byte Array containing the
     * valid file containing this WorkBook
     * and associated Storages (such as VB files
     * and PivotTables.)
     * <p>
     * This is the actual file data and that can be
     * read from and written to FileOutputStreams and
     * ServletOutputStreams.
     *
     * @return byte[] the XLS File's bytes
     */
    byte[] getBytes();

    /** Initialize a Vector of the CellRanges existing in this WorkBook
     *  specifically the Ranges referenced in Formulas, Charts, and
     *  Named Ranges.
     *
     *  This is necessary to allow for automatic updating of references
     * 	when adding/removing/moving Cells within these ranges, as well
     * 	as shifting references to Cells in Formulas when Formula records
     *  are moved.
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
    WorkSheetHandle[] getWorkSheets();

    /**
     * returns the handle to a WorkSheet by name.
     *
     * @param index of worksheet (ie: 0)
     * @return WorkSheetHandle the WorkSheet
     * @throws WorkSheetNotFoundException if the specified WorkSheet is
     *                                    not found in the WorkBook.
     */
    WorkSheetHandle getWorkSheet(int i) throws WorkSheetNotFoundException;

    /**
     * returns the handle to a WorkSheet by name.
     *
     * @param String name of worksheet (ie: "Sheet1")
     * @return WorkSheetHandle the WorkSheet
     * @throws WorkSheetNotFoundException if the specified WorkSheet is
     *                                    not found in the WorkBook.
     */
    WorkSheetHandle getWorkSheet(String handstr) throws WorkSheetNotFoundException;

    /**
     * Returns a low-level WorkBook.
     * <p>
     * NOTE: The WorkBook class is NOT a part of the
     * published API.  Any of the methods and/or
     * variables on a WorkBook object are subject
     * to change without notice in new versions of OpenXLS.
     */
    io.starter.formats.XLS.WorkBook getWorkBook();

    /**
     * Clears all values in a template WorkBook.
     * <p>
     * Use this method to 'reset' the values of your
     * WorkBook in memory to defaults.
     * <p>
     * For example, if you load a Servlet with a
     * single WorkBookHandle instance, then modify
     * values and stream to a Client system, yo
     * should call 'clearAll()' when the request
     * is completed to remove the modified values
     * and set them back to a default.
     */
    @Override
    void reset();

    /**
     * Set Encoding mode of new Strings added to file.
     * <p>
     * OpenXLS has 3 modes for handling the internal encoding of
     * String data that is added to the file.
     * <p>
     * OpenXLS can save space in the file if it knows that all characters
     * in your String data can be represented with a single byte (Compressed.)
     * <p>
     * If your String contains characters which need 2 bytes to represent (such
     * as Eastern-language characters) then it needs to be stored in an uncompressed
     * Unicode format.
     * <p>
     * OpenXLS can either automatically detect the mode for each String, or you
     * can set it explicitly.  The auto mode is the most flexible but requires processing
     * overhead.
     * <p>
     * Default mode is WorkBookHandle.STRING_ENCODING_AUTO.
     * <p>
     * Valid Modes Are:
     * <p>
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
    void setStringEncodingMode(int mode);

    /**
     * Set Duplicate String Handling Mode.
     * <p>
     * The Duplicate String Mode determines the behavior of
     * the String table when inserting new Strings.
     * <p>
     * The String table shares a single entry for multiple
     * Cells containing the same string.  When multiple Cells
     * have the same value, they share the same underlying string.
     * <p>
     * Changing the value of any one of the Cells will change
     * the value for any Cells sharing that reference.
     * <p>
     * For this reason, you need to determine
     * the handling of new strings added to the sheet that
     * are duplicates of strings already in the table.
     * <p>
     * If you will be changing the values of these
     * new Cells, you will need to set the Duplicate
     * String Mode to ALLOWDUPES.  If the string table
     * encounters a duplicate entry being added, it
     * will insert a duplicate that can then be subsequently
     * changed without affecting the other duplicate Cells.
     * <p>
     * Valid Modes Are:
     * <p>
     * WorkBookHandle.ALLOWDUPES - faster, smaller file sizes, dupe Cells share changes
     * WorkBookHandle.SHAREDUPES - slower inserts, changing Cells has no effect on dupe Cells
     *
     * @param int Duplicate String Handling Mode
     */
    void setDupeStringMode(int mode);

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chartname
     * @param sheetname
     */
    void copyChartToSheet(String chartname, String sheetname) throws ChartNotFoundException, WorkSheetNotFoundException;

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chart
     * @param sheet
     */
    void copyChartToSheet(ChartHandle chart, WorkSheetHandle sheet) throws ChartNotFoundException, WorkSheetNotFoundException;

    /**
     * Copy (duplicate) a worksheet in the workbook and add it to the end of the workbook with a new name
     *
     * @param String the Name of the source worksheet;
     * @param String the Name of the new (destination) worksheet;
     * @return the new WorkSheetHandle
     */
    WorkSheetHandle copyWorkSheet(String SourceSheetName, String NewSheetName) throws WorkSheetNotFoundException;

    /**
     * Iterate through the formulas in this WorkBook and call the
     * calculate method on each.
     * <p>
     * May be more expensive than calling update on individual FormulaHandles
     * depending on extent of data changes to your WorkBook, or calling update
     * on only the 'top-level' formula in a calculation.  When a formula references
     * a Cell containing another formula, it will recursively calculate until it
     * reaches non-formula Cells.  Thus, calling this method may calculate formula
     * Cells in a heirarchy more than once.
     *
     * @throws FunctionNotSupportedException
     */
    void calculateFormulas() throws FunctionNotSupportedException;

    /**
     * Removes all of the WorkSheets from this WorkBook.
     * <p>
     * Bytes streamed from this WorkBook will create invalid
     * Spreadsheet files unless a WorkSheet(s) are added to it.
     */
    void removeAllWorkSheets();

    /**
     * Returns a WorkBookHandle containing an empty
     * version of this WorkBook.
     * <p>
     * Use in conjunction with  addSheetFromWorkBook() to create
     * new output WorkBooks containing various sheets from a master
     * template.
     * <p>
     * ie:
     * WorkBookHandle emptytemplate = this.getNoSheetWorkBook();
     * emptytemplate.addSheetFromWorkBook(this, "Sheet1", "TargetSheet");
     *
     * @return WorkBookHandle - the empty WorkBookHandle duplicate
     * @see addSheetFromWorkBook
     */
    WorkBookHandle getNoSheetWorkBook();

    /**
     * Inserts a worksheet from a Source WorkBook.
     *
     * @param sourceBook      - the WorkBook containing the sheet to copy
     * @param sourceSheetName - the name of the sheet to copy
     * @param destSheetName   - the name of the new sheet in this workbook
     * @throws WorkSheetNotFoundException
     */
    boolean addSheetFromWorkBook(WorkBookHandle sourceBook, String sourceSheetName, String destSheetName) throws WorkSheetNotFoundException;

    /**
     * Inserts a new worksheet and places it at the end of the workbook
     *
     * @param WorkSheetHandle the source WorkSheetHandle;
     * @param String          the Name of the new (destination) worksheet;
     */
    WorkSheetHandle addWorkSheet(WorkSheetHandle sht, String NewSheetName);

    /**
     * Creates a new worksheet and places it at the end of the workbook
     *
     * @param String the Name of the newly created worksheet
     * @return the new WorkSheetHandle
     */
    WorkSheetHandle createWorkSheet(String name);

    /**
     * Returns the name of this Sheet.
     *
     * @see java.lang.Object#toString()
     */
    @Override
    String toString();

    StringBuffer writeBytes(OutputStream bbout);
}