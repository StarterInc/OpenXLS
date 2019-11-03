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


import io.starter.formats.XLS.ColumnNotFoundException
import io.starter.formats.XLS.SxStreamID
import io.starter.formats.XLS.Sxview
import io.starter.toolkit.Logger
import org.json.JSONObject


/**
 * PivotTable Handle allows for manipulation of PivotTables within a WorkBook.
 *
 * @see WorkBookHandle
 *
 * @see WorkSheetHandle
 *
 * @see CellHandle
 */
/**
 *
 */
/**
 *
 */

/**
 *
 */
class PivotTableHandle
/**
 * Create a new PivotTableHandle from an Excel PivotTable
 *
 * @param PivotTable
 * - the PivotTable to create a handle for.
 */
(private val pt: Sxview, private val book: WorkBookHandle)//SxStreamID sxid= bk.getWorkBook().getPivotStrean(pt.getICache());
//cellRange= sxid.getCellRange();
{
    val workSheetHandle: WorkSheetHandle? = null

    /**
     * @return Returns the cellRange.
     */
    val dataSourceRange: CellRange?
        get() {
            val sxid = book.workBook!!.getPivotStream(pt.iCache.toInt())
            return sxid!!.cellRange
        }

    /**
     * get the Name of the PivotTable
     *
     * @return String - value of the PivotTable if stored as a String.
     */
    /**
     * set the Name of the PivotTable
     *
     * @param String
     * - value of the PivotTable if stored as a String.
     */
    var tableName: String?
        get() = pt.tableName
        set(tx) {
            pt.tableName = tx
        }

    /**
     * returns the name of the data field
     * @return
     */
    /**
     * sets the name of the data field for this pivot table
     * @param name
     */
    var dataName: String?
        get() = pt.dataName
        set(name) {
            pt.dataName = name
        }

    /**
     * get whether table displays row grand totals
     *
     * @return boolean whether table displays row grand total
     */
    /**
     * Takes a string as a current PivotTable location, and changes that pointer
     * in the PivotTable to the new string that is sent. This can take single
     * cells "A5" and cell ranges, "A3:d4" Returns true if the cell range
     * specified in PivotTableLoc exists & can be changed else false. This also
     * cannot change a cell pointer to a cell range or vice versa.
     *
     * @param String
     * - range of Cells within PivotTable to modify
     * @param String
     * - new range of Cells within PivotTable
     *
     *
     * public boolean changePivotTableLocation(String PivotTableLoc,
     * String newLoc) throws PivotTableNotFoundException{
     *
     * Logger.logInfo("Changing: " + PivotTableLoc + " to: " +
     * newLoc);
     *
     * return false;
     *
     * }
     */

    /**
     * set whether to display row grand totals
     *
     * @param boolean whether to display row grand totals
     */
    var rowsHaveGrandTotals: Boolean
        get() = pt.fRwGrand
        set(b) {
            pt.fRwGrand = b
        }

    /**
     * get whether table displays a column grand total
     *
     * @return boolean whether table displays column grand total
     */
    /**
     * set whether to display a column grand total
     *
     * @param boolean whether to display column grand total
     */
    var colsHaveGrandTotals: Boolean
        get() = pt.fColGrand
        set(b) = pt.setColGrand(b)

    /**
     * get the auto format for the Table
     *
     * @return int the auto format Id for the table
     */
    /**
     * set the auto format for the Table
     *
     * <pre>
     * The valid formats are:
     *
     * PivotTableHandle.AUTO_FORMAT_Report1
     * PivotTableHandle.AUTO_FORMAT_Report2
     * PivotTableHandle.AUTO_FORMAT_Report3
     * PivotTableHandle.AUTO_FORMAT_Report4
     * PivotTableHandle.AUTO_FORMAT_Report5
     * PivotTableHandle.AUTO_FORMAT_Report6
     * PivotTableHandle.AUTO_FORMAT_Report7
     * PivotTableHandle.AUTO_FORMAT_Report8
     * PivotTableHandle.AUTO_FORMAT_Report9
     * PivotTableHandle.AUTO_FORMAT_Report10
     * PivotTableHandle.AUTO_FORMAT_Table1
     * PivotTableHandle.AUTO_FORMAT_Table2
     * PivotTableHandle.AUTO_FORMAT_Table3
     * PivotTableHandle.AUTO_FORMAT_Table4
     * PivotTableHandle.AUTO_FORMAT_Table5
     * PivotTableHandle.AUTO_FORMAT_Table6
     * PivotTableHandle.AUTO_FORMAT_Table7
     * PivotTableHandle.AUTO_FORMAT_Table8
     * PivotTableHandle.AUTO_FORMAT_Table9
     * PivotTableHandle.AUTO_FORMAT_Table10
    </pre> *
     *
     * @param int the auto format Id for the table
     */
    var autoFormatId: Int
        get() = pt.itblAutoFmt.toInt()
        set(b) {
            pt.itblAutoFmt = b.toShort()
        }

    /**
     * get whether table has auto format applied
     *
     * @param boolean whether table has auto format applied
     */
    /**
     * set whether to auto format the Table
     *
     * @param boolean whether to auto format the table
     */
    var usesAutoFormat: Boolean
        get() = pt.fAutoFormat
        set(b) {
            pt.fAutoFormat = b
        }

    /**
     * get whether Width/Height Autoformat is applied
     *
     * @return boolean whether the Width/Height Autoformat is applied
     */
    /**
     * set Width/Height Autoformat is applied
     *
     * @param boolean whether to apply the Width/Height Autoformat
     */
    var autoWidthHeight: Boolean
        get() = pt.fwh
        set(b) {
            pt.fwh = b
        }

    /**
     * get whether Font Autoformat is applied
     *
     * @return boolean whether the Font Autoformat is applied
     */
    /**
     * set whether Font Autoformat is applied
     *
     * @param boolean whether to apply the Font Autoformat
     */
    var autoFont: Boolean
        get() = pt.fFont
        set(b) {
            pt.fFont = b
        }

    /**
     * get whether Alignment Autoformat is applied
     *
     * @return boolean whether the Alignment Autoformat is applied
     */
    /**
     * set whether Alignment Autoformat is applied
     *
     * @param boolean whether to apply the Alignment Autoformat
     */
    var autoAlign: Boolean
        get() = pt.fAlign
        set(b) {
            pt.fAlign = b
        }

    /**
     * get whether Border Autoformat is applied
     *
     * @return boolean whether the Border Autoformat is applied
     */
    /**
     * set whether Border Autoformat is applied
     *
     * @param boolean whether to apply the Border Autoformat
     */
    var autoBorder: Boolean
        get() = pt.fBorder
        set(b) {
            pt.fBorder = b
        }

    /**
     * get whether Pattern Autoformat is applied
     *
     * @return boolean whether the Pattern Autoformat is applied
     */
    /**
     * set whether Pattern Autoformat is applied
     *
     * @param boolean whether to apply the Pattern Autoformat
     */
    var autoPattern: Boolean
        get() = pt.fPattern
        set(b) {
            pt.fPattern = b
        }

    /**
     * get whether Number Autoformat is applied
     *
     * @return boolean whether the Number Autoformat is applied
     */
    /**
     * set whether Number Autoformat is applied
     *
     * @param boolean whether to apply the Number Autoformat
     */
    var autoNumber: Boolean
        get() = pt.fNumber
        set(b) {
            pt.fNumber = b
        }

    /**
     * get the first row in the PivotTable
     */
    /**
     * set the first row in the PivotTable
     */
    // s--; // these are zero-based rows
    var rowFirst: Int
        get() = pt.rwFirst.toInt() + 1
        set(s) {
            pt.rwFirst = s.toShort()
        }

    /**
     * get the last row in the PivotTable
     */
    /**
     * set the last row in the PivotTable
     */
    var rowLast: Int
        get() = pt.rwLast.toInt() + 1
        set(s) {
            var s = s
            s--
            pt.rwLast = s.toShort()
        }

    /**
     * get the first Column in the PivotTable
     */
    /**
     * set the first Column in the PivotTable
     */
    var colFirst: Int
        get() = pt.colFirst.toInt()
        set(s) {
            pt.colFirst = s.toShort()
        }

    /**
     * get the last Column in the PivotTable
     */
    /**
     * set the last Column in the PivotTable
     */
    var colLast: Int
        get() = pt.colLast.toInt()
        set(s) {
            pt.colLast = s.toShort()
        }

    /**
     * get the first header row
     */
    /**
     * set the first header row
     */
    var rowFirstHead: Int
        get() = pt.rwFirstHead.toInt() + 1
        set(s) {
            var s = s
            s--
            pt.rwFirstHead = s.toShort()
        }

    /**
     * get the first Row containing data
     */
    /**
     * set the first Row containing data
     */
    // zero-based rows
    var rowFirstData: Int
        get() = pt.rwFirstData.toInt() + 1
        set(s) {
            var s = s
            s--
            pt.rwFirstData = s.toShort()
        }

    /**
     * get the first Column containing data
     */
    /**
     * set the first Column containing data
     */
    var colFirstData: Int
        get() = pt.colFirstData.toInt()
        set(s) {
            pt.colFirstData = s.toShort()
        }

    /**
     * returns the JSON representation of this PivotTable
     *
     * @return JSON Pivot Table representation
     */
    // copy all methods to this sucker
    //			thePivot.put("cellrange", this.getCellRange().getRange());
    val json: String
        get() {
            val thePivot = JSONObject()

            try {
                thePivot.put("title", this.tableName)

            } catch (e: Exception) {
                throw WorkBookException("PivotTableHandle.getJSON failed:$e",
                        WorkBookException.RUNTIME_ERROR)
            }

            return thePivot.toString()
        }

    /**
     * Sets the Pivot Table Range to represent the Data to analyse
     * <br></br>NOTE: any existing data will be replaced
     * @param cellRange The cellRange to set.
     */
    fun setSourceDataRange(cellRange: CellRange) {
        val sxid = book.workBook!!.getPivotStream(pt.iCache.toInt())
        sxid!!.cellRange = cellRange
        try {
            pt.nPivotFields = cellRange.cols.size.toShort()
        } catch (e: ColumnNotFoundException) {

        }

    }

    /**
     * Sets the Pivot Table Range to represent the Data to analyse
     * <br></br>NOTE: any existing data will be replaced
     * <br></br>If the cell range does not contain sheet information, the sheet that the pivot table is located will be used
     * @param cellRange
     */
    fun setSourceDataRange(range: String) {
        var range = range
        val rc = ExcelTools.getRangeCoords(range)
        if (range.indexOf("!") == -1)
            range = this.workSheetHandle.toString() + "!" + range
        val sxid = book.workBook!!.getPivotStream(pt.iCache.toInt())
        sxid!!.setCellRange(range)

        pt.nPivotFields = (rc[3] - rc[1] + 1).toShort()
    }

    /**
     * Sets the Pivot Table data source from a named range
     * @param namedrange    Named Range
     */
    fun setSource(namedrange: String) {
        // TODO: finish; update DCONNAME
    }

    /**
     * returns whether a given row is contained in this PivotTable
     *
     * @param int the row number
     * @return boolean whether the row is in the table
     */
    fun containsRow(x: Int): Boolean {
        return x <= pt.rwLast && x >= pt.rwFirst
    }

    /**
     * return a more friendly
     *
     * @see java.lang.Object.toString
     */
    override fun toString(): String? {
        return this.tableName
    }

    /**
     * returns whether a given col is contained in this PivotTable
     *
     * @param int the column number
     * @return boolean whether the col is in the table
     */
    fun containsCol(x: Int): Boolean {
        return x <= pt.colLast && x >= pt.colFirst
    }

    fun removeArtifacts() {
        val coords = intArrayOf(pt.rwFirst - 2, pt.colFirst.toInt(), pt.rwLast.toInt(), pt.colLast.toInt())
        try {
            val newr = CellRange(workSheetHandle!!, coords, true)
            val ch = newr.getCells()
            for (r in ch!!.indices) {
                if (ch[r] != null)
                    ch[r].remove(true)
            }
        } catch (e: Exception) {
            Logger.logWarn("could not remove artifacts in PivotTableHandle: $e")
        }

    }

    companion object {

        var AUTO_FORMAT_Report1 = 1
        var AUTO_FORMAT_Report2 = 2
        var AUTO_FORMAT_Report3 = 3
        var AUTO_FORMAT_Report4 = 4
        var AUTO_FORMAT_Report5 = 5
        var AUTO_FORMAT_Report6 = 6
        var AUTO_FORMAT_Report7 = 7
        var AUTO_FORMAT_Report8 = 8
        var AUTO_FORMAT_Report9 = 9
        var AUTO_FORMAT_Report10 = 10
        var AUTO_FORMAT_Table1 = 11
        var AUTO_FORMAT_Table2 = 12
        var AUTO_FORMAT_Table3 = 13
        var AUTO_FORMAT_Table4 = 14
        var AUTO_FORMAT_Table5 = 15
        var AUTO_FORMAT_Table6 = 16
        var AUTO_FORMAT_Table7 = 17
        var AUTO_FORMAT_Table8 = 18
        var AUTO_FORMAT_Table9 = 19
        var AUTO_FORMAT_Table10 = 20
        var AUTO_FORMAT_Classic = 30
    }


}