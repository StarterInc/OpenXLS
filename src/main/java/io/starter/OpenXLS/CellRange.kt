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
import io.starter.formats.XLS.formulas.FormulaParser
import io.starter.formats.XLS.formulas.GenericPtg
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgRef
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject

import java.io.Serializable
import java.util.*

import io.starter.OpenXLS.JSONConstants.*

/**
 * Cell Range is a handle to a range of Workbook Cells
 *
 * <br></br>
 * Contains useful methods for working with Collections of Cells. <br></br>
 * <br></br>
 * for example: <br></br>
 * <br></br>
 * <blockquote> CellRange cr = new CellRange("Sheet1!A1:B10", workbk);<br></br>
 * cr.addCellToRange("C10");<br></br>
 * CellHandle mycells = cr.getCells();<br></br>
 * for(int x=0;x < mycells.length;x++) <br></br>
 * Logger.logInfo(mycells[x].getCellAddress() + mycells[x].toString());<br></br>
 * }<br></br>
 * <br></br>
</blockquote> *
 *
 * [Starter Inc.](http://starter.io)
 *
 * @see DataBoundCellRange
 *
 * @see XLSRecord
 */
class CellRange : Serializable {

    /**
     * returns the conditional format object for this range, if any
     *
     * @return Condfmt object
     */
    protected var conditionalFormat: Condfmt? = null
        internal set
    /**
     * returns the merged state of the CellRange
     *
     * @return true if this CellRange is merged
     */
    var isMerged = false
        private set
    private var parent: BiffRec? = null // if cell range is child of a named range, must ensure update correctly
    // private Ptg myptg = null;
    var DEBUG = false
    private var isDirty = false // true if addCellsToRange without init
    internal var firstcellrow = -1
    internal var firstcellcol = -1
    internal var lastcellrow = -1
    internal var lastcellcol = -1
    /**
     * gets whether this CellRange will add blank records to the WorkBook for any
     * missing Cells contained within the range.
     *
     * @return true if should create blank records for missing Cells
     */
    /**
     * set whether this CellRange will add blank records to the WorkBook for any
     * missing Cells contained within the range.
     *
     * @param boolean b - true if should create blank records for missing Cells
     */
    var createBlanks = false
    private var initializeCells = true
    @Transient
    var cells: Array<CellHandle>? = null
    @Transient
    protected var range: String? = null
    @Transient
    protected var sheetname: String? = null
    /**
     * @return the workbook object attached to this CellRange
     */
    /**
     * attach the workbook for this CellRange
     *
     * @param WorkBook bk
     */
    @Transient
    var workBook: WorkBook? = null
    @Transient
    private var sheet: WorkSheetHandle? = null
    private var myrowints: IntArray? = null
    private var mycolints: IntArray? = null
    @Transient
    private var myrows: Array<RowHandle>? = null
    @Transient
    private var mycols: Array<ColHandle>? = null

    // for OOXML External References
    protected var externalLink1 = 0
    protected var externalLink2 = 0

    internal var fmtr: FormatHandle? = null
    internal var wholeCol = false
    internal var wholeRow = false

    /**
     * Gets the number of columns in the range.
     */
    val width: Int
        get() = lastcellcol - firstcellcol + 1

    /**
     * Gets the number of rows in the range.
     */
    val height: Int
        get() = lastcellrow - firstcellrow + 1

    /**
     * Returns an array of the row numbers referenced by this CellRange
     *
     * @return int[] array of row ints
     */
    val rowInts: IntArray
        get() {
            if (myrowints != null)
                return myrowints
            val numrows = lastcellrow + 1 - firstcellrow
            myrowints = IntArray(numrows)
            for (t in 0 until numrows) {
                myrowints[t] = firstcellrow + t
            }
            return myrowints
        }

    /**
     * returns an array of column numbers referenced by this CellRange
     *
     * @return int[] array of col ints
     */
    val colInts: IntArray
        get() {
            if (mycolints != null)
                return mycolints
            val numcols = lastcellcol + 1 - firstcellcol
            mycolints = IntArray(numcols)
            for (t in 0 until numcols) {
                mycolints[t] = firstcellcol + t
            }
            return mycolints
        }

    /**
     * Returns an array of Rows (RowHandles) referenced by this CellRange
     *
     * @return RowHandle[] array of row handles
     */
    // typically empty rows
    val rows: Array<RowHandle>
        @Throws(RowNotFoundException::class)
        get() {
            if (myrows != null)
                return myrows
            val numrows = lastcellrow + 1 - firstcellrow
            myrows = arrayOfNulls(numrows)
            for (t in 0 until numrows) {
                var rx: RowHandle? = null
                try {
                    rx = sheet!!.getRow(firstcellrow - 1 + t)
                } catch (x: Exception) {
                }

                myrows[t] = rx
            }
            return myrows
        }

    /**
     * Get the number of rows that this CellRange encompasses
     *
     * @return
     */
    val numberOfRows: Int
        get() = lastcellrow + 1 - firstcellrow

    /**
     * Returns an array of Columns (ColHandles) referenced by this CellRange
     *
     * @return ColHandle[] array of columns handles
     */
    val cols: Array<ColHandle>
        @Throws(ColumnNotFoundException::class)
        get() {
            if (mycols != null)
                return mycols
            val numcols = lastcellcol + 1 - firstcellcol
            mycols = arrayOfNulls(numcols)
            for (t in 0 until numcols) {
                mycols[t] = sheet!!.getCol(firstcellcol + t)
            }
            return mycols
        }

    /**
     * Get the number of columns that this CellRange encompasses
     *
     * @return
     */
    val numberOfCols: Int
        get() = lastcellcol + 1 - firstcellcol

    /**
     * Return a list of the cells in this cell range
     *
     * @return List of CellHandles
     */
    val cellList: List<CellHandle>
        get() = Arrays.asList(*cells!!)

    /**
     * get the underlying Cell Records in this range <br></br>
     * NOTE: Cell Records are not a part of the public API and are not intended for
     * use in client applications.
     *
     * @return BiffRec[] array of Cell Records
     */
    val cellRecs: Array<BiffRec>
        get() {
            val ch = this.getCells()
            val ret = arrayOfNulls<BiffRec>(ch!!.size)
            for (t in ret.indices) {
                if (ch[t] != null)
                    ret[t] = ch[t].cell
            }
            return ret
        }

    /*
     * reset the underlying cell records in this range <br>NOTE: Cell Records are
     * not a part of the public API and are not intended for use in client
     * applications.
     *
     * @return BiffRec[] array of Cell Records
     *
     * NOT USED AT THIS TIME public BiffRec[] resetCellRecs() { this.isDirty= true;
     * return getCellRecs(); }
     */

    /**
     * If the cells contain an incrementing value that can be transferred into an
     * int, then return that value, else throw a NPE. I'm sure there is a better
     * exception to be thrown, but not sure what that is, and it doesn't make sense
     * to return a value like -1 in these cases.
     */
    val incrementAmount: Int
        @Throws(Exception::class)
        get() {
            val ch = this.getCells()
            if (ch!!.size == 1) {
                throw Exception("Cannot have increment with non-range cell")
            }
            var initialized = false
            var incAmount = 0
            for (i in 1 until ch.size) {
                val value1 = ch[i - 1].intVal
                val value2 = ch[i].intVal
                if (!initialized) {
                    incAmount = value2 - value1
                    initialized = true
                } else {
                    if (value2 - value1 != incAmount) {
                        throw Exception("Inconsistent values across increment")
                    }
                }
            }
            if (!initialized) {
                throw Exception("Error determining increment")
            }
            return incAmount
        }

    /**
     * Return the XML representation of this CellRange object
     *
     * @return String of XML
     */
    // StringBuffer sb = new StringBuffer(xmlResponsePre);
    // append cellxml
    val xml: String
        get() {
            val sb = StringBuffer("<CellRange Range=\"" + this.getRange() + "\">")
            val cx = this.getCells()
            sb.append("\r\n")
            for (t in cx!!.indices) {
                sb.append(cx[t].xml)
                sb.append("\r\n")
            }
            sb.append(xmlResponsePost)
            return sb.toString()
        }

    /**
     * gets the array of Cells in this Name
     *
     *
     * Thus this method is limited to use with 2D ranges.
     *
     * @param fragment whether to enclose result in NameHandle tag
     * @return Cell[] all Cells defined in this Name
     */
    val cellRangeXML: String
        get() {
            val sbx = StringBuffer()
            sbx.append("<?xml version=\"1\" encoding=\"utf-8\"?>")
            sbx.append(xml)
            return sbx.toString()
        }

    /**
     * Gets the coordinates of this cell range,
     *
     * @return int[5]: [0] first row (zero based, ie row 1=0), [1] first column, [2]
     * last row (zero based, ie row 1=0), [3] last column, [4] number of
     * cells in range
     */
    // qualify sheet and reset range - necessary if sheetname with spaces is used in
    // formula parsing
    /*
         * check for R1C1
         */// no range
    // get the first cell's coordinates
    // get the last cell's coordinates
    // handle swapped cells ie: "B1:A1"
    // not an error if it is a whole column or whole row range
    // what should numcells be for wholerow?
    // what should numcells be for wholecol?
    // trap OOXML external reference link, if any
    val coords: IntArray
        @Throws(CellNotFoundException::class)
        get() {
            var numrows = 0
            var numcols = 0
            var numcells = 0
            val coords = IntArray(5)
            var temprange = range
            val s = ExcelTools.stripSheetNameFromRange(temprange!!)
            temprange = s[1]
            sheetname = GenericPtg.qualifySheetname(s[0])
            if (sheetname != null && sheetname != "") {
                if (s[2] == null)
                    this.range = "$sheetname!$temprange"
                else {
                    s[2] = GenericPtg.qualifySheetname(s[2])
                    this.range = sheetname + ":" + s[2] + "!" + temprange
                }
            }
            if (temprange!!.indexOf("R") == 0 && temprange.indexOf("C") > 1
                    && Character.isDigit(temprange[temprange.indexOf("C") - 1])) {

                val b = ExcelTools.getRangeRowCol(temprange)

                numrows = b[2] - b[0]
                if (numrows <= 0)
                    numrows = 1

                numcols = b[3] - b[2]
                if (numcols <= 0)
                    numcols = 1

                numcells = numrows * numcols

                val retr = IntArray(5)
                retr[0] = b[0]
                retr[1] = b[1]
                retr[2] = b[2]
                retr[3] = b[3]
                retr[4] = numcells
                return retr
            }
            var startcell = ""
            var endcell = ""
            val lastcolon = temprange.lastIndexOf(":")
            endcell = temprange.substring(lastcolon + 1)
            if (lastcolon == -1)
                startcell = endcell
            else
                startcell = temprange.substring(0, lastcolon)

            startcell = StringTool.strip(startcell, "$")
            endcell = StringTool.strip(endcell, "$")
            var charct = startcell.length
            while (charct > 0) {
                if (!Character.isDigit(startcell[--charct])) {
                    charct++
                    break
                }
            }
            val firstcellrowstr = startcell.substring(charct)
            firstcellrow = Integer.parseInt(firstcellrowstr)
            val firstcellcolstr = startcell.substring(0, charct)
            firstcellcol = ExcelTools.getIntVal(firstcellcolstr)
            charct = endcell.length
            while (charct > 0) {
                if (!Character.isDigit(endcell[--charct])) {
                    charct++
                    break
                }
            }
            val lastcellrowstr = endcell.substring(charct)
            lastcellrow = Integer.parseInt(lastcellrowstr)
            val lastcellcolstr = endcell.substring(0, charct)
            lastcellcol = ExcelTools.getIntVal(lastcellcolstr)
            numrows = lastcellrow - firstcellrow + 1
            numcols = lastcellcol - firstcellcol + 1

            numcells = numrows * numcols
            if (numcells < 0)
                numcells *= -1

            coords[0] = firstcellrow - 1
            coords[1] = firstcellcol
            coords[2] = lastcellrow - 1
            coords[3] = lastcellcol
            coords[4] = numcells
            if (firstcellrow < 0 && lastcellrow < 0 || firstcellcol < 0 || lastcellcol < 0) {
                if (firstcellcol == -1 && lastcellcol == -1) {
                    wholeRow = true
                } else if (firstcellrow == -1 && lastcellrow == -1) {
                    wholeCol = true
                } else
                    Logger.logErr("CellRange.getRangeCoords: Error in Range " + range!!)
            }
            if (s[3] != null)
                externalLink1 = Integer.valueOf(s[3].substring(1, s[3].length - 1)).toInt()
            if (s[4] != null)
                externalLink2 = Integer.valueOf(s[4].substring(1, s[4].length - 1)).toInt()

            return coords

        }

    /**
     * Gets the coordinates of this cell range.
     *
     * @return int[5]: [0] first row, [1] first column, [2] last row, [3] last
     * column, [4] number of cells in range
     */
    val rangeCoords: IntArray
        @Deprecated("{@link #getCoords()} instead, which returns zero based values for\n" +
                "      rows.")
        @Throws(CellNotFoundException::class)
        get() {
            val ordinalValues = this.coords
            ordinalValues[0] += 1
            ordinalValues[2] += 1
            return ordinalValues
        }

    /**
     * Return the String cell address of this range in R1C1 format
     *
     * @return String range in R1C1 format
     */
    // rangecoords are already 1-based
    val r1C1Range: String
        get() {
            var rc1x = "R"
            try {
                val rc1 = this.rangeCoords
                rc1x += rc1[0] + 1
                rc1x += "C" + rc1[1]
                rc1x += ":R" + (rc1[2] + 1)
                rc1x += "C" + rc1[3]

            } catch (e: CellNotFoundException) {
                Logger.logErr("CellRange.getR1C1Range failed", e)
            }

            return rc1x
        }

    /**
     * Return a json object representing this cell range, entries contain only
     * address and values for more compact space
     *
     * @param range
     * @param wbh
     * @return
     */
    // should this possibly be full
    val basicJSON: JSONObject
        get() {
            try {
                val crObj = JSONObject()
                crObj.put(JSON_RANGE, this.getRange())
                val rangeArray = JSONArray()
                val cells = this.getCells()
                for (j in cells!!.indices) {
                    val result = JSONObject()
                    val addy = cells[j].cellAddress
                    val `val` = cells[j].`val`!!.toString()
                    result.put(JSON_LOCATION, addy)
                    result.put(JSON_CELL_VALUE, `val`)
                    rangeArray.put(result)
                }
                crObj.put(JSON_CELLS, rangeArray)
                return crObj
            } catch (e: Exception) {
                Logger.logErr("Error obtaining CellRange $range JSON: $e")
            }

            return JSONObject()
        }

    /**
     * Return a json object representing this cell range with full cell information
     * embedded.
     */
    val json: JSONObject
        get() {
            val theRange = JSONObject()
            val cells = JSONArray()
            try {
                theRange.put(JSON_RANGE, getRange())
                val chandles = getCells()
                for (i in chandles!!.indices) {
                    val thisCell = chandles[i]
                    val result = JSONObject()

                    result.put(JSON_CELL, thisCell.jsonObject)
                    cells.put(result)
                }
                theRange.put(JSON_CELLS, cells)
            } catch (e: JSONException) {
                Logger.logErr("Error getting cellRange JSON: $e")
            }

            return theRange
        }

    /**
     * Protected constructor for creating result ranges.
     */
    protected constructor(sheet: WorkSheetHandle, row: Int, col: Int, width: Int, height: Int) {
        this.sheet = sheet
        sheetname = sheet.sheetName
        workBook = sheet.workBook

        firstcellrow = row
        firstcellcol = col
        lastcellrow = row + height - 1
        lastcellcol = col + width - 1

        range = (sheetname + "!"
                + ExcelTools.formatRange(intArrayOf(firstcellcol, firstcellrow, lastcellcol, lastcellrow)))

        cells = arrayOfNulls(width * height)
    }

    /**
     * Initializes a `CellRange` from a `CellRangeRef`. The
     * source `CellRangeRef` instance must be qualified with a single
     * resolved worksheet.
     *
     * @param source the `CellRangeRef` from which to initialize
     * @param init   whether to populate the cell array
     * @param create whether to fill gaps in the range with blank cells
     * @throws IllegalArgumentException if the source range does not have a resolved sheet or has more
     * than one sheet
    `` */
    @JvmOverloads
    constructor(source: CellRangeRef, init: Boolean = false, create: Boolean = true) {
        initializeCells = init
        createBlanks = create

        sheet = source.firstSheet
        if (sheet == null || source.isMultiSheet)
            throw IllegalArgumentException("the source range must have a single resolved sheet")

        workBook = this.sheet!!.workBook
        sheetname = this.sheet!!.sheetName

        // This is inefficient, but fixing it would require rewriting init.
        this.range = source.toString()

        try {
            this.init()
        } catch (e: CellNotFoundException) {
            // this should be impossible
            throw RuntimeException(e)
        }

    }

    fun clearFormats() {
        for (idx in cells!!.indices)
            if (cells!![idx] != null)
                cells!![idx].clearFormats()
    }


    @Deprecated("use clear()")
    fun clearContents() {
        for (idx in cells!!.indices)
            if (cells!![idx] != null)
                cells!![idx].clearContents()
    }

    /**
     * clears the contents and formats of the cells referenced by this range but
     * does not remove the cells from the workbook.
     */
    fun clear() {
        for (idx in cells!!.indices)
            if (cells!![idx] != null)
                cells!![idx].clear()
    }

    /**
     * removes the cells referenced by this range from the sheet.
     *
     *
     * NOTE: method does not shift rows or cols.
     */
    fun removeCells() {
        for (idx in cells!!.indices)
            if (cells!![idx] != null)
                cells!![idx].remove(true)
    }

    /**
     * Un-Merge the Cells contained in this CellRange
     *
     * @throws Exception
     */
    @Throws(Exception::class)
    fun unMergeCells() {
        val mycells = this.cellRecs
        for (t in mycells.indices) {
            mycells[t].mergeRange = null // unset the range of merged cells
            mycells[t].xfRec.merged = false
        }
        val mc = this.getSheet()!!.sheet!!.mergedCellsRec
        mc?.removeCellRange(this)
        this.isMerged = false
    }

    /**
     * Set the format ID of all cells in this CellRange <br></br>
     * FormatID can be obtained through any CellHandle with the getFormatID() call
     *
     * @param int fmtID - the format ID to set the cells within the range to
     */
    @Throws(Exception::class)
    fun setFormatID(fmtID: Int) {
        val mycells = this.cellRecs
        for (t in mycells.indices) {
            mycells[t].setXFRecord(fmtID)
        }
    }

    /**
     * Set a hyperlink on all cells in this CellRange
     *
     * @param String url - the URL String to set
     */
    @Throws(Exception::class)
    fun setURL(url: String) {
        val mycells = this.cellRecs
        for (t in mycells.indices) {
            CellHandle(mycells[t], this.workBook).url = url
        }
    }

    /**
     * Merge the Cells contained in this CellRange
     *
     * @param boolean remove - true to delete the Cells following the first in the range
     */
    fun mergeCells(remove: Boolean) {
        this.createBlanks()
        if (remove)
            this.mergeCellsClearFollowingCells()
        else
            this.mergeCellsKeepFollowingCells()
    }

    /**
     * Merge the Cells contained in this CellRange, clearing or removing the Cells
     * following the first in the range
     */
    private fun mergeCellsClearFollowingCells() {
        val mycells = this.cellRecs

        // mycells[0].setMergeRange(this); // set the range of merged cells
        var r: Xf? = mycells[0].xfRec
        if (r == null) {
            fmtr = FormatHandle(this.workBook)
            fmtr!!.addCell(mycells[0])
            r = mycells[0].xfRec
        }
        r!!.merged = true
        for (t in mycells.indices) {
            if (mycells[t] != null)
                mycells[t].setSheet(this.getSheet()!!.mysheet)
            mycells[t].mergeRange = this // set the range of merged cells
            if (t > 0) {
                if (mycells[t] !is Blank) {
                    val cellname = mycells[t].cellAddress
                    val sheet = mycells[t].sheet
                    mycells[t].remove(true) // blow it out!
                    sheet.addValue(null, cellname)
                }
            }
        }
        var mc = this.getSheet()!!.sheet!!.mergedCellsRec
        if (mc == null)
            mc = this.getSheet()!!.sheet!!.addMergedCellRec()
        mc!!.addCellRange(this)
        this.isMerged = true
    }

    /**
     * Merge the Cells contained in this CellRange *
     */
    private fun mergeCellsKeepFollowingCells() {
        val mycells = this.cellRecs
        for (t in mycells.indices) {
            mycells[t].mergeRange = this // set the range of merged cells
            var r: Xf? = mycells[t].xfRec
            if (r == null) {
                fmtr = FormatHandle(this.workBook)
                fmtr!!.addCellRange(this)
                r = mycells[t].xfRec
            }
            r!!.merged = true
        }

        var mc = this.getSheet()!!.sheet!!.mergedCellsRec
        if (mc == null)
            mc = this.getSheet()!!.sheet!!.addMergedCellRec()
        mc!!.addCellRange(this)
        this.isMerged = true
    }

    /**
     * returns edge status of the desired CellHandle within this CellRange ie: top,
     * left, bottom, right <br></br>
     * returns 0 or 1 for 4 sides <br></br>
     * 1,1,1,1 is a single cell in a range 1,1,0,0 is on the top left edge of the
     * range
     *
     * @param CellHandle ch -
     * @param int        sz -
     * @return int[] array representing the edge positions
     */
    // TODO: documentation: Don't quite understand this!
    fun getEdgePositions(ch: CellHandle, sz: Int): IntArray {
        val coords = intArrayOf(0, 0, 0, 0)
        // get the corners, check for 'edges'
        val adr = ch.cellAddress
        val rc = ExcelTools.getRowColFromString(adr)
        // increment to one-based
        rc[0]++
        if (rc[0] == firstcellrow)
            coords[0] = sz
        if (rc[0] == lastcellrow)
            coords[2] = sz
        if (rc[1] == firstcellcol)
            coords[1] = sz
        if (rc[1] == lastcellcol)
            coords[3] = sz
        return coords
    }

    /**
     * returns whether this CellRange intersects with another CellRange
     *
     * @param CellRange cr - CellRange to test
     * @return boolean true if CellRange cr intersects with this CellRange
     */
    fun intersects(cr: CellRange): Boolean {
        // get the corners, check for 'contains'
        try {
            val rc = cr.rangeCoords
            if (rc[0] >= firstcellrow && rc[2] <= lastcellrow && rc[1] >= firstcellcol
                    && rc[3] <= lastcellcol) {
                return true
            }
        } catch (e: CellNotFoundException) {
            Logger.logWarn("CellRange unable to determine intersection of range: " + cr.toString()!!)
        }

        return false
    }

    /**
     * returns whether this CellRange contains a particular Cell
     *
     * @param CellHandle ch - the Cell to check
     * @return true if CellHandle ch is contained within this CellRange
     */
    operator fun contains(cxx: Cell): Boolean {
        val chsheet = cxx.workSheetName
        var mysheet = ""
        if (this.getSheet() != null)
            mysheet = this.getSheet()!!.sheetName
        if (!chsheet.equals(mysheet, ignoreCase = true))
            return false
        val adr = cxx.cellAddress
        val rc = ExcelTools.getRowColFromString(adr)
        return contains(rc)
    }

    /**
     * returns whether this CellRange contains the specified Row/Col coordinates
     *
     * @param int[] rc - row/col coordinates to test
     * @return true if the coordinates are contained with this CellRange
     */
    operator fun contains(rc: IntArray): Boolean {
        var ret = true
        if (rc[0] + 1 < firstcellrow)
            ret = false
        if (rc[0] + 1 > lastcellrow)
            ret = false
        if (rc[1] < firstcellcol)
            ret = false
        if (rc[1] > lastcellcol)
            ret = false
        return ret
    }

    /**
     * returns the String representation of this CellRange
     */
    override fun toString(): String? {
        return range
    }

    /**
     * Constructor to create a new CellRange from a WorkSheetHandle and a set of
     * range coordinates: <br></br>
     * coords[0] = first row <br></br>
     * coords[1] = first col <br></br>
     * coords[2] = last row <br></br>
     * coords[3] = last col
     *
     * @param WorkSheetHandle sht - handle to the WorkSheet containing the Range's Cells
     * @param int[]           coords - the cell coordinates
     * @param boolean         cb - true if should create blank cells
     * @throws Exception
     */
    @Throws(Exception::class)
    @JvmOverloads
    constructor(sht: WorkSheetHandle, coords: IntArray, cb: Boolean = false) {
        this.createBlanks = cb
        this.sheet = sht
        this.workBook = sht.workBook
        sheetname = sht.sheetName
        sheetname = GenericPtg.qualifySheetname(sheetname)
        var addr = sheetname!! + "!"
        val c1 = ExcelTools.getAlphaVal(coords[1]) + (coords[0] + 1)
        val c2 = ExcelTools.getAlphaVal(coords[3]) + (coords[2] + 1)
        addr += "$c1:$c2"
        this.range = addr
        this.init()
    }

    /**
     * Set this CellRange to be the current Print Area
     */
    fun setAsPrintArea() {
        if (this.sheet == null) {
            Logger.logErr("CellRange.setAsPrintArea() failed: " + this.toString()
                    + " does not have a valid Sheet reference.")
            return
        }
        sheet!!.setPrintArea(this)
    }

    /**
     * Constructor to Create a new CellRange from a String range <br></br>
     * The String range must be in the format Sheet!CR:CR <br></br>
     * For Example, "Sheet1!C9:I19" <br></br>
     * NOTE: You MUST Set the WorkBookHandle explicitly on this CellRange or it will
     * generate NullPointerException when trying to access the Cells.
     *
     * @param String r - range String
     * @see CellRange.setWorkBook
     */
    constructor(r: String) {
        this.range = r
    }

    /**
     * Increase the bounds of the CellRange by including the CellHandle. <br></br>
     * These are the limitations and side-effects of this method: <br></br>
     * - the Cell must be contiguous with the existing Range, ie: you can add a Cell
     * which either increments the row or the column of the existing range by one.
     * <br></br>
     * - the Cell must be on the same sheet as the existing range. <br></br>
     * - as a Cell Range is a 2 dimensional rectangle, expanding a multiple column
     * range by adding a Cell to the end will include the logical Cells on the row
     * in the range. <br></br>
     * Some Examples: <br></br>
     * <br></br>
     * // simple one dimensional range expansion: <br></br>
     * existing Range = A1:A20 <br></br>
     * addCellToRange(A21) new Range = A1:A21 <br></br>
     * <br></br>
     * existing Range = A1:B20 <br></br>
     * addCellToRange(A21) <br></br>
     * new Range = A1:B21 // note B20 is included automatically <br></br>
     * <br></br>
     * existing Range = A1:A20 <br></br>
     * addCellToRange(B1) <br></br>
     * new Range = A1:B20 //note entire B column of cells are included automatically
     *
     * @param CellHandle ch - the Cell to add to the CellRange
     */
    fun addCellToRange(ch: CellHandle): Boolean {
        // check worksheet
        val sheetname = ch.workSheetName
        if (sheetname == null) {
            Logger.logWarn("Cell $ch NOT added to Range: $this")
            return false
        }
        if (!sheetname.equals(this.getSheet()!!.sheetName, ignoreCase = true)) {
            Logger.logWarn("Cell $ch NOT added to Range: $this")
            return false
        }
        val rc = intArrayOf(ch.rowNum, ch.colNum)
        // increment to one-based
        rc[0]++

        // check that it's at the beginning
        if (firstcellrow == -1)
            firstcellrow = rc[0]
        if (firstcellcol == -1)
            firstcellcol = rc[1]
        if (lastcellrow == -1)
            lastcellrow = rc[0]
        if (lastcellcol == -1)
            lastcellcol = rc[1]
        if (rc[0] < firstcellrow)
            firstcellrow--
        if (rc[1] < firstcellcol)
            firstcellcol--
        // check that it's at the end
        if (rc[0] > lastcellrow)
            lastcellrow++
        if (rc[1] > lastcellcol)
            lastcellcol++

        // myptg is never set so myptg access never happens... taking out
        // boolean addPtgInfo = false;
        // if (myptg != null && myptg instanceof PtgRef)
        // addPtgInfo = true;
        // format the new range String
        val newrange = this.getSheet()!!.sheetName + "!"
        var newcellrange = ""
        // if (addPtgInfo && !((PtgRef) myptg).isColRel())
        // newcellrange += "$";
        newcellrange += ExcelTools.getAlphaVal(firstcellcol)
        // if (addPtgInfo && !((PtgRef) myptg).isRowRel())
        // newcellrange += "$";
        newcellrange += firstcellrow.toString()
        newcellrange += ":"
        // if (addPtgInfo && !((PtgRef) myptg).isColRel())
        // newcellrange += "$";
        newcellrange += ExcelTools.getAlphaVal(lastcellcol)
        // if (addPtgInfo && !((PtgRef) myptg).isColRel())
        // newcellrange += "$";
        newcellrange += lastcellrow.toString()
        this.range = newrange + newcellrange
        isDirty = true

        /*
         * if (addPtgInfo) { ReferenceTracker.updateAddressPerPolicy(myptg,
         * newcellrange); return true; }
         */

        if (this.parent != null && this.parent!!.opcode == XLSConstants.NAME) {
            (parent as Name).location = this.range // ensure named range expression is updated, as well as any formula
            // references are cleared of cache
        }

        return false
    }

    /**
     * get the Cells in this cell range
     *
     * @return CellHandle[] all the Cells in this range
     */
    fun getCells(): Array<CellHandle>? {
        if (isDirty)
            try {
                init()
            } catch (e: CellNotFoundException) {
            }

        return cells
    }

    /**
     * Constructor which creates a new CellRange from a String range <br></br>
     * The String range must be in the format Sheet!CR:CR <br></br>
     * For Example, "Sheet1!C9:I19"
     *
     * @param String   range - the range string
     * @param WorkBook bk
     * @param boolean  createblanks - true if blank cells should be created if necessary
     * @param boolean  initcells - true if cells should (be initialized)
     */
    constructor(range: String, bk: io.starter.OpenXLS.WorkBook?, createblanks: Boolean, initcells: Boolean) {
        createBlanks = createblanks
        initializeCells = initcells
        this.range = range
        if (bk == null)
            return
        this.workBook = bk
        try {
            this.init()
        } catch (e: CellNotFoundException) {
        }

    }

    /**
     * Constructor which creates a new CellRange from a String range <br></br>
     * The String range must be in the format Sheet!CR:CR <br></br>
     * For Example, "Sheet1!C9:I19"
     *
     * @param String   range - the range string
     * @param WorkBook bk
     * @param boolean  createblanks - true if blank cells should be created (if
     * necessary)
     */
    @JvmOverloads
    constructor(range: String, bk: io.starter.OpenXLS.WorkBook?, c: Boolean = true) {
        createBlanks = c
        this.range = range
        if (bk == null)
            return
        this.workBook = bk
        try {
            if ("" != this.range)
                this.init()
        } catch (e: CellNotFoundException) {
        } catch (ne: NumberFormatException) {
            // happens upon !REF range
        }

    }

    /**
     * sets the parent of this Cell range <br></br>
     * Used Internally. Not intended for the End User.
     *
     * @param b
     */
    fun setParent(b: BiffRec) {
        parent = b
    }

    /**
     * Re-sort all cells in this cell range according to the column.
     *
     *
     * A custom comparator can be passed in, or the default one can be used with
     * sort(String, boolean).
     *
     *
     * Comparators will be passed 2 CellHandle objects for comparison.
     *
     *
     * Collections.reverse will be called on the results if ascending is set to
     * false;
     *
     * @param rowNumber  the 0 based (row 5 = 4) number of the row to be sorted upon
     * @param comparator
     * @throws RowNotFoundException
     * @throws ColumnNotFoundException
     */
    @Throws(RowNotFoundException::class)
    fun sort(rownumber: Int, comparator: Comparator<CellHandle>, ascending: Boolean) {
        this.createBlanks()
        val sortRow = this.getCellsByRow(rownumber)
        Collections.sort(sortRow, comparator)
        if (!ascending)
            Collections.reverse(sortRow)
        // now we have sorted the array list, come up with a map to resort the rows.
        var coords: IntArray? = null
        try {
            coords = this.rangeCoords
            // fix stupid wrong offsets;
            coords[0] = coords[0]--
            coords[2] = coords[2]--
        } catch (e1: CellNotFoundException) {
        }

        val outputCols = ArrayList<ArrayList<*>>()
        for (i in sortRow.indices) {
            val cell = sortRow[i]
            var cells: ArrayList<*>? = null
            try {
                cells = this.getCellsByCol(ExcelTools.getAlphaVal(cell.colNum))
            } catch (e: ColumnNotFoundException) {
                // if there are no cells in this column ignore it
            }

            outputCols.add(cells)
        }
        this.removeCells()
        for (i in coords!![1]..coords[3]) {
            val cells = outputCols[i - coords[1]]
            for (x in cells.indices) {
                val cell = cells[x] as CellHandle
                val bs = this.getSheet()!!.boundsheet
                cell.cell!!.setCol(i.toShort())
                bs!!.addCell(cell.cell as CellRec)
            }
        }

    }

    /**
     * Changes the cellRange to a createBlanks cellrange and re-initializes the
     * range, creating the missing blanks.
     */
    private fun createBlanks() {
        if (!this.createBlanks) {
            this.createBlanks = true
            this.initializeCells = true
            try {
                this.init()
            } catch (e: CellNotFoundException) {
            }

        }
    }

    /**
     * Resort all cells in the range according to the rownumber passed in.
     *
     * @param rownumber the 0 based row number
     * @param ascending
     * @throws RowNotFoundException
     * @throws ColumnNotFoundException
     */
    @Throws(RowNotFoundException::class)
    fun sort(rownumber: Int, ascending: Boolean) {
        val cp = CellComparator()
        this.sort(rownumber, cp, ascending)
    }

    /**
     * Re-sort all cells in this cell range according to the column.
     *
     *
     * A custom comparator can be passed in, or the default one can be used with
     * sort(String, boolean).
     *
     *
     * Comparators will be passed 2 CellHandle objects for comparison.
     *
     *
     * * Collections.reverse will be called on the results if ascending is set to
     * false;
     *
     * @param columnName
     * @param comparator
     * @throws RowNotFoundException
     * @throws ColumnNotFoundException
     */
    @Throws(ColumnNotFoundException::class)
    fun sort(columnName: String, comparator: Comparator<*>, ascending: Boolean) {
        if (!this.createBlanks) {
            // we cannot have empty cells in this operation
            this.createBlanks = true
            try {
                this.init()
            } catch (e: CellNotFoundException) {
            }

        }
        val sortCol = this.getCellsByCol(columnName)
        Collections.sort<Any>(sortCol, comparator)
        if (!ascending)
            Collections.reverse(sortCol)
        // now we have sorted the array list, come up with a map to resort the rows.
        var coords: IntArray? = null
        try {
            coords = this.rangeCoords
            // fix stupid wrong offsets;
            coords[0] = coords[0]--
            coords[2] = coords[2]--
        } catch (e1: CellNotFoundException) {
            e1.printStackTrace()
        }

        val outputRows = ArrayList<ArrayList<CellHandle>>()
        for (i in sortCol.indices) {
            var cells: ArrayList<CellHandle>? = null
            try {
                cells = this.getCellsByRow(sortCol[i].rowNum)
            } catch (e: RowNotFoundException) {
                // ignore if no cells available
            }

            outputRows.add(cells)
        }
        this.removeCells()
        for (i in coords!![0]..coords[2]) {
            val cells = outputRows[i - coords[0]]
            for (x in cells.indices) {
                val cell = cells[x]
                val bs = this.getSheet()!!.boundsheet
                cell.cell!!.rowNumber = i - 1
                bs!!.addCell(cell.cell as CellRec)
            }
        }
    }

    /**
     * Resort all cells in the range according to the column passed in.
     *
     * @param columnName
     * @param ascending
     * @throws ColumnNotFoundException
     * @throws RowNotFoundException
     */
    @Throws(ColumnNotFoundException::class)
    fun sort(columnName: String, ascending: Boolean) {
        val cp = CellComparator()
        this.sort(columnName, cp, ascending)
    }

    /**
     * gets the array of Cells in this Name
     *
     *
     * NOTE: this method variation also returns the Sheetname for the name record if
     * not null.
     *
     *
     * Thus this method is limited to use with 2D ranges.
     *
     * @param fragment whether to enclose result in NameHandle tag
     * @return Cell[] all Cells defined in this Name
     */
    fun getCellRangeXML(fragment: Boolean): String {
        val sbx = StringBuffer()
        if (!fragment)
            sbx.append("<?xml version=\"1\" encoding=\"utf-8\"?>")
        sbx.append("<CellRange Range=\"" + this.getRange() + "\">")
        sbx.append(xml)
        sbx.append("</NameHandle>")
        return sbx.toString()
    }

    /**
     * Constructor which creates a new CellRange using an array of cells as it's
     * constructor. <br></br>
     * NOTE that if the array of cells you are adding is not a rectangle of data (ie
     * [A1][B1][C1]) that you will have null cells in your cell range and operations
     * on it may cause errors. <br></br>
     * If you wish to populate a cell range that is not contiguous, consider the
     * constructor CellRange(CellHandle[] newcells, boolean createblanks), which
     * will populate null cells with blank records, allowing normal operations such
     * as formatting, merging, etc.
     *
     * @param CellHandle[] newcells - the array of cells from which to create the new
     * CellRange
     * @throws CellNotFoundException
     */
    @Throws(CellNotFoundException::class)
    constructor(newcells: Array<CellHandle>) {
        this.workBook = newcells[0].workBook
        this.sheet = newcells[0].workSheetHandle
        for (x in newcells.indices) {
            this.addCellToRange(newcells[x])
        }
        this.init()
    }

    /**
     * create a new CellRange using an array of cells as it's constructor. <br></br>
     * If you wish to populate a cell range that is not contiguous, set createblanks
     * to true, which will populate null cells with blank records, allowing normal
     * operations such as formatting, merging, etc.
     *
     * @param CellHandle[] newcells - the array of cells from which to create the new
     * CellRange
     * @param boolean      createblanks - true if should create blank cells if necesary
     */
    @Throws(CellNotFoundException::class)
    constructor(newcells: Array<CellHandle>, createblanks: Boolean) {
        this.createBlanks = createblanks
        this.workBook = newcells[0].workBook
        this.sheet = newcells[0].workSheetHandle
        for (x in newcells.indices) {
            this.addCellToRange(newcells[x])
        }
        this.init()
    }

    /**
     * Returns the WorkSheet referenced in this CellRange.
     *
     * @return WorkSheetHandle sheet referenced in this CellRange.
     */
    fun getSheet(): WorkSheetHandle? {
        return sheet
    }

    /**
     * initializes this CellRange
     *
     * @throws CellNotFoundException
     */
    @Throws(CellNotFoundException::class)
    fun init() {
        if (!FormulaParser.isComplexRange(range)) {
            val coords = this.rangeCoords
            var rowctr = coords[0]
            val firstcellcol = coords[1]
            val lastcellcol = coords[3]
            val numcells = coords[4]

            var cellctr = firstcellcol - 1
            try {
                if (sheetname != null) {
                    if (sheetname == "")
                    // 20080214 KSC - is this a good idea?
                    // default to work sheet 0
                        sheetname = this.workBook!!.getWorkSheet(0).sheetName
                }
                if (sheetname == null) {
                    if (this.sheet != null)
                        sheetname = this.sheet!!.sheetName
                    else
                        throw IllegalArgumentException("sheet name not specified: " + range!!)
                }

                var s = sheetname
                if (s != null) {
                    // handle enclosing apostrophes which are added to PtgRefs
                    if (s[0] == '\'') {
                        s = s.substring(1)
                        if (s[s.length - 1] == '\'') {
                            s = s.substring(0, s.length - 1)
                        }
                    }
                }
                sheet = workBook!!.getWorkSheet(s)
                // if wholerow or wholecol, don't gather cells
                if (this.wholeCol || this.wholeRow)
                    return
                cells = arrayOfNulls(numcells)
                var resetFastAdds = false
                if (sheet!!.fastCellAdds && this.createBlanks) {
                    resetFastAdds = true
                    sheet!!.fastCellAdds = false
                }
                for (i in 0 until numcells) {
                    if (cellctr == lastcellcol) {// if its the end of the row,
                        // increment row.
                        cellctr = firstcellcol - 1
                        rowctr++
                    }
                    ++cellctr
                    try {
                        // use caching 20080917 KSC: PROBLEM HERE [Claritas
                        // BugTracker 1862]
                        if (this.initializeCells)
                            cells[i] = sheet!!.getCell(rowctr - 1, cellctr, sheet!!.useCache) // 20080917 KSC: use
                        // cache
                        // setting instead of
                        // defaulting to true);
                    } catch (e: CellNotFoundException) {
                        if (this.createBlanks) {
                            cells[i] = sheet!!.add(null, rowctr - 1, cellctr)
                        }

                    }

                }
                if (resetFastAdds) {
                    sheet!!.fastCellAdds = true
                }
            } catch (e: WorkSheetNotFoundException) {
                throw IllegalArgumentException(e.toString())
            }

        } else { // gather cells for a complex range ...
            val pm = io.starter.formats.XLS.formulas.PtgMemFunc()
            val b = XLSRecord()
            b.workBook = this.workBook!!.workBook
            pm.parentRec = b
            try {
                pm.location = range
                val p = pm.components
                val cellsfromcomplexrange = java.util.ArrayList<CellHandle>()
                for (i in p!!.indices) {
                    try {
                        cellsfromcomplexrange.add(workBook!!.getCell((p[i] as PtgRef).location))
                    } catch (e: CellNotFoundException) {
                        if (this.createBlanks) {
                            cells[i] = sheet!!.add(null, p[i].location)
                        }
                    }

                }
                cells = arrayOfNulls(cellsfromcomplexrange.size)
                cells = cellsfromcomplexrange.toTypedArray<CellHandle>()
            } catch (e: Exception) {
                throw IllegalArgumentException(e.toString())
            }

        }
        isDirty = false
    }

    /**
     * Initializes this `CellRange`'s cell list if necessary. This method
     * is useful if this `CellRange` was created with
     * `initCells` set to `false` and it is later necessary to
     * retrieve the cell list.
     *
     * @param createBlanks whether missing cells should be created as blanks. If this is
     * `false` they will appear in the cell list as
     * `null`s.
     */
    fun initCells(createBlanks: Boolean) {
        // If we don't need to do anything, return
        if (initializeCells == true && (!createBlanks || this.createBlanks))
            return

        this.initializeCells = true
        this.createBlanks = createBlanks

        try {
            this.init()
        } catch (e: CellNotFoundException) {
            // This will never actually happen but we have to catch it anyway
            throw Error()
        }

    }

    /**
     * Return the String representation of this range
     *
     * @return the String range
     */
    fun getRange(): String? {
        return range
    }

    /**
     * Sets the range of cells for this CellRange to a string range
     *
     * @param String rng - Range string
     */
    fun setRange(rng: String) {
        range = rng
        try {
            this.init()
        } catch (e: CellNotFoundException) {
            // don't have to report anything
        }

    }

    /**
     * sets a border around the range of cells
     *
     * @param int            width - line width
     * @param int            linestyle - line style
     * @param java.awt.Color colr - color of border line
     */
    fun setBorder(width: Int, linestyle: Int, colr: java.awt.Color) {
        val ch = getCells()
        for (t in ch!!.indices) {
            val coords = getEdgePositions(ch[t], width)
            // create Excel border -- top, left, bottom, right
            if (coords[0] > 0) {
                ch[t].setTopBorderLineStyle(linestyle.toShort())
                ch[t].setBorderTopColor(colr)
            }
            if (coords[1] > 0) {
                ch[t].setLeftBorderLineStyle(linestyle.toShort())
                ch[t].setBorderLeftColor(colr)
            }
            if (coords[2] > 0) {
                ch[t].setBottomBorderLineStyle(linestyle.toShort())
                ch[t].setBorderBottomColor(colr)
            }
            if (coords[3] > 0) {
                ch[t].setRightBorderLineStyle(linestyle.toShort())
                ch[t].setBorderRightColor(colr)
            }
        }
    }

    /**
     * update the CellRange when the underlying Cells change their location
     *
     * @return boolean true if the CellRange could be updated, false if there are no
     * cells represented by this range
     */
    fun update(): Boolean {
        if (cells == null)
            return false // this is an invalid range -- see Mergedcells prob.
        // jm

        // arbitrarily set the initial vals...
        if (cells!![0] != null) { // 20100106 KSC: if didn't create blanks it's possible that cells are null
            firstcellrow = cells!![0].rowNum + 1
            firstcellcol = cells!![0].colNum
            lastcellrow = cells!![0].rowNum + 1
            lastcellcol = cells!![0].colNum
            for (t in cells!!.indices) {
                val cx = cells!![t]
                // 20090901 KSC: apparently can be null
                if (cx != null)
                    this.addCellToRange(cx)
            }
            this.myrowints = null
            this.mycolints = null
            return true
        } else if (this.range != null) {// 20100106 KSC: handle ranges containing null cell[0] (i.e. ranges referencing
            // cells not present and createBlanks==false)
            if (this.DEBUG)
                Logger.logWarn("CellRange.update:  trying to access blank cells in range " + this.range!!)
            try {
                this.rangeCoords
                return true
            } catch (e: CellNotFoundException) { // shouldn't
                return false
            }

        }
        return false // return false if it doesn't have it's cells defined
    }

    /**
     * Sets the sheet reference for this CellRange.
     *
     * @param WorkSheetHandle aSheet
     */
    fun setSheet(aSheet: WorkSheetHandle) {
        this.sheet = aSheet
        this.sheetname = aSheet.sheetName
    }

    /**
     * Copies this range to another location. At present only contents and complete
     * formats may be copied.
     *
     * @param row  the topmost row of the target area
     * @param col  the leftmost column of the target area
     * @param what a set of flags determining what will be copied
     * @return the destination range
     */
    fun copy(sheet: WorkSheetHandle, row: Int, col: Int, what: Int): CellRange {
        var row = row
        var col = col
        val result = CellRange(sheet, row, col, this.width, this.height)

        val first_col = col

        // note these are not currently used, see setting below
        val copy_contents = what and COPY_CONTENTS != 0
        var copy_formulas = what and COPY_FORMULAS != 0
        var copy_formats = what and COPY_FORMATS != 0

        // set to true until this thing is fully implemented
        copy_formats = true
        copy_formulas = true

        var cur_row = cells!![0].rowNum
        for (idx in cells!!.indices) {
            val source = cells!![idx]

            if (source.rowNum != cur_row) {
                cur_row = source.rowNum
                row++
                col = first_col
            }

            var target: CellHandle? = null
            try {
                target = sheet.getCell(row, col)
            } catch (e: CellNotFoundException) {
            }

            val formatID: Int
            if (copy_formats)
                formatID = source.formatId
            else if (target != null)
                formatID = target.formatId
            else
                formatID = sheet.workBook!!.workBook!!.defaultIxfe

            if (copy_contents) {
                val value: Any?

                if (copy_formulas && source.isFormula) {
                    try {
                        value = source.formulaHandle.formulaString
                    } catch (e: FormulaNotFoundException) {
                        // This shouldn't happen; we known the formula exists.
                        // If it does happen it indicates a bug in OpenXLS,
                        // thus we throw an Error.
                        throw Error("formula cell has no Formula record", e)
                    }

                } else {
                    value = source.`val`
                }

                target = sheet.add(value, row, col, formatID)

                if (target!!.isFormula)
                    try {
                        FormulaHandle.moveCellRefs(target.formulaHandle,
                                intArrayOf(row - source.rowNum, col - source.colNum))
                    } catch (e: FormulaNotFoundException) {
                    }

            } else if (target == null) {
                target = sheet.add(null, row, col, formatID)
            }

            if (copy_formats) {
                target!!.formatId = formatID
            }

            result.cells[idx] = target
            col++
        }

        return result
    }

    /**
     * Fills this range from the given cell.
     *
     * @param source    the cell whose attributes should be copied or `null` to
     * copy from the first cell in the range
     * @param what      a set of flags determining what will be copied
     * @param increment the amount by which to increment numeric values or
     * `NaN` for no increment
     */
    fun fill(source: CellHandle?, what: Int, increment: Double) {
        var source = source
        if (null == source)
            source = cells!![0]

        val copy_contents = what and COPY_CONTENTS != 0
        val copy_formulas = what and COPY_FORMULAS != 0
        val copy_formats = what and COPY_FORMATS != 0

        val sourceRow = source.rowNum
        val sourceCol = source.colNum

        var value: Any? = null
        if (copy_contents) {
            if (copy_formulas && source.isFormula) {
                try {
                    value = source.formulaHandle.formulaString
                } catch (e: FormulaNotFoundException) {
                    throw Error("formula cell has no Formula record", e)
                }

            } else {
                value = source.`val`
            }
        }

        // if increment is set, ensure the value can be incremented
        if (!java.lang.Double.isNaN(increment) && !(copy_contents && value is Number))
            throw IllegalArgumentException("cannot increment unless filling with a numeric value")

        for (idx in cells!!.indices) {
            var target: CellHandle? = cells!![idx]

            // don't overwrite the source cell
            if (source == target)
                continue

            if (!java.lang.Double.isNaN(increment))
                value = (value as Number).toDouble() + increment

            val formatID = (if (copy_formats) source else target).formatId

            if (copy_contents) {
                target = sheet!!.add(value, target!!.rowNum, target.colNum, formatID)
                cells[idx] = target

                if (target!!.isFormula)
                    try {
                        FormulaHandle.moveCellRefs(target.formulaHandle,
                                intArrayOf(target.rowNum - sourceRow, target.colNum - sourceCol))
                    } catch (e: FormulaNotFoundException) {
                    }

            }

            // when adding Date values passing the format ID to sheet.add
            // doesn't set the format so we always set it here
            if (copy_formats) {
                target!!.formatId = formatID
            }

        }
    }

    fun calculateAffectedCellsOnSheet(): Collection<CellHandle> {
        val affected = HashSet<CellHandle>()
        for (cell in cells!!) {
            if (cell != null) {
                affected.add(cell)
                affected.addAll(cell.calculateAffectedCellsOnSheet())
            }
        }
        return affected
    }

    /**
     * Get the cells from a particular rownumber, constrained by the boundaries of
     * the cellRange
     *
     * @param rownumber
     */
    @Throws(RowNotFoundException::class)
    fun getCellsByRow(rownumber: Int): ArrayList<CellHandle> {
        val al = ArrayList<CellHandle>()
        val r = this.getSheet()!!.getRow(rownumber)
        val cells = r.cells
        var coords: IntArray? = null
        try {
            coords = this.rangeCoords
        } catch (e: CellNotFoundException) {
            throw RowNotFoundException("Error getting internal coordinates for CellRange$e")
        }

        for (i in cells.indices) {
            if (cells[i].colNum >= coords[1] && cells[i].colNum <= coords[3]) {
                al.add(cells[i])
            }
        }
        return al
    }

    /**
     * Get the cells from a particular column, constrained by the boundaries of the
     * cellRange
     *
     * @param rownumber
     * @throws ColumnNotFoundException
     */
    @Throws(ColumnNotFoundException::class)
    fun getCellsByCol(col: String): ArrayList<CellHandle> {
        val al = ArrayList<CellHandle>()
        val r = this.getSheet()!!.getCol(col)
        val cells = r.cells
        var coords: IntArray? = null
        try {
            coords = this.rangeCoords
            coords[0] = coords[0] - 1
            coords[2] = coords[2] - 1
        } catch (e: CellNotFoundException) {
            throw ColumnNotFoundException("Error getting internal coordinates for CellRange$e")
        }

        for (i in cells.indices) {
            if (cells[i].rowNum >= coords[0] && cells[i].rowNum <= coords[2]) {
                al.add(cells[i])
            }
        }
        return al
    }

    /**
     * removes the border from all of the cells in this range
     */
    fun removeBorder() {
        val ch = getCells()
        for (t in ch!!.indices) {
            ch[t].removeBorder()
        }
    }

    /**
     * Sets a bottom border on all cells in the cellrange
     *
     *
     * Linestyle should be set through the FormatHandle constants
     */
    fun setInnerBorderBottom(linestyle: Int, colr: java.awt.Color) {
        val ch = getCells()
        for (t in ch!!.indices) {
            ch[t].setBottomBorderLineStyle(linestyle.toShort())
            ch[t].setBorderBottomColor(colr)
        }
    }

    /**
     * Sets a right border on all cells in the cellrange
     *
     *
     * Linestyle should be set through the FormatHandle constants
     */
    fun setInnerBorderRight(linestyle: Int, colr: java.awt.Color) {
        val ch = getCells()
        for (t in ch!!.indices) {
            ch[t].setRightBorderLineStyle(linestyle.toShort())
            ch[t].setBorderRightColor(colr)
        }
    }

    /**
     * Sets a left border on all cells in the cellrange
     *
     *
     * Linestyle should be set through the FormatHandle constants
     */
    fun setInnerBorderLeft(linestyle: Int, colr: java.awt.Color) {
        val ch = getCells()
        for (t in ch!!.indices) {
            ch[t].setLeftBorderLineStyle(linestyle.toShort())
            ch[t].setBorderLeftColor(colr)
        }
    }

    /**
     * Sets a top border on all cells in the cellrange
     *
     *
     * Linestyle should be set through the FormatHandle constants
     */
    fun setInnerBorderTop(linestyle: Int, colr: java.awt.Color) {
        val ch = getCells()
        for (t in ch!!.indices) {
            ch[t].setTopBorderLineStyle(linestyle.toShort())
            ch[t].setBorderTopColor(colr)
        }
    }

    /**
     * Sets a surround border on all cells in the cellrange
     *
     *
     * Linestyle should be set through the FormatHandle constants
     */
    fun setInnerBorderSurround(linestyle: Int, colr: java.awt.Color) {
        val ch = getCells()
        for (t in ch!!.indices) {
            ch[t].setBorderColor(colr)
            ch[t].setBorderLineStyle(linestyle.toShort())
        }
    }

    companion object {

        /**
         *
         */
        private const val serialVersionUID = -3609881364824289079L
        val REMOVE_MERGED_CELLS = true
        val RETAIN_MERGED_CELLS = false

        var xmlResponsePre = "<CellRange>"
        var xmlResponsePost = "</CellRange>"

        /**
         * Whether to copy the cell contents.
         */
        val COPY_CONTENTS = 0x01

        /**
         * Whether formulas should be copied. If this bit is not set the formula result
         * will be copied instead.
         */
        val COPY_FORMULAS = 0x02

        val COPY_FORMATS = 0x0100

        /**
         * return a JSON array of cell values for the given range <br></br>
         * static version
         *
         * @param String   range - a string representation of the desired range of cells
         * @param WorkBook wbh - the source WorkBook for the cell range
         * @return JSONArray - a JSON representation of the desired cell range
         */
        fun getValuesAsJSON(range: String, wbh: WorkBook): JSONArray {
            val rangeArray = JSONArray()
            try {
                val cr = CellRange(range, wbh, true)
                for (j in 0 until cr.getCells()!!.size)
                    rangeArray.put(cr.getCells()!![j].`val`)
            } catch (e: Exception) {
                Logger.logErr("Error obtaining CellRange $range JSON: $e")
            }

            return rangeArray
        }

        /**
         * returns the cells for a given range <br></br>
         * static version
         *
         * @param String   range - a string representation of the desired range of cells
         * @param WorkBook wbh - the source WorkBook for the cell range
         * @return CellHandle[] array of cells represented by the desired cell range
         */
        fun getCells(range: String, wbh: WorkBookHandle): Array<CellHandle>? {
            val cr = CellRange(range, wbh, true)
            return cr.getCells()
        }
    }

}
/**
 * Initializes a `CellRange` from a `CellRangeRef`. The
 * source `CellRangeRef` instance must be qualified with a single
 * resolved worksheet.
 *
 * @param source the `CellRangeRef` from which to initialize
 * @throws IllegalArgumentException if the source range does not have a resolved sheet or has more
 * than one sheet
`` */
/**
 * Constructor to create a new CellRange from a WorkSheetHandle and a set of
 * range coordinates: <br></br>
 * coords[0] = first row <br></br>
 * coords[1] = first col <br></br>
 * coords[2] = last row <br></br>
 * coords[3] = last col
 *
 * @param WorkSheetHandle sht - handle to the WorkSheet containing the Range's Cells
 * @param int[]           coords - the cell coordinates
 */
/**
 * Constructor which creates a new CellRange from a String range <br></br>
 * The String range must be in the format Sheet!CR:CR <br></br>
 * For Example, "Sheet1!C9:I19"
 *
 * @param String   range - the range string
 * @param WorkBook bk
 * @throws CellNotFoundException
 */
