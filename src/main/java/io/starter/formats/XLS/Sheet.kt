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
/*
 * Created on Dec 16, 2004
 *
 * To change the template for this generated file go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
package io.starter.formats.XLS

import io.starter.formats.XLS.charts.Chart

import java.io.IOException

/**
 *
 */
interface Sheet : BiffRec {

    // returns the records belonging to the sheet
    val sheetRecs: List<*>

    val myEof: Eof

    val header: Headerrec

    val footer: Footerrec

    /**
     * get the last BiffRec added to this sheet
     */
    val lastCell: BiffRec

    /**/
    val localRecs: List<*>

    val myBof: Bof

    var window2: Window2

    /**
     * remove rec from the vector, includes firing
     * a changeevent.
     */

    val isChartOnlySheet: Boolean

    /**
     * return the pos of the Bof for this Sheet
     */
    /**
     * set the pos of the Bof for this Sheet
     */
    var lbPlyPos: Long

    /**
     * the beginning of the Dimensions record
     * is the index of the RowBlocks
     */
    var dimensions: Dimensions

    val minRow: Int

    val maxRow: Int

    val minCol: Int

    val maxCol: Int

    /**
     * set the numeric sheet number
     */
    val sheetNum: Int
    //public abstract void shiftCellRow(BiffRec c, int origRownum, int shiftamount, int flag);
    /* we need to handle shifting any cell references
     * ie: moving formula C12=SUM(C10:C11) to C13
     *	should update the formula to C13=SUM(C11:C12)
     *
     *	we only do this if the formula references are 'unlocked'
     *
     *	additional restrictions:
     *		Invoice!F19:F19 - NO, single BiffRec Range
     *		Sheet2!C3 - NO, not on the same sheet as shifting Cells
     *		Invoice!F20:F21 - YES, a real range
     *		Invoice!F20 - YES, a real single ref
     *
     */
    /* NOTE:  MERGED INTO SHEET.SHIFTCELLROW AS ONE CANNOT CHANGE CELL ROW WITHOUT ALL THE CHECKING MACHINERY CONTAINED WITHIN
	public abstract void setCellRow(BiffRec c, String oldaddr, int newrow);
*/

    /**
     * get whether this sheet is hidden upon opening
     */
    val hidden: Boolean

    /**
     * get the number of defined rows on this sheet
     */
    val numRows: Int

    /**
     * get the number of defined cells on this sheet
     */
    val numCells: Int

    /**
     * get the List of columns defined on this sheet
     */
    val colNames: List<*>

    /**
     * get the Number of columns defined on this sheet
     */
    val numCols: Int

    /**
     * get the List of rows defined on this sheet
     */
    val rowNums: List<*>

    /**
     * return an Array of the Rows in no particular order
     */
    val rows: Array<Row>

    /**
     * get the Collection of Colinfos
     */
    val colinfos: Collection<Colinfo>
    /** get an array of all cells for this worksheet
     *
     * public abstract Hashtable getCellMap(); */
    /**
     * get an array of all cells for this worksheet
     */
    val cells: Array<BiffRec>

    val mergedCellsRecs: List<*>

    val mergedCellsRec: Mergedcells

    /**
     * get the name of the sheet
     */
    /**
     * change the displayed name of the sheet
     *
     *
     * Affects the following byte values:
     * 10      cch         1       Length of sheet name
     * 11      grbitChr    1       Compressed/Uncompressed Unicode
     * 12      rgch        var     Sheet name
     */
    var sheetName: String

    /**
     * @param b
     */
    var grbitChr: Byte

    /*
        Returns a serialized copy of this SheetImpl
    */
    val sheetBytes: ByteArray

    /**
     * get the type of sheet as a short
     */
    val sheetType: Short

    /**
     * get the type of sheet as a string
     */
    val sheetTypeString: String

    var guts: Guts

    var wsBool: WsBool

    fun setHeader(hr: BiffRec)

    fun setFooter(r: BiffRec)

    /**
     * do all of the expensive updating here
     * only right before streaming record.
     */
    override fun preStream()

    /**
     * Remove a BiffRec from this WorkSheet.
     */
    fun removeCell(celladdr: String)

    /**
     * Remove a BiffRec from this WorkSheet.
     */
    fun removeCell(c: BiffRec)

    fun removeRecFromVec(rec: BiffRec)

    /**
     * Called from removeCell(), removeMulrk() handles the fact that you
     * are trying to delete a rk that is really just a part of a Mulrk.  This
     * is handled by truncating the mulrk at the cell, then creating individual numbers
     * after the deleted cell.
     */
    fun removeMulrk(thisrk: Rk)

    /**
     * get a handle to a specific column of cells in this sheet
     */
    fun getColInfo(colin: Int): Colinfo

    /**
     * remove all Sheet records from Sheet.
     */
    fun removeAllRecords()

    /**
     * set the Bof record for this SheetImpl
     */
    fun setBOF(b: Bof)

    fun setEOF(f: Eof)

    /**
     * update the INDEX record with the new max Row #
     * why we need so many redundant references to the Min/Max Row/Cols
     * is a question for the Redmond sages.
     */
    fun updateDimensions(row: Int, c: Int)

    /**
     * set the associated sheet index
     */
    fun setSheetIDX(idx: Index)


    /** add a row to the worksheet as well
     * as to the RowBlock which will handle
     * the updating of Dbcell index behavior
     *
     *
     * @param BiffRec the cell being added (can't add a row without one...)
     *
     * public abstract Row addNewRow(BiffRec cell);
     */

    /**
     * inserts a row and shifts all of the other rows down by the number
     * of rows specified
     */
    //public abstract void insertRow(int rownum, int numrows);
    /**
     * inserts a row and shifts all of the other rows down one
     */
    //	public abstract Row insertRow(int rownum, boolean shiftrows);

    /**
     * Shifts a row down
     *
     * @param shiftamount - number of rows shifted
     */
    //public abstract void shiftRow(Row roe, int shiftamount)
    //	throws RowNotFoundException;

    //	deal with shifting Merged cells. 10-15-04 -jm
    fun updateMergedCells()

    /**
     * set whether this sheet is hidden upon opening
     */
    fun setHidden(gr: Int)

    /**
     * returns the selected sheet status
     */
    fun selected(): Boolean

    /**
     * set whether this sheet is selected upon opening
     */
    fun setSelected(b: Boolean)

    /**
     * get a handle to the Row at the specified
     * row index
     */
    fun getRowByNumber(r: Int): Row
    /** Add a Value record to a WorkSheet.
     *
     *
     * @param obj the value of the new Cell
     * @param row & col address of the new Cell
     */
    //public abstract BiffRec add(Object obj, int[] rc);

    /**
     * Add a Value record to a WorkSheet.
     *
     * @param obj     the value of the new Cell
     * @param address the address of the new Cell
     */
    fun addValue(obj: Any, address: String): BiffRec

    /**
     * Add an BiffRec to a WorkSheet
     */
    fun addRecord(rec: BiffRec, rc: IntArray)

    fun setCopyPriorCellFormats(f: Boolean)
    //	public abstract boolean copyPriorCellFormatForNewCells(BiffRec c);
    /** add a new cell to the book
     */
    //public abstract void addNewCell(BiffRec cell);

    /**
     * add a cell to the worksheet cell collection
     *
     *
     * associates with Row and Col
     */
    fun addCell(cell: CellRec)
    /*
		  Returns the *real* last col num.  Unfortunately the dimensions record
		  cannot be counted on to give a correct value.
	  **/
    //public abstract int getRealMaxCol();

    /**
     * column formatting records
     */
    fun addColinfo(c: Colinfo)

    /**
     * get  a colinfo by first col in range
     */
    fun getColinfo(c: String): Colinfo
    /** add a new rec to the book where rec already has a cell
     * or does not need one.
     *
     * three things need to happen
     * 1. get the record index in the recvec
     * 2. get the record offset in the workbook stream
     * 3. put the record in the output vector
     *
     */
    //public abstract void addNewRecord(BiffRec vr);
    /**
     * Moves a cell location from one address to another
     */

    /**
     * get a cell by address from the worksheet
     */
    fun getCell(s: String): BiffRec

    /**
     * Returns a BiffRec for working with
     * the value of a BiffRec on a WorkSheet.
     *
     * @param int Row the integer row of the Cell
     * @param int Col the integer col of the Cell
     * @throws CellNotFoundException is thrown if there is no existing BiffRec at the specified address.
     */
    @Throws(CellNotFoundException::class)
    fun getCell(row: Int, col: Int): BiffRec

    fun addMergedCellsRec(r: Mergedcells)

    /**
     * get the name of the sheet
     */
    override fun toString(): String

    /**
     * initialize the SheetImpl with data from
     * the byte array.
     */
    override fun init()

    /* prior to serializing the worksheet,
       we need to initialize the records which belong to this sheet
       instance.

    */
    fun setLocalRecs()

    /* Inserts a serialized boundsheet into the workboook
     */
    fun addChart(inbytes: ByteArray, NewChartName: String, coords: ShortArray): Chart

    companion object {

        // hidden states from grbit field offset 1
        val VISIBLE: Byte = 0x00

        val HIDDEN: Byte = 0x01

        val VERY_HIDDEN: Byte = 0x02

        // sheet types from grbit field offset 0
        val TYPE_SHEET_DIALOG: Byte = 0x00

        val TYPE_XL4_MACRO: Byte = 0x01

        val TYPE_CHART: Byte = 0x02

        val TYPE_VBMODULE: Byte = 0x06
    }
    //	----- TODO: implement COMBINATOR SECTION ------ //
    //public abstract void initCombinators(SheetImpl sh);
}