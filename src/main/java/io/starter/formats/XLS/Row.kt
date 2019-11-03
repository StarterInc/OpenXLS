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

import io.starter.OpenXLS.FormatHandle
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList

/**
 * **ROW 0x208: Describes a single row on a MS Excel Sheet.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---------------------------------------------------------------
 * 0       rw          2       Row number
 * 2       colMic      2       First defined column in the row
 * 4       colMac      2       Last defined column in the row plus 1
 * 6       miyRw      	2       Row Height in twips (1/20 of a printer's point, or 1/1440 of an inch)
 * 14-0 7FFFH Height of the row, in twips
 * 15 8000H 0 = Row has custom height; 1 = Row has default height
 * 8      irwMac      	2       Optimizing, set to 0
 * 10      reserved    2
 * 12      grBit       2       Option Flags
 * 14      ixfe        2       Index to XF record for row
 *
 * grib options
 * offset	Bits	mask	name					contents
 * ---------------------------------------------------------------
 * 2-0 	07H 	outlineLevel			Outline level of the row
 * 4 		10H 	fCollapsed 				1 = Outline group starts or ends here (depending on where
 * the outline buttons are located, see WSBOOL record and is fCollapsed
 * 5 		20H 	fHidden					1 = Row is fHidden (manually, or by a filter or outline group)
 * 6 		40H 	altered					1 = Row height and default font height do not match
 * 7 		80H 	fFormatted	1 = Row has explicit default format (fl)
 * 8 		100H 							Always 1
 * 27-16 	0FFF0000H 							If fl = 1: Index to default XF record
 * 28 	10000000H 							1 = Additional space above the row. This flag is set, if the
 * upper border of at least one cell in this row or if the lower
 * border of at least one cell in the row above is formatted with
 * a thick line style. Thin and medium line styles are not taken
 * into account.
 * 29  20000000H 							1 = Additional space below the row. This flag is set, if the
 * lower border of at least one cell in this row or if the upper
 * border of at least one cell in the row below is formatted with
 * a medium or thick line style. Thin line styles are not taken
 * into account.
 *
</pre> *
 *
 * @see WorkBook
 *
 * @see BOUNDSHEET
 *
 * @see INDEX
 *
 * @see Dbcell
 *
 * @see ROW
 *
 * @see Cell
 *
 * @see XLSRecord
 */

class Row : io.starter.formats.XLS.XLSRecord {

    private var colMic: Short = 0
    private var colMac: Short = 0
    private var miyRw: Int = 0

    /**
     * get the Dbcell record which contains the
     * cell offsets for this row.
     *
     *
     * needed in computing new INDEX offset values
     */
    /**
     * set the Dbcell record which contains the
     * cell offsets for this row.
     *
     *
     * needed in computing new INDEX offset values
     */
    internal var dbCell: Dbcell? = null
        set
    private var firstcell: BiffRec? = null
    private var lastcell: BiffRec? = null

    private var fCollapsed: Boolean = false
    private var fHidden: Boolean = false
    /**
     * true if row height has been altered from default
     * i.e. set manually
     *
     * @return
     */
    var isAlteredHeight: Boolean = false
        private set
    /**
     * This flag determines if the row has been formatted.
     * If this flag is not set, the XF reference will not affect the row.
     * However, if it's true then the row will be formatted according to
     * the XF ref.
     *
     * @return
     */
    var explicitFormatSet: Boolean = false
        private set
    /**
     * Additional space above the row. This flag is set, if the
     * upper border of at least one cell in this row or if the lower
     * border of at least one cell in the row above is formatted with
     * a thick line style. Thin and medium line styles are not taken
     * into account.
     */
    /**
     * flags this row to have at least one cell that has a thick top border
     *
     *
     * For internal use only
     *
     * @param hasBorder
     */
    var hasAnyThickTopBorder: Boolean = false
    /**
     * Additional space below the row. This flag is set, if the
     * lower border of at least one cell in this row or if the upper
     * border of at least one cell in the row below is formatted with
     * a medium or thick line style. Thin line styles are not taken
     * into account.
     */
    var hasAnyBottomBorder: Boolean = false
        private set
    private var fPhonetic: Boolean = false
    private var outlineLevel = 0

    /**
     * Get the height of a row
     */
    /**
     * Set the height of a row in twips (1/20th of a point)
     */
    //	15th bit set= row size is not default
    // not 100% sure of this ...
    // set bit 6 = row height and default font DO NOT MATCH
    // 10      miyRw       2       Row Height
    var rowHeight: Int
        get() {
            if (this.miyRw < 0)
                this.miyRw = (this.miyRw + 1) * -1
            return this.miyRw
        }
        set(x) {
            if (DEBUGLEVEL > 3) Logger.logInfo("Updating Row Height: " + this.rowNumber + " to: " + x)
            isAlteredHeight = true
            updateGrbit()
            val rw = ByteTools.shortToLEBytes(x.toShort())
            System.arraycopy(rw, 0, this.getData()!!, 6, 2)
            miyRw = x
        }

    /**
     * get the position of the ROW on the Worksheet
     */
    /**
     * set the position of the ROW on the Worksheet
     */
    override var rowNumber: Int
        get() {
            if (rw < 0) {
                val rowi = rw * -1
                rw = XLSConstants.MAXROWS - rowi
            }
            return rw
        }
        set(n) {
            if (DEBUGLEVEL > 3) Logger.logInfo("Updating Row Number: " + this.rowNumber + " to: " + n)
            rw = n
            val rwb = ByteTools.shortToLEBytes(rw.toShort())
            System.arraycopy(rwb, 0, this.getData()!!, 0, 2)
        }

    /**
     * Get the real max col
     */
    internal val realMaxCol: Int
        get() {
            var collast = 0
            val cs = this.cells.iterator()
            while (cs.hasNext()) {
                val c = cs.next()
                if (c.colNumber > collast) collast = c.colNumber.toInt()
            }
            return collast
        }


    internal val maxCol: Int
        get() {
            this.preStream()
            return colMac.toInt()
        }

    internal val minCol: Int
        get() {
            this.preStream()
            return colMic.toInt()
        }


    /**
     * get a collection of cells in column-based order	*
     */
    // no cells in this row
    val cells: Collection<BiffRec>
        get() {
            try {
                return this.sheet!!.getCellsByRow(this.rowNumber)
            } catch (e: CellNotFoundException) {
            }

            return ArrayList()
        }

    /**
     * get the cells as an array.  Needed when
     * operations will be used on the cell array causing concurrentModificationException
     * problems on the TreeMap collection
     *
     * @return
     */
    val cellArray: Array<Any>
        get() {
            val cells = this.cells
            var br = arrayOfNulls<Any>(cells.size)
            br = cells.toTypedArray()
            return br
        }

    val numberOfCells: Int
        get() = this.cells.size


    /**
     * iterate and get the record index of the last val record
     */
    // empty row
    val lastRecIndex: Int
        get() = if (lastcell != null) {
            lastcell!!.recordIndex
        } else
            this.recordIndex

    /**
     * return the min/max column for this row
     *
     * @return
     */
    val colDimensions: IntArray
        get() = intArrayOf(colMic.toInt(), colMac.toInt())

    /**
     * Returns whether the row is collapsed
     *
     * @return
     */
    /**
     * Set whether the row is fCollapsed
     * hides all contiguous rows with the same outline level
     *
     * @param b
     */
    // implement bit masking set on grbit
    var isCollapsed: Boolean
        get() = fCollapsed
        set(b) {
            fCollapsed = b
            fHidden = b
            var keepgoing = true
            var counter = 1
            while (keepgoing) {
                val r = this.sheet!!.getRowByNumber(this.rowNumber + counter)
                if (r != null && r.outlineLevel == this.getOutlineLevel()) {
                    r.isHidden = b
                } else {
                    keepgoing = false
                }
                counter++
            }
            counter = 1
            keepgoing = true
            while (keepgoing) {
                val r = this.sheet!!.getRowByNumber(this.rowNumber - counter)
                if (r != null && r.outlineLevel == this.getOutlineLevel()) {
                    r.isHidden = b
                } else {
                    keepgoing = false
                }
                counter++
            }
            updateGrbit()
        }

    /**
     * Returns whether the row is hidden
     * TODO:  same issue as setHidden above!
     *
     * @return
     */
    /**
     * Set whether the row is fHidden
     *
     * @param b
     */
    //		implement bit masking set on grbit
    var isHidden: Boolean
        get() = fHidden
        set(b) {
            fHidden = b
            updateGrbit()
        }

    /**
     * Gets the ID of the format currently applied to this row.
     *
     * @return the ID of the current format,
     * or the default format ID if no format has been applied
     */
    /**
     * Applies the format with the given ID to this row.
     *
     * @param ixfe the format ID. Must be between 0x0 and 0xFFF.
     * @throws IllegalArgumentException if the given format ID cannot be
     * encoded in the 1.5 byte wide field provided for it
     */
    override var ixfe: Int
        get() = if (explicitFormatSet) ixfe else this.workBook!!.defaultIxfe
        set(ixfe) {
            if (ixfe and 0xFFF.inv() != 0)
                throw IllegalArgumentException(
                        "ixfe value 0x" + Integer.toHexString(ixfe)
                                + " out of range, must be between 0x0 and 0xfff")
            this.ixfe = ixfe
            if (ixfe != this.workBook!!.defaultIxfe)
                explicitFormatSet = true
        }

    /**
     * returns true if there is a thick bottom border set on the row
     */
    /**
     * sets this row to have a thick top border
     */
    // reset
    var hasThickTopBorder: Boolean
        get() {
            if (!explicitFormatSet) return false
            if (this.hasAnyThickTopBorder) {
                try {
                    val bs = this.xfRec!!.topBorderLineStyle.toInt()
                    return bs == FormatHandle.BORDER_DOUBLE.toInt() || bs == FormatHandle.BORDER_THICK.toInt()
                } catch (e: Exception) {
                }

            }
            return false
        }
        set(hasBorder) {
            this.hasAnyThickTopBorder = hasBorder
            if (hasBorder) {
                val fh = FormatHandle(null, this.xfRec)
                fh.topBorderLineStyle = FormatHandle.BORDER_THICK
                ixfe = fh.formatId
                myxf = null
            }
            explicitFormatSet = true
        }

    /**
     * returns true if there is a thick bottom border set on the row
     */
    /**
     * sets this row to have a thick bottom border
     */
    // reset
    var hasThickBottomBorder: Boolean
        get() {
            if (!explicitFormatSet) return false
            if (this.hasAnyBottomBorder) {
                try {
                    val bs = this.xfRec!!.bottomBorderLineStyle.toInt()
                    return bs == FormatHandle.BORDER_DOUBLE.toInt() || bs == FormatHandle.BORDER_THICK.toInt()
                } catch (e: Exception) {
                }

            }
            return this.hasAnyBottomBorder
        }
        set(hasBorder) {
            this.hasAnyBottomBorder = hasBorder
            if (hasBorder) {
                val fh = FormatHandle(null, this.xfRec)
                fh.bottomBorderLineStyle = FormatHandle.BORDER_THICK
                ixfe = fh.formatId
                myxf = null
            }
            explicitFormatSet = true
        }

    constructor() : super() {}


    constructor(rowNum: Int, book: WorkBook) {
        this.workBook = book
        this.length = defaultsize.toShort()
        this.opcode = XLSConstants.ROW
        val dta = ByteArray(defaultsize)
        dta[6] = 0xff.toByte()
        dta[13] = 0x1.toByte()
        dta[14] = 0xf.toByte()
        this.setData(dta)
        this.originalsize = defaultsize
        this.init()
        this.rowNumber = rowNum
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        dbCell = null
        firstcell = null
        lastcell = null
        this.workBook = null
        this.setSheet(null)
    }


    /**
     * add a cell to the Row.  Instead of using the full
     * cell address as the treemap identifier, just use the column.
     * this allows the natural ordering of the treemap to work to our
     * advantage on output, ordering cells from lowest to highest col.
     */
    internal fun addCell(c: BiffRec) {
        c.row = this

        /*
		 * I'm not clear on this operation, it seems as if we add blank records, this just
		 * applies a format to the cell rather than actually replacing the valrec?

		if (c.getOpcode()!=MULBLANK) {	// KSC: Added
			BiffRec existing = (BiffRec)cells.get(Short.valueOf(cellCol));
			if( existing != null){
	            if (this.getWorkBook().getFactory().iscompleted()) {
	    		    if((c instanceof Blank)) {
	    		    	existing.setIxfe(c.getIxfe());
	    		        return;
	    		    }else {
	    		        cells.remove(Short.valueOf(cellCol));
	                    c.setRow(this);
	                    cells.put(Short.valueOf(cellCol), c);
	                    this.lastcell = c;
	    		    }
	            }
			}else {
	    		c.setRow(this);
	    		cells.put(Short.valueOf(cellCol), c);
	    		this.lastcell = c;
	        }
		} else { // expand mulblanks to each referenced cell
		 */

        /* We should be able to handle this with cellAddressible, hopefully.
			short colFirst= ((Mulblank)c).colFirst;
			short colLast= ((Mulblank) c).colLast;
			for (short i= colFirst; i <= colLast; i++) {
                cells.put(Short.valueOf(i), c);
                this.lastcell = c;
			}
		}
		 */
    }

    internal fun removeCell(c: BiffRec) {
        this.sheet!!.removeCell(c)
    }

    /**
     * remove cell via column number
     */
    internal fun removeCell(c: Short) {
        this.sheet!!.removeCell(this.rowNumber, c.toInt())
    }

    /**
     * Get a cell from the row
     *
     * @throws CellNotFoundException
     */
    @Throws(CellNotFoundException::class)
    fun getCell(d: Short): BiffRec {
        return this.sheet!!.getCell(this.rowNumber, d.toInt())
    }

    /**
     * Return an ordered array of the BiffRecs associated with this row.
     *
     *
     * This includes child records, and other non-cell-associated records
     * that should be in the row block (such as Formula Shrfmlas and Arrays)
     *
     * @int outputId = random id passed in that is specific to a worksheets output.  Allows tracking what
     * internal records (such as shared formulas) have been written already
     */
    fun getValRecs(outputId: Int): List<*> {
        val v = ArrayList()
        val cx = this.cells
        val it = cx.iterator()
        while (it.hasNext()) {
            val br = it.next()
            v.add(br)
            if (br is Formula) {
                br.preStream()    // must do now so can ensure internal records are properly set
                val itx = br.internalRecords
                val brints = itx.toTypedArray() as Array<BiffRec>
                for (x in brints.indices) {
                    //Don't allow dupes!
                    if (!v.contains(brints[x])) {
                        v.add(brints[x])
                    }
                }
            }
            if (br.hyperlink != null) {
                v.add(br.hyperlink)
            }
        }
        return v
    }

    override fun init() {
        super.init()
        this.getData()

        // get the number of the row
        rw = ByteTools.readUnsignedShort(this.getByteAt(0), this.getByteAt(1))
        colMic = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        colMac = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        miyRw = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt()).toInt()

        // bytes 8 - 11 are reserved
        val byte12 = this.getByteAt(12)

        /**
         * A - iOutLevel (3 bits): An unsigned integer that specifies the outline level (1) of the row.
         * B - reserved2 (1 bit): MUST be zero, and MUST be ignored.
         * C - fCollapsed (1 bit): A bit that specifies whether the rows that are one level of outlining deeper than the current row are included in the collapsed outline state.
         * D - fDyZero (1 bit): A bit that specifies whether the row is hidden.
         * E - fUnsynced (1 bit): A bit that specifies whether the row height was manually set.
         * F - fGhostDirty (1 bit): A bit that specifies whether the row was formatted.
         */
        outlineLevel = byte12 and 0x7
        fCollapsed = byte12 and 0x10 != 0    // ?? 0x8 ??
        fHidden = byte12 and 0x20 != 0
        isAlteredHeight = byte12 and 0x40 != 0
        explicitFormatSet = byte12 and 0x80 != 0
        /**/
        // byte 13 is reserved
        val byte15 = this.getByteAt(15)
        val byte14 = this.getByteAt(14)
        if (explicitFormatSet) { // then explicit ixfe set
            // The low-order byte is sbyte 14. The low-order nybble of the
            // high-order byte is stored in the high-order nybble of byte 15.
            ixfe = byte14 and 0xFF or (byte15 and 0xFF shl 8 and 0xFFF) // 12 bits
            if (ixfe < 0 || ixfe > this.workBook!!.numXfs) {    // KSC: TODO: ixfe calc is wrong ...?
                ixfe = 15//this.getWorkBook().getDefaultIxfe();
                explicitFormatSet = false
            }
        } else
            ixfe = 15//this.getWorkBook().getDefaultIxfe();

        /** f
         * from excel documentation:
         * fBorderTop=
         * G - fExAsc (1 bit): A bit that specifies whether any cell in the row has a thick top border, or any cell in the row directly above the current row has a thick bottom border.
         * Thick borders are specified by the following enumeration values from BorderStyle: THICK and DOUBLE.
         * fBorderBottom=
         * H - fExDes (1 bit): A bit that specifies whether any cell in the row has a medium or thick bottom border, or any cell in the row directly below the current row has a medium or thick top border.
         * Thick borders are previously specified. Medium borders are specified by the following enumeration values from BorderStyle: MEDIUM, MEDIUMDASHED, MEDIUMDASHDOT, MEDIUMDASHDOTDOT, and SLANTDASHDOT.
         */
        hasAnyThickTopBorder = byte15 and 0x10 != 0
        hasAnyBottomBorder = byte15 and 0x20 != 0
        fPhonetic = byte15 and 0x40 != 0
    }

    override fun preStream() {
        if (this.sheet != null) {
            this.updateColDimensions(colMac)
            this.updateColDimensions(colMic)
        } else {
            if (DEBUGLEVEL > -1)
                Logger.logWarn("Missing Boundsheet in Row.prestream for Row: " + this.rowNumber + this.cellAddress)
        }

        val data = ByteArray(16)
        data[0] = (rw and 0x00FF).toByte()
        data[1] = (rw and 0xFF00).ushr(8).toByte()
        data[2] = (colMic and 0x00FF).toByte()
        data[3] = (colMic and 0xFF00).ushr(8).toByte()
        data[4] = (colMac and 0x00FF).toByte()
        data[5] = (colMac and 0xFF00).ushr(8).toByte()
        data[6] = (miyRw and 0x00FF).toByte()
        data[7] = (miyRw and 0xFF00).ushr(8).toByte()
        // bytes 8 - 11 are reserved

        /**
         * A - iOutLevel (3 bits): An unsigned integer that specifies the outline level (1) of the row.
         * B - reserved2 (1 bit): MUST be zero, and MUST be ignored.
         * C - fCollapsed (1 bit): A bit that specifies whether the rows that are one level of outlining deeper than the current row are included in the collapsed outline state.
         * D - fDyZero (1 bit): A bit that specifies whether the row is hidden.
         * E - fUnsynced (1 bit): A bit that specifies whether the row height was manually set.
         * F - fGhostDirty (1 bit): A bit that specifies whether the row was formatted.
         */
        if (outlineLevel != 0)
            data[12] = data[12] or outlineLevel.toByte()
        if (fCollapsed) data[12] = data[12] or 0x10    // 0x8 ???
        if (fHidden) data[12] = data[12] or 0x20
        if (isAlteredHeight) data[12] = data[12] or 0x40
        if (explicitFormatSet) data[12] = data[12] or 0x80
        /**/
        // byte 13 is reserved
        data[13] = 1

        // The low-order byte is byte 14. The low-order nybble of the
        // high-order byte is stored in the high-order nybble of byte 15.
        data[14] = (ixfe and 0x00FF).toByte()
        data[15] = (ixfe shr 8).toByte()    //& 0x0F00) >>> 4);
        if (hasAnyThickTopBorder) data[15] = data[15] or 0x10
        if (hasAnyBottomBorder) data[15] = data[15] or 0x20
        if (fPhonetic) data[15] = data[15] or 0x40
        // byte 15 bit 0x01 is reserved

        this.setData(data)
    }


    /**
     * sets or clears the Unsynced flag
     * <br></br>The Unsynched flag is true if the row height is manually set
     * <br></br>If false, the row height should auto adjust when necessary
     *
     * @param bUnsynced
     */
    fun setUnsynched(bUnsynced: Boolean) {
        isAlteredHeight = bUnsynced
        updateGrbit()
    }

    /**
     * update the col indexes
     */
    fun updateColDimensions(col: Short) {
        var col = col
        if (col > Row.MAXCOLS)
            return
        var cl: ByteArray? = null
        if (col < colMic) {
            colMic = col
            cl = ByteTools.shortToLEBytes(colMic)
            System.arraycopy(cl!!, 0, this.getData()!!, 2, 2)
        }
        if (col > colMac) {
            colMac = col
            colMac = ++col
            cl = ByteTools.shortToLEBytes(colMac)
            System.arraycopy(cl!!, 0, this.getData()!!, 4, 2)
        }
    }

    override fun toString(): String {
        val celladdrs = StringBuffer()
        val cx = this.cells
        val it = cx.iterator()
        while (it.hasNext()) {
            celladdrs.append("{")
            celladdrs.append(it.next().toString())
            celladdrs.append("}")
        }
        return this.rowNumber.toString() + celladdrs.toString()
    }

    fun setHeight(twips: Int) {
        if (twips < 2 || twips > 8192)
            throw IllegalArgumentException(
                    "twips value " + twips
                            + " is out of range, must be between 2 and 8192 inclusive")

        miyRw = twips
        isAlteredHeight = true
    }

    fun clearHeight() {
        isAlteredHeight = false
    }

    /**
     * Set the Outline level (depth) of the row
     *
     * @param x
     */
    fun setOutlineLevel(x: Int) {
        this.outlineLevel = x
        this.sheet!!.guts!!.rowGutterSize = 10 + 10 * x
        this.sheet!!.guts!!.maxRowLevel = x + 1
        updateGrbit()
        //		implement bit masking set on grbit
    }

    /**
     * Update the internal Grbit based
     * on values existant in the row
     */
    private fun updateGrbit() {
        preStream()
    }


    /**
     * Returns the Outline level (depth) of the row
     *
     * @return
     */
    fun getOutlineLevel(): Int {
        return outlineLevel

    }

    /**
     * Removes the format currently applied to this row, if any.
     */
    fun clearIxfe() {
        ixfe = 0
        explicitFormatSet = false
    }

    /**
     * Returns whether a format has been set for this row.
     */
    fun hasIxfe(): Boolean {
        return explicitFormatSet
    }

    /**
     * flags this row to have at least one cell that has a thick bottom border
     *
     *
     * For internal use only
     *
     * @param hasBorder
     */
    fun setHasAnyThickBottomBorder(hasBorder: Boolean) {
        this.hasAnyBottomBorder = hasBorder
    }

    companion object {

        private val serialVersionUID = 6848429681761792740L

        internal val defaultsize = 16
    }
}