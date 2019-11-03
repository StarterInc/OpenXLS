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

import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.io.Serializable


/**
 * **Index: Index Record 0x20B**<br></br>
 *
 *
 * Index records are written after the Bof record for each Boundsheet.
 *
 * <pre>
 *
 * offset  name        size    contents
 * ---
 * 4       Reserved    4       Must be zero
 * 8       rwMic       4       First row that exists on the sheet
 * 12      rwMac       4       Last row that exists on the sheet, plus 1
 * 16      Reserved    4       Pointer to DIMENSIONS record for Boundsheet
 * 20      rgibRw      var     Array of file offsets to the Dbcell records
 * for each block of ROW records.  A block
 * contains up to 32 ROW records.
 *
 *
 * When a record value changes in size, it fires a CellChangeEvent
 * which cascades through the other associated objects.
 *
 * The record size change has the following effect on INDEX record fields:
 *
 * 1. ALL Dbcell records in the file starting with the one
 * containing the changed record move.  This requires updating
 * the dbcellpointers (rgibRw) in ALL INDEX records which
 * are located after the changed Dbcell within the file stream.
 *
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Dbcell
 *
 * @see Row
 *
 * @see Cell
 *
 * @see XLSRecord
 */

class Index : io.starter.formats.XLS.XLSRecord(), XLSConstants {
    private var rwMic: Int = 0
    private var rwMac: Int = 0
    private var dbnum = 0
    //    private dbCellPointer[] dbcellarray; not used
    private var dbcells = CompatibleVector()
    /**
     * get the index number for
     * addressing.
     */
    /**
     * set the index number for
     * addressing.
     */
    var indexNum: Int = 0
    private var dims: Dimensions? = null

    /**
     * return the associated Dbcell objects
     * for this INDEX.
     */
    internal val dbCells: Array<Dbcell>
        get() {
            val obj = dbcells.toTypedArray()
            val dbcs = arrayOfNulls<Dbcell>(obj.size)
            System.arraycopy(obj, 0, dbcs, 0, obj.size)
            return dbcs
        }

    internal var offsetStart = 0

    /**
     * fire the cell change event
     *
     *
     * public void fireCellChangeEvent(CellChangeEvent c){
     * // do its thing
     * // this.doCellSizeChangeAction(c);
     * // then pass it along...
     * //  this.getSheet().fireCellChangeEvent(c);
     * }
     */

    internal fun setDimensions(d: Dimensions) {
        this.dims = d
    }


    // Not used??
    internal fun setDimensionsOffset(offset: Int) {
        val recData = this.getData()
        val newoff = ByteTools.cLongToLEBytes(offset)
        System.arraycopy(newoff, 0, recData!!, 12, 4)
        this.setData(recData)
    }


    /**
     * add an associated Dbcell object
     * to this INDEX.
     */
    internal fun addDBCell(dbc: Dbcell) {
        val bAdd = true
        /* KSC: TESTING testLostBorders bug   for (int i= 0; i < dbcells.size() && bAdd; i++) {
    	if (Arrays.equals(((Dbcell) dbcells.get(i)).data,  dbc.data))
			bAdd= false;
    }*/
        if (bAdd) {
            //        if(!dbcells.contains(dbc)){
            dbc.dbcNum = dbnum++
            dbcells.add(dbc)
            //        }
        }
    }

    internal fun resetDBCells() {
        dbnum = 0
        dbcells = CompatibleVector()
    }

    /**
     * Update the Dbcell Pointers
     *
     *
     * At this point the index should have all of it's dbcells in place, as well as
     * thier offsets being populated, so all that should need to be done is to create
     * the index correctly out of its values
     */

    internal fun updateDbcellPointers() {
        streamer = sheet!!.streamer
        // first, get the collection of Rows from sheet
        val bs = this.sheet
        val rowz = bs!!.rows
        if (rowz.size != 0) {
            this.updateRowDimensions(rowz[0].rowNumber, rowz[rowz.size - 1].rowNumber)
        }
        // create the new Dbcells if any rows exist within the sheet

        // rebuild the record with the correct length body data to fit the new dbcells
        val arrsize = 16 + dbcells.size * 4
        val newBytes = ByteArray(arrsize)
        val dbc: Dbcell? = null
        System.arraycopy(this.getData()!!, 0, newBytes, 0, 16)
        var offset = 16
        for (i in dbcells.indices) {
            val db = dbcells.elementAt(i) as Dbcell
            val dbOff = db.getOffset()
            val b = ByteTools.cLongToLEBytes(dbOff)
            newBytes[offset++] = b[0]
            newBytes[offset++] = b[1]
            newBytes[offset++] = b[2]
            newBytes[offset++] = b[3]
        }
        this.setData(newBytes)
    }


    /**
     * update the dimensions info based on Dimensions rec
     */
    internal fun updateDimensions() {
        val rkdata = this.getData()
        val newb = ByteTools.cLongToLEBytes(dims!!.offset)
        rkdata[12] = newb[0]
        rkdata[13] = newb[1]
        rkdata[14] = newb[2]
        rkdata[15] = newb[3]

        /* these should match rwMic/rwMac on dimensions rec
        dims.rwMic;
        dims.rwMac;
        */
    }

    override fun init() {
        super.init()
        // 1st 4 are reseverd-0
        rwMic = ByteTools.readInt(this.getBytesAt(4, 4)!!)
        rwMac = ByteTools.readInt(this.getBytesAt(8, 4)!!)
        // next 4 are position of defColWidth record - skip
        /* no need to read in dbcell offsets as we don't do anything with it
 * 		int pos= 16;
		int recsize= 4; // KSC added
		int numdbcells = (this.getLength()-pos)/recsize;
//		dbcellarray = new dbCellPointer[numdbcells];
		// rest of data is position of dbCell records: rgibRw (variable): An array of FilePointer. Each FilePointer specifies the file position of each referenced DBCell record
		for(int i = 0;i< numdbcells;i++){
		    if(DEBUGLEVEL > 6)Logger.logInfo("Index -> initializing dbcell pointer: " + i);
		    byte[] bite = this.getBytesAt(pos,4);
		    dbCellPointer pointer = new dbCellPointer(bite);
		    pos += 4;
//		    dbcellarray[i] = pointer;
		}
//		*/
    }

    /** compute the location of Dbcell records using
     * the INDEX dbcellpointers and the firstBof record position
     * in the workbook.
     *
     * NOT USED
     *
     * int getDbcellPosition(int pointernum){
     * int firstBofloc = wkbook.getFirstBof().offset;
     * byte[] b = dbcellarray[pointernum].cdb;
     * int pointerloc = ByteTools.readInt(b[0],b[1],b[2],b[3]);
     * return pointerloc + firstBofloc;
     * } */

    /**
     * file offset to the Dbcell record
     */
    internal inner class dbCellPointer(b: ByteArray) : Serializable {
        var cellloc = 0
        var datasiz = 0
        var s2: Short = 0
        var s3: Short = 0
        var cdb = ByteArray(4)

        val bytes: ByteArray
            get() {
                val bite = ByteArray(4)
                System.arraycopy(ByteTools.shortToLEBytes(cellloc.toShort()), 0, bite, 0, 2)
                System.arraycopy(ByteTools.shortToLEBytes(datasiz.toShort()), 0, bite, 2, 2)
                return bite
            }

        init {
            cdb = b
            cellloc = ByteTools.readShort(b[0].toInt(), b[1].toInt()).toInt()
            datasiz = ByteTools.readShort(b[2].toInt(), b[3].toInt()).toInt()
        }

        /*** Updates location of Dbcell pointer and data size(?)
         */
        fun adjustPosition(i: Int) {
            cellloc += i
            datasiz += i
        }

        companion object {

            /**
             * serialVersionUID
             */
            private const val serialVersionUID = -5132922970171084839L
        }
    }

    /**
     * update the min/max cols and rows
     * 8       rwMic       4       First row that exists on the sheet
     * 12      rwMac       4       Last row that exists on the sheet, plus 1
     */
    fun updateRowDimensions(lowRow: Int, hiRow: Int) {
        var rw = ByteTools.cLongToLEBytes(lowRow)
        System.arraycopy(rw, 0, this.getData()!!, 4, 4)
        rw = ByteTools.cLongToLEBytes(hiRow + 1)
        System.arraycopy(rw, 0, this.getData()!!, 8, 4)
    }

    /**
     * Called from streamer, this updates individual dbcell offset values.
     *
     *
     * Will only run correctly if called sequentially, ie dboffset [0], [1], [2]
     *
     * @param DbcellNumber - which dbcell to update
     * @param DbOffset     - the pure offset from beginning of file
     */
    internal fun setDbcellPosition(DbcellNumber: Int, DbOffset: Int) {
        if (offsetStart == 0) {
            offsetStart = this.sheet!!.myBof!!.getOffset()
        }
        val insertOffset = DbOffset - offsetStart
        if (DEBUGLEVEL > 10) {
            Logger.logInfo("Setting DBBiffRec Position, offsetStart:$offsetStart & InsertOffset = $insertOffset")
        }
        offsetStart += insertOffset
        val insertloc = 16 + DbcellNumber * 4
        val off = ByteTools.cLongToLEBytes(insertOffset)
        System.arraycopy(off, 0, data!!, insertloc, 4)

    }


    /**
     * Prestream for Index is going to create the correct size record, and populate the correct number of dbcells.
     * The actual values will not yet be populated, but the record sizes will be correct in order to get offsets
     * working correctly.
     *
     *
     * Once offsets are correctly calculated in bytestreamer.stream, we can come back and
     * populate without the getIndex call overhead.
     */
    override fun preStream() {
        // rebuild the record with the correct length body data to fit the new dbcells
        this.getData()
        val arrsize = 16 + dbcells.size * 4
        val newBytes = ByteArray(arrsize)
        val dbc: Dbcell? = null
        // KSC: Changed from copying 12 bytes to copying 16 bytes to keep DIMENSIONS reference
        System.arraycopy(this.getData()!!, 0, newBytes, 0, 16)
        this.setData(newBytes)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -753407655976707961L
        private val defaultsize = 16

        /**
         * create a new INDEX rec
         */
        // default val
        val prototype: XLSRecord?
            get() {
                val idx = Index()
                val dt = ByteArray(defaultsize)
                idx.originalsize = defaultsize
                idx.setData(dt)
                idx.opcode = XLSConstants.INDEX
                idx.length = defaultsize.toShort()
                idx.init()
                return idx
            }
    }
}

