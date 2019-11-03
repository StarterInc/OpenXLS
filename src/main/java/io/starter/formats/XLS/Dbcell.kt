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
import io.starter.toolkit.Logger


/**
 * **Dbcell: Stream Offsets 0xD7**<br></br>
 *
 *
 * Offsets for value records.  There is one DBCELL for
 * each 32-row block of Row records and associated cell records.
 *
 * <br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       dbRtrw      4       Offset from the start of the DBCELL
 * to the start of the first Row in the block
 * 8       rgdb        var     Array of stream offsets, 2 bytes each.
 *
 * The internal format layout is basically:
 *
 *
 * When a record value changes in size, it fires a CellChangeEvent
 * which cascades through the other associated objects.
 *
 * The record size change has the following effects on DBCELL record fields:
 *
 * 1. Row records for the data block move relative to the
 * DBCELL for the block.  The dbRtrw field tracks the position
 * of the first Row record, so this needs to be updated in the
 * DBCELL only for the changed record block.
 *
 * 2. The *record*  offsets for the Row stored in the rgdb array
 * change starting with the changed value.  These only change
 * for the record block containing the changed value -- subsequent
 * blocks maintain their relative position to their row and row
 * records.
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Index
 *
 * @see Dbcell
 *
 * @see Row
 *
 * @see Cell
 *
 * @see XLSRecord
 */
class Dbcell : io.starter.formats.XLS.XLSRecord(), XLSConstants {
    private var mybof: Bof? = null

    private val rkdata = ByteArray(0)
    private var numrecs = -1
    private var dbRtrw: Int = 0
    /**
     * get the DBCELL number for use by the Index
     */
    /** handle a cell change event
     *
     * public void fireCellChangeEvent(CellChangeEvent c){
     * // do its thing,
     * this.doCellSizeChangeAction(c);
     * // then pass it along...
     * this.getIDX().fireCellChangeEvent(c);
     * } */

    /**
     * set the DBCELL number for use by the Index
     */
    internal var dbcNum: Int = 0
    private var rgdb: ShortArray? = null
    /**
     * get the associated dbcell index
     */
    /**
     * set the associated dbcell index
     */
    var idx: Index? = null
        internal set

    internal var myrows: Array<Row>? = arrayOfNulls<Row>(32)

    /** set the rows for this dbcell
     *
     * public void setRows(AbstractList r){
     * this.myrows = r;
     * for(int i = 0;i<myrows.size></myrows.size>();i++){
     * Row rt = (Row) myrows.get(i);
     * rt.setDBCell(this);
     * }
     * } */

    /**
     * returns the number of rows
     * contained in this DBCELL.  There should
     * never be more than 32.
     */
    // if(ret > 32)if(DEBUGLEVEL > -1)Logger.logInfo("DBCELL has too many rows: "+ String.valueOf(ret));
    //return ret;
    var numRows = 0
        internal set

    /** iterates through XLSRecords Contained in Rows
     * and gets offset pointers
     *
     * Deprecated?  THinkso
     *
     * void updateIndexes(){
     * Iterator e = myrows.iterator();
     * int startingSize = this.getLength();
     * int[] idxes = new int[myrows.size()];
     * int i = 0, pos = 4;
     * byte[] newrgdb = new byte[(myrows.size()*2)+4];
     * int offst = -1;
     * while(e.hasNext()){
     * Row c = (Row) e.next();
     * idxes[i]=c.getFirstCellOffset();
     * if(i==0){  // read page 443 of Excel Format book for info on this...
     * offst = idxes[i];
     * idxes[i] += 6;
     * Row r1 = null;
     * if(myrows.size() > 1) r1 = (Row) this.myrows.get(1);
     * else r1 = (Row) this.myrows.get(0);
     * int dd = r1.offset + r1.getLength();
     * if(false)Logger.logInfo("Initializing of Dbcell: " + this.getdbRtrw() + " - " + dd);
     * idxes[i] = idxes[i] - dd;
     * if(myrows.size()>1)idxes[i]+=r1.getLength();
     * }else{
     * idxes[i] -= offst;
     * //  offst+=idxes[i];
     * }
     * if(false)Logger.logInfo(" dbcellpointer: " + String.valueOf( idxes[i] ));
     * byte[] barr = ByteTools.shortToLEBytes((short)idxes[i++]);
     * System.arraycopy(barr, 0, newrgdb, pos, 2);
     * pos+=2;
     * }
     * this.setData(newrgdb);
     * this.setDbRtrw(this.getdbRtrw());
     * this.init();
     * }
     */

    /** returns whether this RowBlock can
     * hold any more Row records
     *
     * boolean isFull(){
     * if(myrows.size() >= 32)return true;
     * return false;
     * } */

    //Row[] getRows(){return myrows;}

    /**
     * returns the byte array containing the DBCELL location
     * as an offset from the Bof for the BOUNDSHEET.
     *
     *
     * this is used by the Index recordto locate RowBlocks.
     */
    internal val dbcellPointerPos: ByteArray
        get() {
            val bofpos = mybof!!.offset
            val thispos = this.offset
            val diff = thispos - bofpos
            return ByteTools.cLongToLEBytes(diff)
        }

    internal fun setBof(b: Bof) {
        mybof = b
    }


    /**
     * Init the dbcell with it's new values
     *
     *
     * TODO: review these calls to look for performance gain possibilities.  This method can
     * get called a lot on output.
     *
     * @param offsetToFirstRow - offset from the dbcell to the top row
     * @param valrecOffsets    - array of offsets from row to row in valrecs
     */
    fun initDbCell(offsetToFirstRow: Int, valrecOffsets: List<*>) {
        val newData = ByteArray(4 + (valrecOffsets.size - 1) * 2)
        val firstOffset = ByteTools.cLongToLEBytes(offsetToFirstRow)
        System.arraycopy(firstOffset, 0, newData, 0, 4)
        val it = valrecOffsets.iterator()
        var pointer = 4
        while (it.hasNext()) {
            val s = it.next() as Short
            // we don't want the last length, not needed, end of chain!
            if (it.hasNext()) {
                val b = ByteTools.shortToLEBytes(s)
                newData[pointer++] = b[0]
                newData[pointer++] = b[1]
            }
        }
        this.setData(newData)
    }

    /**
     * set the Index record for this RowBlock
     *
     *
     * there are many RowBlocks per Index record
     * Index records need to update their DBCELL
     * pointers when we add or move a DBCELL.
     */
    override fun setIndex(idx: Index) {
        this.idx = idx
    }

    /**
     * this method adds a Row to the Dbcell, as well
     * as updating the Dbcell position based on existing
     * row data.  For use in Index updateDbcellPOinters.
     */
    internal fun addRow(r: Row): Boolean {
        //   if(!this.isFull()){
        myrows[numRows++] = r
        return true
        //    }else return false;
    }


    /**
     * gets the value and updates the pointer.  this is the position
     * of the first row in the row block.
     */
    fun getdbRtrw(): Int {
        if (numRows > 0) {
            val rw1 = this.myrows!![0]
            val i = rw1.getOffset()
            val y = this.getOffset()
            return y - i
        }
        return -1
    }

    /**
     * Initialize the Dbcell
     */
    override fun init() {
        super.init()
        dbRtrw = ByteTools.readInt(this.getByteAt(0), this.getByteAt(1), this.getByteAt(2), this.getByteAt(3))
        numrecs = (this.length - 8) / 2
        var pos = 4
        rgdb = ShortArray(numrecs)
        for (i in 0 until numrecs) {
            rgdb[i] = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt())
        }
        if (DEBUGLEVEL > 10) {
            //     Logger.logInfo("DBCELL POINTER at: " + String.valueOf(dbRtrw));
            for (t in rgdb!!.indices) {
                Logger.logInfo(" rgdb" + t + ":" + rgdb!![t])
            }
            Logger.logInfo(" num idxs: $numrecs")
        }
    }

    /**
     * set the dbRtrw pointer location.
     */
    internal fun setDbRtrw(`val`: Int) {
        dbRtrw = `val`
        val b = ByteTools.cLongToLEBytes(`val`)
        System.arraycopy(b, 0, getData()!!, 0, 4)
    }

    /**
     *
     */
    override fun preStream() {
        //this.updateIndexes();
    }

    override fun close() {
        super.close()
        rgdb = null
        this.idx = null
        myrows = null
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3169134298616400374L


        /**
         * get a new, empty DBCELL
         */
        // default val
        val prototype: XLSRecord?
            get() {
                val dbc = Dbcell()
                val dt = ByteArray(4)
                dbc.originalsize = 4
                dbc.setData(dt)
                dbc.opcode = XLSConstants.DBCELL
                dbc.length = 4.toShort()
                dbc.init()
                return dbc
            }
    }

}