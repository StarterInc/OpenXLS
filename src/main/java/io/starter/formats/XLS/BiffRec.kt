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

import io.starter.OpenXLS.CellRange
import io.starter.formats.LEO.BlockByteReader

/**
 * To change the template for this generated type comment go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
interface BiffRec {


    // integrating Cell Methods
    val internalVal: Any

    var hyperlink: Hlink

    var mergeRange: CellRange

    var row: Row

    val formatPattern: String

    val font: Font

    val xfRec: Xf

    val formulaRec: Formula

    var intVal: Int

    val dblVal: Double

    var booleanVal: Boolean

    val isReadOnly: Boolean

    /**
     * return the record index of this object
     */
    val recordIndex: Int

    /**
     * returns the array of CONTINUE records for this record, containing all
     * data beyond the 8224 byte record size limit.
     */

    val continueVect: List<*>

    /**
     * get the WorkSheet for this record.
     */
    val sheet: Boundsheet

    var workBook: WorkBook


    val colNumber: Short

    var rowNumber: Int

    /**
     * get a string address for the cell based on row and col ie: "H22"
     */
    val cellAddress: String

    /**
     * get a default "empty" data value for this record
     */
    val defaultVal: Any

    /**
     * get the data type name for this record
     */
    val dataType: String

    /**
     * Get the value of the record as a Float. Value must be parseable as an
     * Float or it will throw a NumberFormatException.
     */
    val floatVal: Float

    /**
     * Get the value of the record as a String.
     */
    var stringVal: String

    /**
     * get the ixfe
     */
    var ixfe: Int


    var opcode: Short

    /**
     * Returns the length of this record, including the 4 header bytes
     */
    var length: Int

    var byteReader: BlockByteReader

    var encryptedByteReader: BlockByteReader

    /**
     * gets the full record bytes for this record including header bytes
     *
     * @return byte[] of all rec bytes
     */
    val bytes: ByteArray

    /**
     * gets the record data merging any Continue record data.
     */
    //** Thread Safing OpenXLS **//
    var data: ByteArray

    /**
     * Get the relative position within the data underlying the block vector
     * represented by the BlockByteReader.
     *
     * @return relative position
     */
    /**
     * Set the relative position within the data underlying the block vector
     * represented by the BlockByteReader.
     *
     *
     * In other words, this is the relative position used by the BlockByteReader
     * to offset the Consumer's read position within the collection of Data
     * Blocks.
     *
     *
     * This may be an offset relative to the data in a file, or within a Storage
     * contained in a file.
     *
     *
     * The Workbook Storage for example will contain a non-contiguous collection
     * of Blocks containing data from any number of positions in a file.
     *
     *
     * This collection forms a contiguous span of bytes comprising an XLS
     * Workbook. The XLSRecords within this span of bytes will set their
     * relative position within this 'virtual' array. Thus the XLSRecord
     * positions are relative to the order of bytes contained in the Block
     * collection. The BOF record then is at offset 0 within the data of the
     * first Block, even though the underlying data of this first Block may be
     * anywhere on disk.
     *
     * @param pos
     */
    var offset: Int

    /**
     * @return
     */
    /**
     * @param streamer
     */
    var streamer: ByteStreamer

    /**
     * @return
     */
    /**
     * set whether this record contains the value of the Cell.
     */
    var isValueForCell: Boolean

    fun copyFormat(source: BiffRec)

    fun remove(t: Boolean): Boolean

    /**
     * set the DEBUG level
     */
    fun setDebugLevel(b: Int)

    /**
     * adds a CONTINUE record to the array of CONTINUE records for this record,
     * containing all data beyond the 8224 byte record size limit.
     */
    fun addContinue(c: Continue)

    /**
     * remove all Continue records
     */
    fun removeContinues()

    /**
     * returns whether this record has a CONTINUE record containing data beyond
     * the 8228 byte record size limit.
     *
     *
     * XLSRecords can have 0 or more CONTINUE records.
     */
    fun hasContinues(): Boolean

    /**
     * associate this record with its Index record
     */
    fun setIndex(id: Index)

    /**
     * Associate this record with a worksheet. First checks to see if there is
     * already a cell with this address.
     */
    fun setSheet(b: Sheet)

    /**
     * set the column
     */
    fun setCol(i: Short)

    fun setRowCol(x: IntArray)

    /**
     * perform record initialization
     */
    fun init()

    /**
     * Get the value of the record as a String.
     */
    fun getStringVal(encoding: String): String

    fun setDoubleVal(v: Double)

    /**
     * do any pre-streaming processing such as expensive index updates or other
     * deferrable processing.
     */
    fun preStream()

    /**
     * do any post-streaming cleanup such as expensive index updates or other
     * deferrable processing.
     */
    fun postStream()

    /**
     * set the XF (format) record for this rec
     */
    fun setXFRecord()

    /**
     * set the XF (format) record for this rec
     */
    fun setXFRecord(i: Int)

    /**
     * Gets the byte from the specified position in the record byte array.
     *
     * @param off
     * @return
     */
    fun getByteAt(off: Int): Byte

    /**
     * Gets the byte from the specified position in the record byte array.
     *
     * @param off
     * @return
     */
    fun getBytesAt(off: Int, len: Int): ByteArray

    /**
     * Dumps this record as a human-readable string.
     * This method's output is more verbose than that of [.toString].
     * It generally includes a hex dump of the record's contents.
     */
    fun toHexDump(): String

    /**
     * collect cell change listeners and fire cell change event upon, well ...
     * @param t
     * 20080204 KSC
     *
     * public void addCellChangeListener(CellChangeListener t);
     * public void removeCellChangeListener(CellChangeListener t);
     * // public void fireCellChangeEvent();
     * public void setCachedValue(Object newValue);
     * public Object getCachedValue();
     */
}