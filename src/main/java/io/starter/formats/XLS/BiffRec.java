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
package io.starter.formats.XLS;

import io.starter.OpenXLS.CellRange;
import io.starter.formats.LEO.BlockByteReader;

import java.util.List;

/**
 * To change the template for this generated type comment go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
public interface BiffRec {


    // integrating Cell Methods
    Object getInternalVal();

    void setHyperlink(Hlink h);

    Hlink getHyperlink();

    CellRange getMergeRange();

    void setMergeRange(CellRange r);

    Row getRow();

    String getFormatPattern();

    void copyFormat(BiffRec source);

    Font getFont();

    Xf getXfRec();

    boolean remove(boolean t);

    Formula getFormulaRec();

    void setRow(Row r);

    void setRowNumber(int i);


    void setIxfe(int xf);

    int getIntVal();

    double getDblVal();

    boolean getBooleanVal();

    boolean isReadOnly();

    /**
     * set the DEBUG level
     */
    void setDebugLevel(int b);

    /**
     * return the record index of this object
     */
    int getRecordIndex();

    /**
     * adds a CONTINUE record to the array of CONTINUE records for this record,
     * containing all data beyond the 8224 byte record size limit.
     */
    void addContinue(Continue c);

    /**
     * remove all Continue records
     */
    void removeContinues();

    /**
     * returns the array of CONTINUE records for this record, containing all
     * data beyond the 8224 byte record size limit.
     */

    List getContinueVect();

    /**
     * returns whether this record has a CONTINUE record containing data beyond
     * the 8228 byte record size limit.
     * <p>
     * XLSRecords can have 0 or more CONTINUE records.
     */
    boolean hasContinues();

    /**
     * set whether this record contains the value of the Cell.
     */
    void setIsValueForCell(boolean b);

    /**
     * associate this record with its Index record
     */
    void setIndex(Index id);

    /**
     * Associate this record with a worksheet. First checks to see if there is
     * already a cell with this address.
     */
    void setSheet(Sheet b);

    /**
     * get the WorkSheet for this record.
     */
    Boundsheet getSheet();

    void setWorkBook(WorkBook wk);

    WorkBook getWorkBook();

    /**
     * set the column
     */
    void setCol(short i);

    void setRowCol(int[] x);


    short getColNumber();

    int getRowNumber();

    /**
     * get a string address for the cell based on row and col ie: "H22"
     */
    String getCellAddress();

    /**
     * perform record initialization
     */
    void init();

    /**
     * get a default "empty" data value for this record
     */
    Object getDefaultVal();

    /**
     * get the data type name for this record
     */
    String getDataType();

    /**
     * Get the value of the record as a Float. Value must be parseable as an
     * Float or it will throw a NumberFormatException.
     */
    float getFloatVal();

    /**
     * Get the value of the record as a String.
     */
    String getStringVal();

    /**
     * Get the value of the record as a String.
     */
    String getStringVal(String encoding);

    void setStringVal(String v);

    void setBooleanVal(boolean b);

    void setIntVal(int v);

    void setDoubleVal(double v);

    /**
     * do any pre-streaming processing such as expensive index updates or other
     * deferrable processing.
     */
    void preStream();

    /**
     * do any post-streaming cleanup such as expensive index updates or other
     * deferrable processing.
     */
    void postStream();

    /**
     * set the XF (format) record for this rec
     */
    void setXFRecord();

    /**
     * set the XF (format) record for this rec
     */
    void setXFRecord(int i);

    /**
     * get the ixfe
     */
    int getIxfe();


    short getOpcode();

    void setOpcode(short opcode);

    /**
     * Returns the length of this record, including the 4 header bytes
     */
    int getLength();

    void setLength(int len);

    void setByteReader(BlockByteReader db);

    BlockByteReader getByteReader();

    void setEncryptedByteReader(BlockByteReader db);

    BlockByteReader getEncryptedByteReader();

    //** Thread Safing OpenXLS **//
    void setData(byte[] b);

    /**
     * gets the full record bytes for this record including header bytes
     *
     * @return byte[] of all rec bytes
     */
    byte[] getBytes();

    /**
     * gets the record data merging any Continue record data.
     */
    byte[] getData();

    /**
     * Gets the byte from the specified position in the record byte array.
     *
     * @param off
     * @return
     */
    byte getByteAt(int off);

    /**
     * Gets the byte from the specified position in the record byte array.
     *
     * @param off
     * @return
     */
    byte[] getBytesAt(int off, int len);

    /**
     * Set the relative position within the data underlying the block vector
     * represented by the BlockByteReader.
     * <p>
     * In other words, this is the relative position used by the BlockByteReader
     * to offset the Consumer's read position within the collection of Data
     * Blocks.
     * <p>
     * This may be an offset relative to the data in a file, or within a Storage
     * contained in a file.
     * <p>
     * The Workbook Storage for example will contain a non-contiguous collection
     * of Blocks containing data from any number of positions in a file.
     * <p>
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
    void setOffset(int pos);

    /**
     * Get the relative position within the data underlying the block vector
     * represented by the BlockByteReader.
     *
     * @return relative position
     */
    int getOffset();

    /**
     * @return
     */
    ByteStreamer getStreamer();

    /**
     * @param streamer
     */
    void setStreamer(ByteStreamer streamer);

    /**
     * @return
     */
    boolean isValueForCell();

    /**
     * Dumps this record as a human-readable string.
     * This method's output is more verbose than that of {@link #toString()}.
     * It generally includes a hex dump of the record's contents.
     */
    String toHexDump();

    /**
     * collect cell change listeners and fire cell change event upon, well ...  
     * @param t
     * 20080204 KSC

    public void addCellChangeListener(CellChangeListener t);
    public void removeCellChangeListener(CellChangeListener t);
    // public void fireCellChangeEvent();
    public void setCachedValue(Object newValue);
    public Object getCachedValue();
     */
}