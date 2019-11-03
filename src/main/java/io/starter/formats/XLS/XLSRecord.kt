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
import io.starter.OpenXLS.ExcelTools
import io.starter.OpenXLS.FormatHandle
import io.starter.formats.LEO.BlockByteConsumer
import io.starter.formats.LEO.BlockByteReader
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.io.ByteArrayOutputStream
import java.io.IOException
import java.io.Serializable
import java.util.AbstractList
import java.util.ArrayList


/**
 * <pre>
 * The XLS byte stream is composed of records delimited by a header with type
 * and data length information, followed by a record Body which contains
 *
 * the byte[] data for the record.
 *
 * XLSRecord subclasses provide specific functionality.
 *
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
 */

open class XLSRecord : BiffRec, BlockByteConsumer, Serializable, XLSConstants {

    override var opcode: Short = 0
    internal var reclen: Int = 0
    internal var data: ByteArray? = null
    @Transient
    private var databuf: BlockByteReader? = null
    /**
     * Hold onto the original encrypted bytes so we can do a look ahead on records
     */
    @Transient
    override var encryptedByteReader: BlockByteReader? = null
    protected var isContinueMerged = false
    @Transient
    protected var DEBUGLEVEL = 0
    /**
     * @return Returns the isValueForCell.
     */
    /**
     * set whether this record contains the  value
     * of the Cell.
     */
    override var isValueForCell: Boolean = false
    internal var isFPNumber: Boolean = false
    internal var isDoubleNumber = false
    internal var isIntNumber: Boolean = false
    var isString: Boolean = false
    var isBoolean: Boolean = false
    /**
     * return whether this is a formula record
     */
    var isFormula: Boolean = false
        internal set
    var isBlank: Boolean = false
    /**
     * @return Returns the isReadOnly.
     */
    override var isReadOnly = false
        internal set
    protected var rw = -1
    override var colNumber: Short = 0
        protected set
    var offset = 0 // byte stream offset of rec
    @Transient
    protected var idx: Index? = null
    protected var worksheet: Sheet? = null
    @Transient
    var myxf: Xf? = null
    @Transient
    var wkbook: WorkBook? = null
    /**
     * @return Returns the streamer.
     */
    /**
     * @param streamer The streamer to set.
     */
    @Transient
    override var streamer: ByteStreamer? = null
    var continues: AbstractList<*>? = null
    internal var originalsize = -1
    internal var originalIndex = 0
    internal var originalOffset = 0
    internal var ixfe = -1
    /**
     * @return Returns the hyperlink.
     */
    override var hyperlink: Hlink? = null
    internal var myrow: Row? = null
    /**
     * @return
     */
    /**
     * @param range
     */
    override var mergeRange: CellRange? = null
    /**
     * Returns the index of the first block
     *
     * @return
     */
    /** Get the blocks containing this Consumer's data
     *
     * @return
     */
    /*unused: public Block[] getBlocks() {
		return myblocks;
	}*/

    /** Set the blocks containing this Consumer's data
     *
     * @param myblocks
     */
    /* unused: 	public void setBlocks(Block[] myb) {
		myblocks = myb;
	}*/

    /**
     * Sets the index of the first block
     *
     * @return
     */
    override var firstBlock: Int = 0
    /**
     * Returns the index of the last block
     *
     * @return
     */
    /**
     * Sets the index of the last block
     *
     * @return
     */
    override var lastBlock: Int = 0

    override val formulaRec: Formula?
        get() {
            if (this is Formula) {
                this.populateExpression()
                return this
            }
            return null
        }

    /**
     * get the row of this cell
     */
    override var row: Row
        get() {
            if (myrow == null) {
                myrow = this.worksheet!!.getRowByNumber(rw)
            }
            return myrow
        }
        set(r) {
            myrow = r
        }

    override val xfRec: Xf?
        get() {
            if (myxf == null) {
                if (ixfe > -1 && ixfe < wkbook!!.numXfs) {
                    this.myxf = wkbook!!.getXf(this.ixfe)
                }
            }
            return myxf
        }


    /**
     * returns the existing font record
     * for this Cell
     */
    override val font: Font?
        get() {
            val b = this.workBook ?: return null
            this.xfRec
            return if (myxf == null) null else myxf!!.font
        }

    /**
     * set the Formatting for this BiffRec from the pattern
     * match.
     *
     *
     * case insensitive pattern match is performed...
     */
    override val formatPattern: String?
        get() = if (myxf == null) "" else myxf!!.formatPattern

    /**
     * get the int val of the type for the valrec
     */
    val cellType: Int
        get() {
            if (this.isBlank) return XLSConstants.TYPE_BLANK
            if (this.isString) return XLSConstants.TYPE_STRING
            if (this.isDoubleNumber) return XLSConstants.TYPE_DOUBLE
            if (this.isFPNumber) return XLSConstants.TYPE_FP
            if (this.isIntNumber) return XLSConstants.TYPE_INT
            if (this.isFormula) return XLSConstants.TYPE_FORMULA
            return if (this.isBoolean) XLSConstants.TYPE_BOOLEAN else -1
        }

    // record name
    // hex record number
    // stream offset
    // size
    // file offset
    /*originalFileOffset*/// cell address, if applicable
    val recDesc: String
        get() {
            var ret = ""
            val name = this.javaClass.getSimpleName()
            ret += if (name == "XLSRecord") "unknown" else name.toUpperCase()
            ret += " (" + Integer.toHexString(opcode.toInt()).toUpperCase() + "h)"
            ret += " at " + Integer.toHexString(offset).toUpperCase() + "h"
            ret += " length " + Integer.toHexString(reclen).toUpperCase() + "h"
            ret += " file " + if (databuf == null)
                "no file"
            else
                databuf!!.getFileOffsetString(offset, reclen)
            if (this.isValueForCell) ret += " cell " + this.cellAddress

            return ret
        }

    /**
     * return whether this is a numeric type
     */
    val isNumber: Boolean
        get() = if (this.opcode == XLSConstants.RK) true else this.opcode == XLSConstants.NUMBER

    var lastidx = -1

    /**
     * return the real (not just boundsheet) record index of this object
     */
    val realRecordIndex: Int
        get() = streamer!!.getRealRecordIndex(this)

    /**
     * return the record index of this object
     */
    override// KSC: Added
    val recordIndex: Int
        get() = if (streamer == null) {
            if (this.sheet != null) this.sheet!!.sheetRecs.indexOf(this) else -1
        } else streamer!!.getRecordIndex(this)

    override val continueVect: List<*>
        get() {
            if (continues != null) return this.continues
            continues = CompatibleVector()
            return continues
        }

    /**
     * get the WorkSheet for this record.
     */
    override val sheet: Boundsheet?
        get() = worksheet as Boundsheet?

    override var workBook: WorkBook?
        get() {
            if (wkbook == null && worksheet != null)
                wkbook = worksheet!!.workBook

            return wkbook
        }
        set(wk) {
            wkbook = wk
        }

    /**
     * set the row
     */
    override var rowNumber: Int
        get() {
            if (rw < 0) {
                val rowi = rw * -1
                if (wkbook!!.isExcel2007)
                    rw = XLSConstants.MAXROWS - rowi
                else
                    rw = XLSConstants.MAXCOLS_BIFF8 - rowi
            }
            return rw
        }
        set(i) {
            val r = ByteTools.cLongToLEBytes(i)
            System.arraycopy(r, 0, getData()!!, 0, 2)
            rw = i
        }

    /**
     * get a string address for the
     * cell based on row and col ie: "H22"
     */
    override// > 32k and the rows go negative... !
    // the very last row...MAXROWS_BIFF8
    val cellAddress: String
        get() {
            var rownum = rw + 1
            if (rownum < 0 && colNumber >= 0) {
                rownum = XLSConstants.MAXROWS + rownum
            } else if (rownum == 0 && colNumber >= 0) {
                rownum = XLSConstants.MAXROWS
            }
            if (rownum > XLSConstants.MAXROWS || colNumber < 0) {
                if (DEBUGLEVEL > -1)
                    Logger.logWarn("XLSRecord.getCellAddress() Row/Col info incorrect for Cell:" + ExcelTools.getAlphaVal(colNumber.toInt()) + rownum)
                return ""
            }
            return ExcelTools.getAlphaVal(colNumber.toInt()) + rownum
        }

    /**
     * returns the cell address in int[] {row, col} format
     */
    open val intLocation: IntArray
        get() = intArrayOf(rw, colNumber.toInt())


    /**
     * return the cell address with sheet reference
     * eg Sheet!A12
     *
     * @return String
     */
    val cellAddressWithSheet: String
        get() = if (this.sheet != null) this.sheet!!.sheetName + "!" + this.cellAddress else this.cellAddress

    /**
     * get a default "empty" data value for this record
     */
    override val defaultVal: Any?
        get() {
            if (this.isDoubleNumber) return 0.0
            if (this.isFPNumber) return 0.0f
            if (this.isBoolean) return java.lang.Boolean.valueOf(false)
            if (this.isIntNumber) return Integer.valueOf(0)
            if (this.isString) return ""
            if (this.isFormula) return this.formulaRec!!.formulaString
            return if (this.isBlank) "" else null
        }

    /**
     * get the data type name for this record
     */
    override val dataType: String?
        get() {
            if (this.isValueForCell) {
                if (this.isBlank) return "Blank"
                if (this.isDoubleNumber) return "Double"
                if (this.isFPNumber) return "Float"
                if (this.isBoolean) return "Boolean"
                if (this.isIntNumber) return "Integer"
                if (this.isString) return "String"
                if (this.isFormula) return "Formula"
            }
            return null
        }

    /**
     * get the int val of the type for the valrec
     */
    override//essentially return "";
    // always use Doubles to avoid loss of precision... see:
    // details http://stackoverflow.com/questions/916081/convert-float-to-double-without-losing-precision
    // OK this is broken, obviously we need to return a calced Object
    //					return getStringVal();
    // should never happen here...
    val internalVal: Any?
        get() {
            try {
                when (this.cellType) {
                    XLSConstants.TYPE_BLANK -> return stringVal

                    XLSConstants.TYPE_STRING -> return stringVal

                    XLSConstants.TYPE_FP -> return dblVal

                    XLSConstants.TYPE_DOUBLE -> return dblVal

                    XLSConstants.TYPE_INT -> return Integer.valueOf(intVal)

                    XLSConstants.TYPE_FORMULA -> {
                        return (this as Formula).calculateFormula()
                    }

                    XLSConstants.TYPE_BOOLEAN -> return java.lang.Boolean.valueOf(booleanVal)

                    else -> return null
                }
            } catch (e: Exception) {
                return null
            }

        }


    /**
     * Get the value of the record as an Object.
     * To use the Object, cast it to the native
     * type for the record.  Ie: a String value
     * would need to be cast to a String, an Integer
     * to an Integer, etc.
     */
    internal open val `val`: Any?
        get() = null

    /**
     * Get the value of the record as a Boolean.
     * Value must be parseable as a Boolean.
     */
    override var booleanVal: Boolean
        get() = false
        set(b) = Logger.logErr("Setting Boolean Val on generic XLSRecord, value not held")

    /**
     * Get the value of the record as an Integer.
     * Value must be parseable as an Integer or it
     * will throw a NumberFormatException.
     */
    override var intVal: Int
        get() = java.lang.Float.NaN.toInt()
        set(v) = Logger.logErr("Setting int Val on generic XLSRecord, value not held")

    /**
     * Get the value of the record as a Double.
     * Value must be parseable as an Double or it
     * will throw a NumberFormatException.
     */
    override val dblVal: Double
        get() = java.lang.Float.NaN.toDouble()

    /**
     * Get the value of the record as a Float.
     * Value must be parseable as an Float or it
     * will throw a NumberFormatException.
     */
    override var floatVal: Float
        get() = java.lang.Float.NaN
        set(v) = Logger.logErr("Setting float Val on generic XLSRecord, value not held")

    /**
     * Get the value of the record as a String.
     */
    override var stringVal: String?
        get() = null
        set(v) = Logger.logErr("Setting String Val on generic XLSRecord, value not held")

    override var byteReader: BlockByteReader?
        get() = databuf
        set(db) {
            databuf = db
            data = null
        }


    /**
     * Gets the byte from the specified position in the
     * record byte array.
     *
     * @param off
     * @return
     */
    override val bytes: ByteArray?
        get() = this.getBytesAt(0, this.length)

    /**
     * Returns the length of this
     * record, including the 4 header bytes
     */
    override// a new rec
    // returns the original len always
    // returns updated lengths
    var length: Int
        get() {
            if (data != null)
                return data!!.size + 4
            else if (this.databuf == null)
                return -1
            if (hasContinues() &&
                    !isContinueMerged &&
                    this.opcode != XLSConstants.SST &&
                    this.opcode != XLSConstants.SXLI) {
                if (this.reclen > XLSConstants.MAXRECLEN) {
                    setData(this.getBytesAt(0, XLSConstants.MAXRECLEN))
                } else {
                    setData(this.getBytesAt(0, this.reclen))
                }
                mergeContinues()
                return data!!.size + 4
            }
            return this.reclen + 4
        }
        set(len) {
            if (this.originalsize <= 0) this.originalsize = len
            this.reclen = len
        }

    /**
     * Get the color table for the associated workbook.  If the workbook is null
     * then the default colortable will be returned.
     */
    val colorTable: Array<java.awt.Color>
        get() {
            try {
                return this.workBook!!.colorTable
            } catch (e: Exception) {
                return FormatHandle.COLORTABLE
            }

        }

    open fun shouldEncrypt(): Boolean {
        return true
    }


    /**
     * Removes this BiffRec from the WorkSheet
     *
     * @param whether to nullify this Cell
     */
    override fun remove(nullme: Boolean): Boolean {
        val success = false
        if (worksheet != null && isValueForCell)
            sheet!!.removeCell(this)

        if (streamer != null)
            streamer!!.removeRecord(this)

        if (nullme) {
            try {
                this.finalize()
            } catch (t: Throwable) {
            }

        }
        this.worksheet = null
        return true
    }

    // Methods from BlockByteConsumer //

    /**
     * Set the relative position within the data
     * underlying the block vector represented by
     * the BlockByteReader.
     *
     *
     * In other words, this is the relative position
     * used by the BlockByteReader to offset the Consumer's
     * read position within the collection of Data Blocks.
     *
     *
     * This may be an offset relative to the data in a file,
     * or within a Storage contained in a file.
     *
     *
     * The Workbook Storage for example will contain a non-contiguous
     * collection of Blocks containing data from any number of
     * positions in a file.
     *
     *
     * This collection forms a contiguous span of bytes comprising
     * an XLS Workbook.  The XLSRecords within this span of bytes will
     * set their relative position within this 'virtual' array.  Thus
     * the XLSRecord positions are relative to the order of bytes contained
     * in the Block collection.  The BOF record then is at offset 0 within the
     * data of the first Block, even though the underlying data of this
     * first Block may be anywhere on disk.
     *
     * @param pos
     */
    override fun setOffset(pos: Int) {
        if (originalOffset < 1) originalOffset = pos
        offset = pos
    }

    /**
     * Get the relative position within the data
     * underlying the block vector represented by
     * the BlockByteReader.
     *
     * @return relative position
     */
    override fun getOffset(): Int {
        return if (this.data == null) this.originalOffset else offset
    }


    protected fun initRowCol() {
        var pos = 0

        val bt = this.getBytesAt(pos, 2)
        rw = ByteTools.readUnsignedShort(bt!![0], bt[1])
        pos += 2
        colNumber = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt())
    }


    override fun toString(): String {
        return recDesc
    }

    /**
     * Dumps this record as a human-readable string.
     */
    override fun toHexDump(): String {
        return recDesc + "\n" + ByteTools.getByteDump(this.getData(), 0)
    }

    /**
     * Copy all formatting info from source biffrec
     *
     * @see io.starter.formats.XLS.BiffRec.copyFormat
     */
    override fun copyFormat(source: BiffRec) {
        val clone = source.xfRec.clone() as Xf
        val fontClone = source.xfRec.font!!.clone() as Font

        Logger.logInfo(source.toString() + ":" + source.xfRec + ":" + clone)
        var fid = -1
        var xid = -1
        // see if we have an equivalent Xf/Font combo
        if (this.workBook!!.fontRecs.contains(fontClone)) {
            fid = this.workBook!!.fontRecs.indexOf(fontClone)
        }
        if (this.workBook!!.xfrecs.contains(clone)) {
            xid = this.workBook!!.xfrecs.indexOf(clone)
        }

        // add the xf/font and set the ixfe
        this.workBook!!.addRecord(clone, false)
        this.workBook!!.addRecord(fontClone, false)
        clone.setFont(fontClone.idx)
        this.setXFRecord(clone.idx)

    }

    /**
     * clone a record
     */
    override fun clone(): Any? {
        try {
            val cn = javaClass.getName()
            val rec = Class.forName(cn).newInstance() as XLSRecord
            val inb = bytes
            rec.setData(inb)
            rec.streamer = this.streamer
            rec.workBook = workBook
            rec.opcode = opcode
            rec.length = length
            rec.setSheet(sheet)    //20081120 KSC: otherwise may set sheet incorrectly in init
            rec.init()
            return rec
        } catch (e: Exception) {
            Logger.logInfo("cloning XLSRecord " + this.cellAddress + " failed: " + e)
        }

        return null
    }

    /** methods from CompatibleVectorHints
     *
     * protected transient int recordIdx = -1;
     */
    /** provide a hint to the CompatibleVector
     * about this objects likely position.
     *
     * public int getRecordIndexHint(){return recordIdx;}
     */

    /** set index information about this
     * objects likely position.
     *
     * public void setRecordIndexHint(int i){
     * lastidx = i;
     * recordIdx = i;
     * }
     */
    /**
     * set the DEBUG level
     */
    override fun setDebugLevel(b: Int) {
        DEBUGLEVEL = b
    }

    /**
     * adds a CONTINUE record to the array of CONTINUE records
     * for this record, containing all data
     * beyond the 8224 byte record size limit.
     */
    override fun addContinue(c: Continue) {
        if (continues == null) continues = ArrayList()
        continues!!.add(c)
    }

    /**
     * remove all Continue records
     */
    override fun removeContinues() {
        if (continues != null) continues!!.clear()
    }

    /**
     * returns whether this record has a CONTINUE
     * record containing data beyond the 8228 byte
     * record size limit.
     *
     *
     * XLSRecords can have 0 or more CONTINUE records.
     */
    override fun hasContinues(): Boolean {
        return if (continues == null) false else this.continues!!.size > 0
    }

    /**
     * associate this record with its Index record
     */
    override fun setIndex(id: Index) {
        idx = id
    }

    /**
     * Associate this record with a worksheet.
     * First checks to see if there is already
     * a cell with this address.
     */
    override fun setSheet(b: Sheet?) {
        this.worksheet = b
    }

    /**
     * set the column
     */
    override fun setCol(i: Short) {
        val c = ByteTools.shortToLEBytes(i)
        System.arraycopy(c, 0, getData()!!, 2, 2)
        colNumber = i
    }

    override fun setRowCol(x: IntArray) {
        this.rowNumber = x[0]
        this.setCol(x[1].toShort())
    }

    /**
     * perform record initialization
     */
    override fun init() {
        if (originalsize == 0) originalsize = reclen
    }


    /**
     * Get the value of the record as a String.
     */
    override fun getStringVal(encoding: String): String? {
        return null
    }

    override fun setDoubleVal(v: Double) {
        Logger.logErr("Setting Double Val on generic XLSRecord, value not held")
    }

    /**
     * do any pre-streaming processing such as expensive
     * index updates or other deferrable processing.
     */
    override fun preStream() {
        // override in sub-classes
    }

    /**
     * set the XF (format) record for this rec
     */
    override fun setXFRecord() {
        if (wkbook == null)
            return
        if (ixfe > -1 && ixfe < wkbook!!.numXfs) {
            if (myxf == null || myxf!!.idx != ixfe) {
                this.myxf = wkbook!!.getXf(this.ixfe)
                this.myxf!!.incUseCount()
            }
        }
    }

    /**
     * set the XF (format) record for this rec
     */
    override fun setXFRecord(i: Int) {
        if (i != ixfe || myxf == null) {
            this.setIxfe(i)
            this.setXFRecord()
        }
    }

    /**
     * set the XF (format) record for this rec
     */
    override fun setIxfe(i: Int) {
        this.ixfe = i
        val newxfe = ByteTools.cLongToLEBytes(i)
        val b = this.getData()
        if (b != null)
            System.arraycopy(newxfe, 0, b, 4, 2)
        this.setData(b)
    }

    /**
     * get the ixfe
     */
    override fun getIxfe(): Int {
        return this.ixfe
    }

    override fun setData(b: ByteArray?) {
        data = b
        this.databuf = null
    }

    /**
     * gets the record data merging any Continue record
     * data.
     */
    override fun getData(): ByteArray? {
        var len = 0
        if ((len = this.length) == 0)
            return byteArrayOf()
        if (data != null)
            return data
        if (len > XLSConstants.MAXRECLEN) {
            setData(this.getBytesAt(0, XLSConstants.MAXRECLEN))
        } else {
            setData(this.getBytesAt(0, len - 4))
        }

        if (this.opcode == XLSConstants.SST || this.opcode == XLSConstants.SXLI)
            return data

        if (!isContinueMerged && hasContinues()) {
            mergeContinues()
        }
        return data
    }

    /**
     * Merge continue data in to the record data array
     */
    fun mergeContinues() {
        val cx = this.continueVect
        if (cx != null) {
            val out = ByteArrayOutputStream()
            // get the main bytes first!!!
            try {
                out.write(data!!)
            } catch (e: IOException) {
                Logger.logWarn(
                        "ERROR: parsing record continues failed: "
                                + this.toString()
                                + ": "
                                + e)
            }

            val it = cx.iterator()
            while (it.hasNext()) {
                val c = it.next() as Continue
                var nb = c.data
                if (c.hasGrbit) {    // remove it - happens in continued StringRec ... a rare case
                    val newnb = ByteArray(nb!!.size - 1)
                    System.arraycopy(nb, 1, newnb, 0, newnb.size)
                    nb = newnb
                }
                try {
                    out.write(nb!!)
                } catch (a: IOException) {
                    Logger.logWarn(
                            "ERROR: parsing record continues failed: "
                                    + this.toString()
                                    + ": "
                                    + a)
                }

            }
            this.data = out.toByteArray()
        }
        isContinueMerged = true
    }

    /**
     * Gets the byte from the specified position in the
     * record byte array.
     *
     * @param off
     * @return
     */
    override fun getBytesAt(off: Int, len: Int): ByteArray? {
        var len = len
        if (this.data != null) {
            if (len + off > data!!.size) len = data!!.size - off // deal with bad requests
            val ret = ByteArray(len)
            System.arraycopy(data!!, off, ret, 0, len)
            return ret
        }
        return if (databuf == null) null else this.databuf!!.get(this, off, len)
    }

    /** Sets a subset of temporary bytes for fast access during init methods...
     *
     *
     * public void initCacheBytes(int start, int len) {
     * data = this.getBytesAt(start, len);
     * }  */

    /**
     * resets the cache bytes so they do not take up space
     */
    fun resetCacheBytes() {
        //  data = null;
    }

    /**
     * Gets the byte from the specified position in the
     * record byte array.
     *
     * @param off
     * @return
     */
    override fun getByteAt(off: Int): Byte {
        if (this.data != null) {
            return data!![off]
        }
        if (this.databuf == null)
            throw InvalidRecordException("XLSRecord has no data buffer." + this.cellAddress)
        return this.databuf!!.get(this, off)
    }

    /**
     * @param isVal true if this is a cell-type record
     */
    fun setValueForCell(isVal: Boolean) {
        this.isValueForCell = isVal
    }

    override fun postStream() {
        // nothing here -- use to blow out data
    }

    /**
     * clear out object references in prep for closing workbook
     */
    open fun close() {
        if (databuf != null) {
            databuf!!.clear()
            databuf = null
        }
        idx = null
        worksheet = null
        myxf = null
        wkbook = null
        streamer = null
        if (continues != null) {
            for (i in continues!!.indices)
                (continues!![i] as XLSRecord).close()
            continues!!.clear()
        }
        if (hyperlink != null) {
            hyperlink!!.close()
            hyperlink = null
        }
        mergeRange = null
        if (myrow != null) {
            myrow = null
        }
        mergeRange = null

    }

    companion object {

        private const val serialVersionUID = -106915096753184441L


        /**
         * get a new, generic instance of a Record.
         */
        protected val prototype: XLSRecord?
            get() {
                Logger.logWarn("Attempt to get prototype XLSRecord failed.  There is no prototype record defined for this record type.")
                return null
            }
    }


}