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

import io.starter.formats.OOXML.Text
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList


/**
 * **Note: A cell annotation (1CDh)**<br></br>
 *
 *
 * Note records contain the row/col position of the annotation, plus an ordinal id that links it to the
 * set of records that define it:  Mso/obj/mso/Txo/Continue/Continue.  These record sets appear one after
 * another after the DIMENSIONS record and before WINDOW2.  After the associated records, all Note records
 * appear, in order.
 *
 *
 *
 *
 * Kaia's notes on notes:
 *
 *
 * The Note record is the last of a set of records necessary to define each note:
 * Mso - shapeType= msosptTextBox.  See Msodrawing.createTextBoxStyle for a list of the necessary sub-records
 * Obj containing an ftNts sub-record
 * Mso - odd Mso cotnaining only 8 bytes - see Msodrawing.getAttachedTextPrototype
 * Txo - contains the size of the text + options such as text rotation + size of the formatting runs
 * Continue - 0 + text
 * Continue - formatting runs for text
 *
 *
 * For each note, the above records repeat, then each Note record appears in order after all of the above recs
 * Notes appear to have an ordinal id that is necessary for proper display
 * It's hidden/shown state appears to be defined by byte 5 but there is also other necessary info that I don't know yet
 *
 *
 * The tricky part of Notes is the Mso msosptTextBox plus the Mso msosptClientTextBox; they don't follow the
 * normal rules we are used to - the SPCONTAINERLENGTH of the 1st Mso MUST BE inc by 8, the length of the
 * second Mso ... but when multiple Notes are on the same sheet, that logic doesn't seem to be followed 100%
 * plus the number of shapes of the mso header is NOT incremented by the secondary Mso ...
 *
 *
 * Also, the Mso's don't appear to follow the normal rules for the drawing Id - seem to be repeated ...
 *
 *
 * Lastly, the 1st Mso contains an OPT sub-record msofbtlTxid; this contains the host-defined id for the note;
 * I'm not sure what to set for this so I put a random int.
 *
 *
 * More info:
 *
 *
 * Some templates define notes differently; instead of the above set of records, it goes
 * Obj (ftNts)
 * Continue [00, 00, (byte)0x0D, (byte)0xF0, 00, 00, 00, 00]  --> masks the 2nd mso
 * Txo
 * Continue
 * Continue
 * Continue --> masks the 1st mso
 *
 *
 * Also:
 * formatting runs: This is almost finished but - of course - doc on formatting runs format doesn't appear to match reality ((:)
 * Also:
 * There needs to be a good, easy interface for setting formatting runs for a text string
 *
 *
 *
 *
 * offset  name        size    contents
 * ---
 * 0 		2 Index to row
 * 2 		2 Index to column
 * 4		2 grbit controls hidden state - 0=hidden, 2=shown
 * 6		2 ordinal id MUST MATCH Obj record id
 * 8		1 lenght of author string or /2 if encoding=1
 * 9       1 encoding 0= non-unicode? 1= unicode?
 * 10		(var) 0 + author bytes
 * (author encoding)
 *
 *
 * @see Note
 */
class Note : io.starter.formats.XLS.XLSRecord() {


    /**
     *
     */
    private var id: Short = 0
    internal var hidden = true    // default
    private var author: String? = null
    private var auth_encoding: Byte = 0
    private var mso: MSODrawing? = null
    // pointer to associated Txo, stores actual text + formatting runs
    internal var txo: Txo? = null

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, /* row */
            0, 0, /* col */
            0, /* hidden state */
            0, /* unknown */
            1, 0, /* ordinal id */
            4, /* length of author */
            0)/* string encoding*///		/* 0-padded author string*/

    /**
     * returns the Text of this Note or Comment
     *
     * @return String
     */
    /**
     * sets the Text of this Note or Comment
     *
     * @param txt
     */
    var text: String?
        get() = if (txo != null) txo!!.stringVal else null
        set(txt) {
            if (txo != null) {
                try {
                    txo!!.stringVal = txt
                } catch (e: IllegalArgumentException) {
                    Logger.logErr(e.toString())
                }

            }
        }

    /**
     * returns the Txo associated with this Note
     *
     * @return
     */
    private// shouldn't!
    // should't!
    // if it's of type Note + has the same id, this is it
    // got it!
    // now find the next TXO
    val associatedTxo: Txo?
        get() {
            val bs = this.sheet
            var idx = -1
            if (bs != null) {
                idx = bs.getIndexOf(XLSConstants.OBJ)
                if (idx == -1) return null
                while (idx < bs.sheetRecs.size) {
                    if ((bs.sheetRecs[idx] as BiffRec).opcode == XLSConstants.OBJ) {
                        val o = bs.sheetRecs[idx] as Obj
                        if (o.objType == 0x19 && o.objId == this.id.toInt()) {
                            idx++
                            while (idx < bs.sheetRecs.size && (bs.sheetRecs[idx] as BiffRec).opcode != XLSConstants.TXO)
                                idx++
                            break
                        }
                    }
                    idx++
                }
            }
            return if (idx < bs!!.sheetRecs.size) bs.sheetRecs[idx] as Txo else null
        }

    /**
     * returns the bounds (size and position) of the Text Box for this Note
     * <br></br>bounds are relative and based upon rows, columns and offsets within
     * <br></br>bounds are as follows:
     * <br></br>bounds[0]= column # of top left position (0-based) of the shape
     * <br></br>bounds[1]= x offset within the top-left column (0-1023)
     * <br></br>bounds[2]= row # for top left corner
     * <br></br>bounds[3]= y offset within the top-left corner	(0-1023)
     * <br></br>bounds[4]= column # of the bottom right corner of the shape
     * <br></br>bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
     * <br></br>bounds[6]= row # for bottom-right corner of the shape
     * <br></br>bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
     *
     * @return
     */
    /**
     * sets the bounds (size and position) of the Text Box for this Note
     * <br></br>bounds are relative and based upon rows, columns and offsets within
     * <br></br>bounds are as follows:
     * <br></br>bounds[0]= column # of top left position (0-based) of the shape
     * <br></br>bounds[1]= x offset within the top-left column (0-1023)
     * <br></br>bounds[2]= row # for top left corner
     * <br></br>bounds[3]= y offset within the top-left corner	(0-1023)
     * <br></br>bounds[4]= column # of the bottom right corner of the shape
     * <br></br>bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
     * <br></br>bounds[6]= row # for bottom-right corner of the shape
     * <br></br>bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
     *
     * @param bounds
     */
    var textBoxBounds: ShortArray?
        get() {
            if (mso == null)
                mso = this.associatedMso
            return mso!!.getBounds()
        }
        set(bounds) {
            if (mso == null)
                mso = this.associatedMso
            mso!!.setBounds(bounds)
        }

    /**
     * Method to retrieve formatting runs for this Note (Comment)
     * <br></br>Formatting Runs are an Internal Structure and are not relavent to the end user
     *
     * @param formattingRuns
     */
    /**
     * Store formatting runs for this Note (Comment)
     * <br></br>Formatting Runs are an Internal Structure and are not relavent to the end user
     *
     * @param formattingRuns
     */
    var formattingRuns: ArrayList<*>?
        get() {
            if (txo != null)
                txo!!.formattingRuns
            return null
        }
        set(formattingRuns) {
            if (txo != null)
                txo!!.formattingRuns = formattingRuns
        }

    /**
     * get the id of drawing object which defines the text box for this Note
     * <br></br>For Internal Use Only
     *
     * @return
     */
    val spid: Int
        get() {
            if (mso == null)
                mso = this.associatedMso
            return mso!!.spid
        }

    private// shouldn't!
    // should't!
    // if it's of type Note + has the same id, this is it
    // got it!
    // first check if this is one of the odd configurations
    // normal case: mso, obj (note), mso, txo, continue, continue
    // continue masking mso - AFTER object record
    // the 1st mso is actually at position 5: Obj (note)/Continue/Txo/Continue/Continue/Continue
    val associatedMso: MSODrawing?
        get() {
            val bs = this.sheet
            var idx = -1
            if (bs != null) {
                idx = bs.getIndexOf(XLSConstants.OBJ)
                if (idx == -1) return null
                while (idx < bs.sheetRecs.size) {
                    if ((bs.sheetRecs[idx] as BiffRec).opcode == XLSConstants.OBJ) {
                        val o = bs.sheetRecs[idx] as Obj
                        if (o.objType == 0x19 && o.objId == this.id.toInt()) {
                            val opcodeprev = (bs.sheetRecs[idx - 1] as BiffRec).opcode.toInt()
                            val opcodenext = (bs.sheetRecs[idx + 1] as BiffRec).opcode.toInt()
                            if (opcodeprev == XLSConstants.MSODRAWING.toInt())
                                return bs.sheetRecs[idx - 1] as MSODrawing
                            else if (opcodenext == XLSConstants.CONTINUE.toInt()) {
                                if (idx + 5 < bs.sheetRecs.size)
                                    return (bs.sheetRecs[idx + 5] as Continue).maskedMso
                            }
                        }
                    }
                    idx++
                }
            }
            return null
        }

    /**
     * returns the OOXML representation of this Note
     *
     * @return
     */
    val ooxml: String
        get() {
            if (txo == null) return ""
            val t = Text(Sst.createUnicodeString(this.text, txo!!.formattingRuns, Sst.STRING_ENCODING_UNICODE))
            return t.getOOXML(this.workBook)
        }

    override fun init() {
        super.init()
        rw = ByteTools.readShort(this.getData()!![0].toInt(), this.getData()!![1].toInt()).toInt()
        colNumber = ByteTools.readShort(this.getData()!![2].toInt(), this.getData()!![3].toInt())
        hidden = this.getData()!![4] != 2.toByte()    // not entirely sure of this
        // bytes 5 unknnown
        id = ByteTools.readShort(this.getData()!![6].toInt(), this.getData()!![7].toInt())
        // rest are fairy known :)
        val authorlen = this.getData()!![8].toShort()
        auth_encoding = this.getData()!![9]
        val authorbytes = ByteArray(authorlen)
        System.arraycopy(this.getData()!!, 11, authorbytes, 0, authorlen.toInt())
        if (auth_encoding.toInt() == 0)
            author = String(authorbytes)
        else
            try {
                author = String(authorbytes, WorkBookFactory.UNICODEENCODING)
            } catch (e: Exception) {
            }

    }

    override fun setSheet(bs: Sheet?) {
        super.setSheet(bs)
        txo = associatedTxo
    }

    /**
     * Is this a useful to string?  possibly just the note itself, at a minimum include it
     *
     * @see io.starter.formats.XLS.XLSRecord.toString
     */
    override fun toString(): String {
        var s: String? = ""
        if (txo != null)
            s = txo!!.stringVal
        if (author != null)
            s += " author:" + author!!
        return "NOTE at [" + this.cellAddressWithSheet + "]: " + s
    }

    /**
     * set the row and column that this note is attached to
     *
     * @param row
     * @param col
     */
    fun setRowCol(row: Int, col: Int) {
        this.rw = row.toShort().toInt()
        this.colNumber = col.toShort()
        var b = ByteTools.shortToLEBytes(this.rw.toShort())
        this.getData()[0] = b[0]
        this.getData()[1] = b[1]
        b = ByteTools.shortToLEBytes(this.colNumber)
        this.getData()[2] = b[0]
        this.getData()[3] = b[1]
    }

    /**
     * return the ordinal id of this note
     *
     * @param id
     */
    fun setId(id: Int) {
        this.id = id.toShort()
        val b = ByteTools.shortToLEBytes(this.id)
        this.getData()[6] = b[0]
        this.getData()[7] = b[1]
    }

    /**
     * retrieve the ordinal id of this note
     *
     * @return
     */
    fun getId(): Int {
        return this.id.toInt()
    }

    /**
     * returns true if this note is hidden (default state)
     *
     * @return
     */
    fun getHidden(): Boolean {
        return this.hidden
    }

    /**
     * show or hide this note
     * <br></br>NOTE: this is still experimental
     *
     * @param hidden
     */
    fun setHidden(hidden: Boolean) {
        this.hidden = hidden
        if (mso == null)
            mso = this.associatedMso
        if (this.hidden) {    // hide
            this.getData()[4] = 0
            // ALSO MUST SET the associated Mso opt subrec to actually hide the note textbox ...
            mso!!.setOPTSubRecord(MSODrawingConstants.msooptGroupShapeProperties, false, false, 131074, null)
        } else {            //show
            this.getData()[4] = 2
            // ALSO MUST SET the associated Mso opt subrec to actually show the note textbox ...
            mso!!.setOPTSubRecord(MSODrawingConstants.msooptGroupShapeProperties, false, false, 131072, null)
        }
    }

/**
 * /** set the text of this Note (Comment) as a unicode string,
 * with formatting information
 *
 * @param txt
*/
fun setText(txt:Unicodestring) {
if (txo != null)
{
try
{
txo!!.setStringVal(txt)
}
catch (e:IllegalArgumentException) {
Logger.logErr(e.toString())
}

}

}

/**
 * return the author of this note, if set
 *
 * @return
*/
fun getAuthor():String? {
return author
}

/**
 * sets the author of this note
 *
 * @param author
*/
fun setAuthor(author:String) {
this.author = author
val authbytes = this.author!!.toByteArray()
val oldData = this.getData()
val newData = ByteArray(authbytes.size + 11)
System.arraycopy(oldData!!, 0, newData, 0, 8)
newData[8] = author.length.toByte()
// encoding= 0
System.arraycopy(authbytes, 0, newData, 11, authbytes.size)
this.setData(newData)
this.init()
}

/**
 * sets the text box width for the note in pixels
 *
 * @param width
 * @return short[] end column and end column offset
*/
fun setTextBoxWidth(width:Short) {
if (mso == null)
mso = this.associatedMso
mso!!.setWidth(Math.round(width / 6.4).toShort().toInt())    // convert pixels to points
}


/**
 * sets the text box width for the note in pixels
 *
 * @param width
 * @return short[] end column and end column offset
*/
fun setTextBoxHeight(height:Short) {
if (mso == null)
mso = this.associatedMso
mso!!.setHeight(Math.round(height.toFloat()).toShort().toInt())
}

companion object {
private val serialVersionUID = -1571461658267879478L

fun getPrototype(author:String):XLSRecord {
val n = Note()
n.opcode = XLSConstants.NOTE
/*
 *  $row, $col, $visible, $obj_id,
$num_chars, $author_enc
$length     = length($data) + length($author);
my $header  = pack("vv", $record, $length);
*/
val authbytes = author.toByteArray()
val data = ByteArray(n.PROTOTYPE_BYTES.size + authbytes.size + 1)
data[8] = author.length.toByte()
// encoding= 0
System.arraycopy(authbytes, 0, data, 11, authbytes.size)
n.setData(data)
n.init()
return n
}
}
}
