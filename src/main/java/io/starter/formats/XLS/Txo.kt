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

import java.io.UnsupportedEncodingException
import java.util.ArrayList


/**
 * **Txo: Text Object (1B6h)**<br></br>
 * This record stores a text object.  This record is followed
 * by two CONTINUE records which contain the text data and the
 * formatting runs respectively.
 *
 *
 * If there is no text, the two CONTINUE records are absent.
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       grbit       2       Option flags.  See table.
 * 6       rot         2       Orientation of text within the object.
 * 8       Reserved    6       Must be zero.
 * 14      cchText     2       Length of text in first CONTINUE rec.
 * 16      cbRuns      2       Length of formatting runs in second CONTINUE rec.
 * 18      Reserved    4       Must be zero.
 *
 *
 * The grbit field contains the following option flags:
 *
 * bits    mask    name                contents
 * ----
 * 0       0x01    Reserved
 * 3-1     0x0e    alcH                Horizontal text alignment
 * 1 = left
 * 2 = centered
 * 3 = right
 * 4 = justified
 * 6-4     0x70    alcV                Vertical text alignment
 * 1 = top
 * 2 = center
 * 3 = bottom
 * 4 = justified
 * 8-7     0x180   Reserved
 * 9       0x200   fLockText           1 = lock text option is on
 * 15-10   0xfc00  Reserved
 *
 *
 * The first CONTINUE record contains text -- the length is cchText of this object.
 * The first byte of the CONTINUE record's body data is 0x0 = compressed unicode.  The rest is text.
 *
</pre> *
 *
 * @see LABELSST
 *
 * @see STRING
 *
 * @see CONTINUE
 */

class Txo : io.starter.formats.XLS.XLSRecord() {
    internal var text: Continue? = null
    internal var formattingruns: Continue? = null //20100430 KSC: garbagetxo is really a masked mso, garbagetxo;  // garbagetxo is a third continue that appears to be cropping up in infoteria files.  We are removing from the file stream currently, but may need to integrate
    internal var state = 0
    internal var grbit: Short = 0
    internal var cchText: Short = 0
    internal var cbRuns: Short = 0
    internal var rot: Short = 0
    internal var compressedUnicode = false


    /**
     * returns the String value of this Text Object
     */
    /**
     * sets the String value of this Text Object
     * <br></br>if present, will parse embedded formats within the string as:
     * <br></br>the format of the embedded information is as follows:
     * <br></br>&lt;font specifics>text segment&lt;font specifics for next segment>text segment...
     * <br></br>where font specifics can be one or more of:
     * b		if present, bold
     * <br></br>i		if present, italic
     * <br></br>s		if present, strikethru
     * <br></br>u		if present, underlined
     * <br></br>f=""		font name e.g. "Arial"
     * <br></br>sz=""	font size in points e.g. "10"
     * <br></br>delimited by ;'s
     * <br></br>For Example:
     * <br></br>&lt;f="Tahoma";b;sz="16">Note: &lt;f="Tahoma";sz="12">This is an important point
     *
     * @throws IllegalArgumentException if String is incorrect format
     */
    // TODO: if length of string is > 8218, must have another continues ******************
    override// normal case (Default encoding)
    // encoding=1  (are there other encoding options??
    // extracts text string from formats and sets formatting runs
    // no formatting present:
    // reset formatting runs
    // get the length of the first CONTINUE
    // TODO: checked for Compressed UNICODE in text CONTINUE
    // create new text-type continues - though should be already set (see getPrototype)
    // reset text length
    var stringVal: String?
        get() {
            var s = ""
            if (text == null) return null
            val barr: ByteArray
            val encoding = text!!.data!![0].toInt()
            barr = ByteArray(text!!.data!!.size - 1)
            System.arraycopy(text!!.data!!, 1, barr, 0, barr.size)
            try {
                if (encoding == 0) {
                    s = String(barr, WorkBookFactory.DEFAULTENCODING)
                } else
                    s = String(barr, WorkBookFactory.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("reading Text Object failed: $e")
            }

            return s
        }
        @Throws(IllegalArgumentException::class)
        set(v) {
            var v = v
            if (v != null && v.indexOf('<') >= 0)
                v = parseFormatting(v)
            else
                this.formattingRuns = null
            val a = v!!.toByteArray()
            var b = ByteArray(a.size + 1)
            System.arraycopy(a, 0, b, 1, a.size)
            if (DEBUGLEVEL > 1) Logger.logInfo("Txo CHANGING: " + this.stringVal!!)
            b[0] = 0x0
            if (text != null)
                text!!.setData(b)
            else
                text = Continue.getTextContinues(v)
            b = ByteTools.shortToLEBytes(a.size.toShort())
            this.getData()[10] = b[0]
            this.getData()[11] = b[1]
            cchText = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
            if (DEBUGLEVEL > 1) Logger.logInfo(" TO: " + this.stringVal!!)
        }

    /**
     * get the formatting runs - fonts per character index, basically, for this text object,
     * as an arraylist of short[] {char index, font index}
     *
     *
     * NOTE: formatting runs in actuality differ from doc:
     * apparently each formatting run in the Continue record occupies 8 bytes (not 4)
     * plus there is an additional entry appended to the end, that indicates the last char index of the
     * string (this last entry is added by Excel 2003 in the Continues Record and is not present in Unicode Strings
     */
    /**
     * set the formatting runs - fonts per character index, basically, for this text object
     * as an arraylist of short[] {char index, font index}
     *
     * @param formattingRuns // NOTES: apparently must have a minimum of 24 bytes in the Continue formatting run to work
     * // 		  even though each formatting run=4 bytes, must have 4 bytes of padding between runs...?
     *
     *
     * // NOTES: apparently formatting runs are different than documentation indicates:
     * // each formatting run is 8 bytes:  charIndex (2), fontIndex (2), 4 bytes (ignored)
     * // at end of formatting runs, must have an extra 8 bytes:  charIndex==length/fontIndex, 4 bytes
     * // for the purposes of this method, the extra entry must be already present
     */
    // only have the "NO FONT" entry
    // skip the 4 "reserved" bytes
    // minimum "null formatting run" is 4? apparently must always have a "null" formatting run
    // formatting runs (charindex, fontindex)*n after string data
    var formattingRuns: java.util.ArrayList<*>?
        get() {
            val formattingRuns = java.util.ArrayList()
            val frcontinues = this.sheet!!.sheetRecs.indexOf(this) + 2
            val fr = this.sheet!!.sheetRecs[frcontinues] as Continue
            val frdata = fr.data
            val nFormattingRuns = frdata!!.size / 8
            if (nFormattingRuns <= 1) return null
            var i = 0
            while (i < nFormattingRuns * 8) {
                val idx: Short
                val font: Short
                idx = ByteTools.readShort(frdata[i++].toInt(), frdata[i++].toInt())
                font = ByteTools.readShort(frdata[i++].toInt(), frdata[i++].toInt())
                formattingRuns.add(shortArrayOf(idx, font))
                i += 4
            }
            return formattingRuns
        }
        set(formattingRuns) {
            var frs = ByteArray(4)
            if (formattingRuns != null) {
                frs = ByteArray(formattingRuns.size * 8)
                for (i in formattingRuns.indices) {
                    val o = formattingRuns[i] as ShortArray
                    val charIndex = ByteTools.shortToLEBytes(o[0])
                    val fontIndex = ByteTools.shortToLEBytes(o[1])
                    System.arraycopy(charIndex, 0, frs, i * 8, 2)
                    System.arraycopy(fontIndex, 0, frs, i * 8 + 2, 2)
                }
            }
            val frcontinues = this.sheet!!.sheetRecs.indexOf(this) + 2
            val fr = this.sheet!!.sheetRecs[frcontinues] as Continue
            fr.setData(frs)
            cbRuns = frs.size.toShort()
            val b = ByteTools.shortToLEBytes(cbRuns)
            this.getData()[12] = b[0]
            this.getData()[13] = b[1]

        }

    private val PROTOTYPE_BYTES = byteArrayOf(18, 2, /* grbit= 530 */
            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, /* cch */
            4, 0, /* formatting run length */
            0, 0, 0, 0  /* reserved must be 0 */)

    override fun init() {
        super.init()
        val datalen = this.length // should be 18

        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        rot = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        cchText = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
        cbRuns = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())

        this.isValueForCell = false
        // -- not always true... does it bother anybody though?
        this.isString = true
        this.isContinueMerged = true
    }


    /**
     * String s contains formatting information; extracts text string from
     * formats and sets formatting runs
     * <br></br>the format of the embedded information is as follows:
     * <br></br>< font specifics>text segment< font specifics for next segment>text segment...
     * <br></br>where font specifics can be one or more of:
     * b		if present, bold
     * <br></br>i		if present, italic
     * <br></br>s		if present, strikethru
     * <br></br>u		if present, underlined
     * <br></br>f=""		font name e.g. "Arial"
     * <br></br>sz=""	font size in points e.g. "10"
     * <br></br>delimited by ;'s
     * <br></br>For Example:
     * <br></br>< f="Tahoma";b;sz="16">Note: < f="Tahoma";sz="12">This is an impotant point
     *
     * @param s String with formatting info
     * @return txt- string with all formatting info removed
     */
    @Throws(IllegalArgumentException::class)
    private fun parseFormatting(s: String): String {
        // parse string for formatting run info
        try {
            var informatting = false
            val txt = StringBuffer()
            var frs = ShortArray(2)
            val formattingRuns = ArrayList()
            var b: Boolean
            var it: Boolean
            var st: Boolean
            var u: Boolean
            u = false
            st = u
            it = st
            b = it
            var font = "Arial"    // default
            var sz = 10                // default
            var i = 0
            while (i < s.length) {
                val c = s[i]
                if (!informatting) {
                    if (c != '<') {
                        txt.append(c)
                    } else {
                        informatting = true
                        // initialize formatting run
                        frs = ShortArray(2)
                        frs[0] = txt.length.toShort()
                        frs[1] = 0
                        u = false
                        st = u
                        it = st
                        b = it
                        sz = 10            // default
                        font = "Arial"    // default
                    }
                } else {
                    val z = s.substring(i).split("[;>]".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                    if (z == null || z.size == 0 /*would happen on invalid formats such as <;>*/
                            || z.size == 1 && !(z[0].endsWith(">") || z[0].endsWith(";"))) {
                        txt.append('<')    // not a real embedded format
                        i--
                        informatting = false
                        i++
                        continue
                    }
                    val section = z[0]
                    // gather up font info
                    if (section == "b") {
                        b = true
                    } else if (section == "i") {
                        it = true
                    } else if (section == "u") {
                        u = true
                    } else if (section == "s") {
                        st = true
                    } else if (section.startsWith("f=")) {
                        // font name
                        font = section.substring(3)
                        font = font.substring(0, font.indexOf('"'))
                    } else if (section.startsWith("sz=")) {
                        var ssz = section.substring(4)
                        ssz = ssz.substring(0, ssz.indexOf('"'))
                        sz = Integer.valueOf(ssz).toInt()
                    }
                    i += section.length
                    if (i < s.length && s[i] == '>') {    // if got end of formatting section
                        // store formatting run
                        informatting = false
                        val f = Font(font, 400, sz * 20)// sz must be in points
                        if (b) f.bold = b
                        if (it) f.italic = it
                        if (u) f.underlined = u
                        if (st) f.stricken = st
                        var fIndex = this.workBook!!.getFontIdx(f)  // index for specific font formatting
                        if (fIndex == -1)
                        // must insert new font
                            fIndex = this.workBook!!.insertFont(f) + 1
                        frs[1] = fIndex.toShort()
                        formattingRuns.add(frs)
                    }
                }
                i++
            }
            if (formattingRuns.size > 0) {
                formattingRuns.add(shortArrayOf(txt.toString().length.toShort(), 15)) // 20100430 KSC: add "extra" formatting run -- necessary for Excel 2003
                this.formattingRuns = formattingRuns
            }
            return txt.toString()
        } catch (e: Exception) {
            throw IllegalArgumentException("Unable to parse String Pattern: $s")
        }

    }

    /**
     * sets the text for this object, including formatting information
     *
     * @param txt
     */
    fun setStringVal(txt: Unicodestring) {
        this.stringVal = txt.stringVal
        this.formattingRuns = txt.formattingRuns
    }

    override fun toString(): String? {
        return stringVal
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7043468034346138525L

        /**
         * generates a skeleton Txo with 0 for text length
         * and the minimum length for formatting runs (=2)
         */
        val prototype: XLSRecord?
            get() {
                val t = Txo()
                t.opcode = XLSConstants.TXO
                t.setData(t.PROTOTYPE_BYTES)
                t.init()
                t.text = Continue.getTextContinues("")
                return t
            }
    }

}