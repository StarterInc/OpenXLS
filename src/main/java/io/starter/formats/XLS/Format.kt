/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.formats.XLS

import io.starter.toolkit.ByteTools
import io.starter.toolkit.StringTool

/**
 * Stores a custom number format pattern.
 */
class Format() : XLSRecord() {

    private var ifmt: Short = -1
    /**
     * Gets the format pattern string represented by this Format.
     */
    var format: String? = null
        private set
    private var ustring: Unicodestring? = null

    override/*
         * Not sure why this is here, but I think it has something
         * to do with
         * worksheet cloning. It's harmless, so might as well leave
         * it alone.
         * - Sam
         */ var workBook: WorkBook?
        get
        set(book) {
            super.workBook = book
            if (ifmt.toInt() != -1)
                book!!.addFormat(this)
        }

    init {
        opcode = XLSConstants.FORMAT
    }

    /**
     * Makes a new Format record for the given pattern.
     *
     * @param book    the workbook to which the pattern should belong
     * @param pattern the number format pattern to ensure exists
     */
    constructor(book: WorkBook, pattern: String) : this(book, (-1).toShort(), pattern) {}

    /**
     * Makes a new Format record with the given ID and pattern.
     * This should only be used when parsing non-BIFF8 files. BIFF8 parsing
     * will use the normal XLSRecord init sequence. For programmatic creation
     * of custom formats use [.Format] instead.
     *
     * @param book    the workbook to which the pattern should belong
     * @param id      the format ID to use or -1 to generate one
     * @param pattern the number format pattern to ensure exists
     */
    constructor(book: WorkBook, id: Short, pattern: String) : this() {
        workBook = book

        this.format = pattern
        this.ustring = Sst
                .createUnicodeString(pattern, null, WorkBook.STRING_ENCODING_AUTO)

        val idbytes: ByteArray
        if (id > 0) {
            this.ifmt = id
            idbytes = ByteTools.shortToLEBytes(id)
        } else {
            // WorkBook.insertFormat will call setIfmt
            idbytes = ByteArray(2)
        }

        setData(ByteTools.append(ustring!!.read(), idbytes))

        book.insertFormat(this)
    }

    /**
     * Initializes the record from bytes.
     * This method should only be called as part of the normal XLSRecord
     * init sequence when parsing from bytes.
     *
     * @throws IllegalStateException if the record has already been parsed
     */
    override fun init() {
        if (format != null)
            throw IllegalStateException(
                    "the record has already been parsed")

        super.init()

        ifmt = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())

        ustring = Unicodestring()
        ustring!!.init(this.getBytesAt(2, this.length - 2)!!, false)

        format = ustring!!.toString()

        /*
         * Strip double quotes and backslashes from the format
         * string.
         * The quoting characters are an integral part of the format
         * string,
         * so this is almost certainly wrong. However, it's what the
         * previous
         * implementation did and I'm trying to preserve behavior.
         * TODO: revisit stripping of quotes when writing number
         * format parser
         * - Sam
         */
        format = StringTool.replaceText(format!!, "\"", "", 0)
        format = StringTool.replaceText(format!!, "\\", "", 0)

        this.workBook!!.addFormat(this)
    }

    /**
     * Sets the format ID of this Format record.
     */
    fun setIfmt(id: Short) {
        ifmt = id

        // Update the record bytes
        System.arraycopy(ByteTools.shortToLEBytes(ifmt), 0, this
                .getData()!!, 0, 2)
    }

    /**
     * Gets the format index of this Format in its workbook.
     */
    fun getIfmt(): Short {
        return ifmt
    }

    override fun toString(): String? {
        return format
    }

    fun init(workBook: WorkBook, fID: Int, format: String) {
        // TODO Auto-generated method stub

    }

    companion object {
        private val serialVersionUID = 1199947552103220748L
    }
}