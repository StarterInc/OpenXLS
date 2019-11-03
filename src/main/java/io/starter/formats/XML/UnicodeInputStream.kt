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
/**
 * BOMInputstream.java
 *
 *
 *
 *
 * Oct 14, 2010
 */
package io.starter.formats.XML

/**
 * Original pseudocode   : Thomas Weidenfeller
 * Implementation tweaked: Aki Nieminen
 *
 * http://www.unicode.org/unicode/faq/utf_bom.html
 * BOMs:
 * 00 00 FE FF    = UTF-32, big-endian
 * FF FE 00 00    = UTF-32, little-endian
 * FE FF          = UTF-16, big-endian
 * FF FE          = UTF-16, little-endian
 * EF BB BF       = UTF-8
 *
 * Win2k Notepad:
 * Unicode format = UTF-16LE
 */

import java.io.IOException
import java.io.InputStream
import java.io.PushbackInputStream

/**
 * This inputstream will recognize unicode BOM marks
 * and will skip bytes if getEncoding() method is called
 * before any of the read(...) methods.
 *
 * Usage pattern:
 * String enc = "ISO-8859-1"; // or NULL to use
 * systemdefault
 * FileInputStream fis = new FileInputStream(file);
 * UnicodeInputStream uin = new UnicodeInputStream(fis,
 * enc);
 * enc = uin.getEncoding(); // check for BOM mark and skip
 * bytes
 * InputStreamReader in;
 * if (enc == null) in = new InputStreamReader(uin);
 * else in = new InputStreamReader(uin, enc);
 */
class UnicodeInputStream(`in`: InputStream, defaultEnc: String) : InputStream() {
    internal var internalIn: PushbackInputStream
    internal var isInited = false
    var defaultEncoding: String
        internal set
    internal var encoding: String

    init {
        internalIn = PushbackInputStream(`in`, BOM_SIZE)
        this.defaultEncoding = defaultEnc
    }

    fun getEncoding(): String {
        if (!isInited) {
            try {
                init()
            } catch (ex: IOException) {
                throw IllegalStateException("Init method failed.$ex")
                //              (Throwable)ex);
            }

        }
        return encoding
    }

    /**
     * Read-ahead four bytes and check for BOM marks. Extra
     * bytes are
     * unread back to the stream, only BOM bytes are skipped.
     */
    @Throws(IOException::class)
    protected fun init() {
        if (isInited) return

        val bom = ByteArray(BOM_SIZE)
        val n: Int
        val unread: Int
        n = internalIn.read(bom, 0, bom.size)

        if (bom[0] == 0xEF.toByte() && bom[1] == 0xBB.toByte()
                &&
                bom[2] == 0xBF.toByte()) {
            encoding = "UTF-8"
            unread = n - 3
        } else if (bom[0] == 0xFE.toByte() && bom[1] == 0xFF.toByte()) {
            encoding = "UTF-16BE"
            unread = n - 2
        } else if (bom[0] == 0xFF.toByte() && bom[1] == 0xFE.toByte()) {
            encoding = "UTF-16LE"
            unread = n - 2
        } else if (bom[0] == 0x00.toByte() && bom[1] == 0x00.toByte() &&
                bom[2] == 0xFE.toByte() && bom[3] == 0xFF.toByte()) {
            encoding = "UTF-32BE"
            unread = n - 4
        } else if (bom[0] == 0xFF.toByte() && bom[1] == 0xFE.toByte() &&
                bom[2] == 0x00.toByte() && bom[3] == 0x00.toByte()) {
            encoding = "UTF-32LE"
            unread = n - 4
        } else {
            // Unicode BOM mark not found, unread all bytes
            encoding = defaultEncoding
            unread = n
        }
        // io.starter.toolkit.Logger.log("read=" + n + ", unread=" + unread);

        if (unread > 0)
            internalIn.unread(bom, n - unread,
                    unread)

        isInited = true
    }

    @Throws(IOException::class)
    override fun close() {
        //init();
        isInited = true
        internalIn.close()
    }

    @Throws(IOException::class)
    override fun read(): Int {
        //init();
        isInited = true
        return internalIn.read()
    }

    companion object {

        private val BOM_SIZE = 4
    }
}
