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

import java.io.Serializable
import java.io.UnsupportedEncodingException
import java.util.ArrayList


/**
 * To change the template for this generated type comment go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */

/** **Unicode String: BIFF8 Compressed String format**<br></br>
 *
 * <pre>
 * offer  name        size    contents
 * ---
 * 0       cch         2       Count of characters in string (NOT the number of bytes)
 * 2       grbit       1       Option flags
 * 3       rgb         var     Array of string characters and formatting runs.
 *
 * GRBIT Option codes:
 * bits    mask        name        contents
 * ---
 * 0       0x1         fHighByte   =0 if all chars have high byte of 0x0
 * (low bytes only are stored.)
 * =1 if at least one char in string has nonzero
 * highbyte.  All chars are double byte (uncompressed)
 * 1       0x2         (Reserved)  Must be zero.
 * 2       0x4         fExtSt      Far East Extended Strings.  Asian Phonetic Settings (phonetic)
 * 3       0x8         fRichSt     Rich String
 * 7-4     0xF0        (Reserved)  Must be zero.
 *
 * rgb:
 * [2 or 3] 	2 		(optional, only if richtext=1) Number of Rich-Text formatting runs (rt)
 * [var.] 		4 		(optional, only if phonetic=1) Size of Asian phonetic settings block (in bytes, sz)
 * var. 	ln or 2∙ln 	Character array (8-bit characters or 16-bit characters, dependent on ccompr)
 * [var.] 		4∙rt 	(optional, only if richtext=1) List of rt formatting runs
 * [var.] 		sz 		(optional, only if phonetic=1) Asian Phonetic Settings Block
 *
 *
 * Asian phonetic text can be used to provide extended phonetic information for specific characters or words. It
 * appears above the regular text (or to the right of vertical text), and can refer to single characters, groups of characters, or
 * entire words.
 * Offset Size Contents
 * 0 		2 	identifier 0001H
 * 2 		2 	Size of the following data (10 + 2∙ln + 6∙np)
 * 4 		2 	Index to FONT record used for the Asian phonetic text
 * 6 		2 	Additional settings for the Asian phonetic text:
 * Bit Mask Contents
 * 1-0 0003H Type of Japanese phonetic text (type):
 * 00 = Katakana (narrow) 10 = Hiragana
 * 01 = Katakana (wide)
 * 3-2 000CH Alignment of all portions of the Asian phonetic text (align):
 * 00 = Not specified (Japanese only) 10 = Centered
 * 01 = Left (Top for vertical text)  11 = Distributed
 * 5-4 0030H 11 (always set)
 * 8 		2 	Number of portions the Asian phonetic text is broken into (np).
 * If np = 0, the Asian phonetic text refers to the entire cell text.
 * 10 		2 	Total length of the following Asian phonetic text (number of characters, ln)
 * 12 		2 	Repeated total length of the text
 * 14 2∙ln or 2 Character array of Asian phonetic text, no Unicode string header, always 16-bit characters.
 * Note: If ln = 0, this field is not empty but contains 0000H.
 * 14+2∙ln 6∙np List of np structures that describe the position of each portion in the main text. Each
 * structure contains the following fields:
 * Offset Size Contents
 * 0 2 First character in the Asian phonetic text of this portion (cpa)
 * 2 2 First character of the main text belonging to this portion (cpm)
 * 4 2 Number of characters in main text belonging to this portion (ccm)
 *
 *
</pre> *
 * @see Sst
 *
 * @see Labelsst
 */
class Unicodestring : XLSConstants, Serializable {
    private val DEBUGLEVEL = -1
    internal var cch: Int = 0
    private var cchExtRst = 0
    private var ExtRst: ByteArray? = null
    // String stringVal;
    private var fHighByte = false
    private var formatlen = 0
    private var formatrunnum = ByteArray(2)
    private var formattingarray: ByteArray? = null
    internal var isRichString = false
        private set
    /** return the grbit for this UString
     */
    internal var grbit: Byte = 0
        private set

    // private Isstinf mybucket = null;
    private var numformattingruns = 0
    /** provide a hint to the CompatibleVector
     * about this objects likely position.
     */
    // public int getRecordIndexHint(){return this.recordIdx;}

    /** set the position of this string in the SST
     */
    /** set the Extsst Isstinf bucket for the USTRING see page 313 XL'97 book
     */
    //void setBucket(Isstinf i){this.mybucket = i;}

    /** set the position of this string in the SST
     */
    internal var sstPos = -1
    private var stringarray: ByteArray? = null
    private var stringlen: Int = 0
    private var offer: Int = 0
    /**
     * returns true if this unicode string is in Western Character Set
     * @return
     */
    var isWesternString = true
        private set


    /* return the eastern bytes for this Ustring
     */
    internal//cchExtRst = ExtRst.length;
    // put the number of formatting runs back
    val easternBytes: ByteArray
        get() {
            val data = ByteArray(stringarray!!.size + formatlen + offer + ExtRst!!.size)
            val cchbyte = ByteTools.cLongToLEBytes(cch)
            System.arraycopy(cchbyte, 0, data, 0, 2)
            data[2] = grbit
            val cchExtByte = ByteTools.cLongToLEBytes(cchExtRst)

            var tempoffset = 3
            if (isRichString) tempoffset = 5
            System.arraycopy(cchExtByte, 0, data, tempoffset, 4)

            stringlen = stringarray!!.size
            System.arraycopy(stringarray!!, 0, data, offer, stringlen)
            if (isRichString) {
                formatrunnum = ByteTools.shortToLEBytes(numformattingruns.toShort())
                data[3] = formatrunnum[0]
                data[4] = formatrunnum[1]
                System.arraycopy(formattingarray!!, 0, data, stringlen + offer, formatlen)
            }
            System.arraycopy(ExtRst!!, 0, data, stringlen + offer + formatlen, ExtRst!!.size)
            if (data[0].toInt() == 0x0) Logger.logInfo("Unicodestring has zero length.")
            return data
        }

    /** returns the length of the UNICODESTRING's bytes
     */
    internal val length: Int
        get() = stringarray!!.size + formatlen + offer

    // ride'em cowboy!
    internal// put the number of formatting runs back
    val westernBytes: ByteArray
        get() {
            val data = ByteArray(stringarray!!.size + formatlen + offer)
            val cchbyte = ByteTools.cLongToLEBytes(cch)
            System.arraycopy(cchbyte, 0, data, 0, 2)
            data[2] = grbit
            stringlen = stringarray!!.size
            System.arraycopy(stringarray!!, 0, data, offer, stringlen)
            if (isRichString) {
                formatrunnum = ByteTools.shortToLEBytes(numformattingruns.toShort())
                data[3] = formatrunnum[0]
                data[4] = formatrunnum[1]
                System.arraycopy(formattingarray!!, 0, data, stringlen + offer, formatlen)
            }
            return data
        }

    var len = 0

    /**
     * returns true if this unicode string is of Eastern Character Set
     * @return
     */
    val isEasternString: Boolean
        get() = !isWesternString

    val stringVal: String?
        get() = toString()

    /**
     * return formatting runs, if any, for this unicode strings
     * formatting runs are a list of one or more pairs of [char index, font index]
     * where char index determines where to apply font specified by font index
     * @return
     */
    val formattingRuns: ArrayList<*>?
        get() {
            if (numformattingruns == 0) return null
            val formattingruns = ArrayList()
            for (i in 0 until numformattingruns) {
                val charIndex = ByteTools.readShort(formattingarray!![i * 4].toInt(), formattingarray!![i * 4 + 1].toInt())
                val fontIndex = ByteTools.readShort(formattingarray!![i * 4 + 2].toInt(), formattingarray!![i * 4 + 3].toInt())
                formattingruns.add(shortArrayOf(charIndex, fontIndex))
            }
            return formattingruns
        }

    /** override equals based on value
     * of String.
     */
    override fun equals(obj: Any?): Boolean {
        if (obj is Unicodestring) {
            // compare Unicode strings at byte level rather than just string value
            // if(val.equals(obj.toString()))return true;
            return if (this.isWesternString) {
                java.util.Arrays.equals(this.westernBytes, obj.westernBytes)
            } else {
                java.util.Arrays.equals(this.easternBytes, obj.easternBytes)
            }
        } else if (obj is String) {
            val `val` = this.toString()
            return `val` == obj
        }
        return false
    }


    /**Your standard init method.  Choose the format of the string and init it.
     *
     * @param ustrdata - the data to init the String from
     * @param extrstbrk - whether the string ExtRst data spans a continue boundary
     */
    internal fun init(ustrdata: ByteArray, extrstbrk: Boolean) {
        len = ustrdata.size
        grbit = ustrdata[2]
        try {
            if (grbit and 0x4 == 0x4) {
                isWesternString = false
                this.initEasternEncoding(ustrdata, extrstbrk)
            } else {
                this.initWesternEncoding(ustrdata)
            }
        } catch (e: NegativeArraySizeException) {
            // error processing string but don't fail entirely
        }

    }

    /**
     * Init method for japanese/chinese/? style UnicodeStrings
     */
    @Throws(NegativeArraySizeException::class)
    internal fun initEasternEncoding(ustrdata: ByteArray, extrstbrk: Boolean) {
        // get the row, col and ixfe information
        cch = ByteTools.readShort(ustrdata[0].toInt(), ustrdata[1].toInt()).toInt()
        var dlen = cch
        grbit = ustrdata[2]
        offer = 7

        // double byte? almost always here
        if (grbit and 0x1 == 0x1) {
            dlen *= 2
            fHighByte = true
        }

        stringarray = ByteArray(dlen)
        // is rich text?  has formatting runs?
        if (grbit and 0x8 == 0x8) {
            isRichString = true
        }

        if (isRichString) {
            offer = 9
            try {
                System.arraycopy(ustrdata, offer, stringarray!!, 0, dlen) // string data
                formatrunnum[0] = ustrdata[3]
                formatrunnum[1] = ustrdata[4]
                numformattingruns = ByteTools.readShort(formatrunnum[0].toInt(), formatrunnum[1].toInt()).toInt()
                formatlen = numformattingruns * 4
                formattingarray = ByteArray(formatlen)
                System.arraycopy(ustrdata, dlen + offer, formattingarray!!, 0, formatlen)  // formatting info
                cchExtRst = ByteTools.readInt(ustrdata[5], ustrdata[6], ustrdata[7], ustrdata[8])
                // report error???
                if (ustrdata.size - dlen - offer - cchExtRst != formatlen)
                    io.starter.toolkit.Logger.log("Unicodestring.initEasternEncoding: Format runs are not correct")
                ExtRst = ByteArray(cchExtRst)
                System.arraycopy(ustrdata, dlen + offer + formatlen, ExtRst!!, 0, cchExtRst) // Extendadata
                //report error???
                if (ExtRst!![0].toInt() != 1 || ExtRst!![1].toInt() != 0)
                    Logger.logWarn("Unicodestring.initEasternEncoding: Phonetic Data is not correct")
            } catch (e: Throwable) {
                Logger.logInfo("Problem parsing rich text Eastern Unicodestring.  len:$dlen rich:$isRichString.  $e")
            }

        } else {
            try {
                cchExtRst = ByteTools.readInt(ustrdata[3], ustrdata[4], ustrdata[5], ustrdata[6])
                val extrstLen = ustrdata.size - (formatlen + offer + dlen)
                val off = 0

                // the most important code in the world
                if (cchExtRst != extrstLen || extrstbrk) {
                    if (DEBUGLEVEL > 0) Logger.logWarn("Unicodestring ExtRst Inconsistent.")
                    //off = 1;
                    cchExtRst = extrstLen
                }
                ExtRst = ByteArray(cchExtRst)

                System.arraycopy(ustrdata, 7, stringarray!!, 0, dlen) // string data
                System.arraycopy(ustrdata, ustrdata.size - ExtRst!!.size - off, ExtRst!!, 0, ExtRst!!.size) // Extendadata
                // report error???
                if (ExtRst!![0].toInt() != 1 || ExtRst!![1].toInt() != 0)
                    Logger.logWarn("Unicodestring.initEasternEncoding: Phonetic Data is not correct")
            } catch (t: Throwable) {
                Logger.logInfo("Problem Parsing non-rich Eastern Unicodestring.  len:$dlen rich:$isRichString.  $t")
            }

        }
    }

    /**
     * Init method for default unicode string types.
     */
    @Throws(NegativeArraySizeException::class)
    internal fun initWesternEncoding(ustrdata: ByteArray) {
        // get the row, col and ixfe information
        cch = ByteTools.readShort(ustrdata[0].toInt(), ustrdata[1].toInt()).toInt()
        var dlen = cch + 0
        grbit = ustrdata[2]
        offer = 3

        // is is double byte stream?
        if (grbit and 0x1 == 0x1) {
            dlen *= 2
            fHighByte = true
        } else {
            fHighByte = false
        }

        // handle data being greater than the length of the bytes
        if (dlen + offer > ustrdata.size) {
            dlen = ustrdata.size - offer
        }

        // is rich text?  does it have formatting runs?
        if (grbit and 0x8 == 0x8) {
            isRichString = true
        }

        if (isRichString) {
            formatrunnum[0] = ustrdata[3]
            formatrunnum[1] = ustrdata[4]
            numformattingruns = ByteTools.readShort(formatrunnum[0].toInt(), formatrunnum[1].toInt()).toInt()
            offer = 5
            formatlen = ustrdata.size - dlen - offer
            formattingarray = ByteArray(formatlen)
            System.arraycopy(ustrdata, dlen + offer, formattingarray!!, 0, formatlen)
        }
        stringarray = ByteArray(dlen + 0)
        try {
            System.arraycopy(ustrdata, offer, stringarray!!, 0, dlen)
        } catch (e: Exception) {
            Logger.logInfo("Problem Parsing Western Unicodestring.  len:$dlen rich:$isRichString.  $e")
        }

    }


    /** returns whether the Stream position passed in
     * falls within the UString's character data as opposed
     * to its length/formatting runs and other non character data
     *
     * Read as "typical plot by microsoft to make our lives miserable"
     *
     * @param pos - the position of the String in the Sst data
     * @return whether the string can be broken into parts to span a continue
     *
     * FROM OPENOFFICE.ORG DOCS...
     * Unicode strings are split in a special way. At the beginning of each CONTINUE record the option flags byte is repeated.
     * Only the character size flag will be set in this flags byte, the Rich-Text flag and the Far-East flag are set to zero.
     *
     * In each CONTINUE record it is possible that the character size changes from 8-bit characters to 16-bit characters
     * and vice versa. !
     * Never a Unicode string is split until and including the first character. That means, all header fields (string length,
     * option flags, optional Rich-Text size, and optional Far-East data size) and the first character of the string have to
     * occur together in the leading record, or have to be moved completely into the CONTINUE record. !
     * Formatting runs (?2.1) cannot be split between their components (character index and FONT record index). If a
     * string is split between two formatting runs, the option flags field will not be repeated in the CONTINUE record.
     */
    internal fun isBreakable(pos: Int): Boolean {
        val mypos = this.sstPos
        val strstart = mypos + offer
        val strend = strstart + stringarray!!.size
        val bts = this.read()
        var brkpos = strstart - pos
        if (brkpos < 0) brkpos *= -1
        if (this.cch < 2) return false // don't break one-char strings -- the 0x1 is confused as a grbit

        if (cchExtRst > 0) return false
        if (pos > mypos + cchExtRst + stringarray!!.size + offer) return false // in ExtRst data...

        return if (pos > strstart && pos < strend) {
            this.length > 8220
        } else false
    }

    /**
     * Returns wheter the specified index is breaking in the middle of a character.  Only useful or needed
     * for double byte strings, where the continue boundary cannot be in the middle of a char.
     * @return true if illegal break
     */
    fun charBreakOnBounds(pos: Int): Boolean {
        val mypos = this.sstPos
        val strstart = mypos + offer
        val strend = strstart + stringarray!!.size
        val bts = this.read()
        var brkpos = strstart - pos
        if (brkpos < 0) brkpos *= -1
        return brkpos % 2 != 0

    }

    /** returns the UNICODESTRING's bytes
     */

    fun read(): ByteArray {
        return if (isWesternString) this.westernBytes else this.easternBytes
    }

    /**
     * returns just the string array portion of this Unicode String
     * minus header, formatting, etc.
     * @return
     */
    fun readStr(): ByteArray? {
        return stringarray
    }

    /** return the String representation of this Unicodestring
     */
    override fun toString(): String? {
        try {
            return if (fHighByte) {
                String(stringarray!!, XLSConstants.UNICODEENCODING) // defaultEncoding);
            } else String(stringarray!!, XLSConstants.DEFAULTENCODING)
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("Problem decoding Unicodestring.  " + e + " Resorting to default encoding: " + XLSConstants.DEFAULTENCODING)
            try {
                return String(stringarray!!, XLSConstants.DEFAULTENCODING) // supported by JDK1.1 +
            } catch (t: UnsupportedEncodingException) {
                if (DEBUGLEVEL > -1) Logger.logInfo("Problem decoding Unicodestring.  $t")
            }

        }

        return null
    }

    /** return the String representation of this Unicodestring.  Sets a caching string value in the Unicode String,
     * primarily needed for lookups when shared strings = true
     */
    fun toCachingString(): String? {
        try {
            return if (fHighByte) {
                /* KSC: TESTING:  for  handling japanese fonts...            	boolean isjp; // KSC: testing
            	if (!westernencoding || Unicodestring.containsJapanese(this.getStringVal()))
            		isjp= true;
*/
                String(stringarray!!, XLSConstants.UNICODEENCODING) // defaultEncoding);
            } else {
                String(stringarray!!, XLSConstants.DEFAULTENCODING)
            }
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("Problem decoding Unicodestring.  " + e + " Resorting to default encoding: " + XLSConstants.DEFAULTENCODING)
            try {
                return String(stringarray!!, XLSConstants.DEFAULTENCODING) // supported by JDK1.1 +
            } catch (t: UnsupportedEncodingException) {
                if (DEBUGLEVEL > -1) Logger.logInfo("Problem decoding Unicodestring.  $t")
                return null
            }

        }

    }

    /** updates the unicode string
     */
    fun updateUnicodeString(s: String) {
        if (s == this.toString()) return
        // 0       cch         2       Count of characters in string (NOT the number of bytes)
        // 2       grbit       1       Option flags
        // make sure to get formatting runs if present
        var strbytes: ByteArray? = null
        try {
            if (!ByteTools.isUnicode(s)) {
                strbytes = s.toByteArray(charset(XLSConstants.DEFAULTENCODING))
                grbit = (grbit and 0xFE).toByte()
            } else {
                strbytes = s.toByteArray(charset(XLSConstants.UNICODEENCODING))
                if (grbit and 0x1 != 0x1) {
                    grbit += 0x1
                }
            }
        } catch (e: UnsupportedEncodingException) {
            Logger.logInfo("Problem encoding string: $e")
        }

        val strdatalen = strbytes!!.size
        val strlen = s.length
        val blen = ByteTools.cLongToLEBytes(strlen)
        if (grbit and 0x4 == 0x4) { // it was an eastern string.  Blow that extrst out...
            grbit = (grbit and 0xFB).toByte()
            ExtRst = null
            offer -= 4
            isWesternString = true
        }
        if (grbit and 0x8 == 0x8) { // had formatting runs - remove (since cannot keep with new string)
            formatlen = 0
            isRichString = false
            grbit = grbit xor 0x8
            offer -= 2
        }
        val newbytes = ByteArray(strdatalen + offer + formatlen)
        newbytes[2] = grbit
        System.arraycopy(blen, 0, newbytes, 0, 2)

        System.arraycopy(strbytes, 0, newbytes, offer, strdatalen)
        if (this.isRichString) {
            System.arraycopy(this.formattingarray!!, 0, newbytes, strdatalen + this.offer, this.formatlen)
            // put the number of formatting runs back
            newbytes[3] = this.formatrunnum[0]
            newbytes[4] = this.formatrunnum[1]
        }
        this.init(newbytes, false)
    }

    /**
     * Return true if the string contains formatting runs embedded within it
     *
     * @return
     */
    fun hasFormattingRuns(): Boolean {
        return numformattingruns > 0
    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -1800227752945355535L

        /** returns true if the char c is a double-byte character  */
        private fun isJapanese(c: Char): Boolean {
            return c >= '\u0100' && c <= '\uffff'
            // simpler:  return c>'\u00ff';
        }

        /** returns true if the String s contains any Japanese characters  */
        fun containsJapanese(s: String): Boolean {
            for (i in 0 until s.length) {
                if (isJapaneseII(s[i])) {
                    return true
                }
            }
            return false
        }

        /** returns true if the char c is a Japanese character.  */
        private fun isJapaneseII(c: Char): Boolean {
            // katakana:
            if (c >= '\u30a0' && c <= '\u30ff') return true
            // hiragana
            if (c >= '\u3040' && c <= '\u309f') return true
            // CJK Unified Ideographs
            if (c >= '\u4e00' && c <= '\u9fff') return true
            // CJK symbols & punctuation
            if (c >= '\u3000' && c <= '\u303f') return true
            // KangXi (kanji)
            if (c >= '\u2f00' && c <= '\u2fdf') return true
            // KanBun
            if (c >= '\u3190' && c <= '\u319f') return true
            // CJK Unified Ideographs Extension A
            if (c >= '\u3400' && c <= '\u4db5') return true
            // CJK Compatibility Forms
            if (c >= '\ufe30' && c <= '\ufe4f') return true
            // CJK Compatibility
            return if (c >= '\u3300' && c <= '\u33ff') true else c >= '\u2e80' && c <= '\u2eff'
            // CJK Radicals Supplement
            // other character..
        }
    }

    /** info on encoding:
     *
     *
     * 16 code from 0 x0000 to 0 x007F) is the ASCII characters
     * the next 128 Unicode characters (x0080 code from 0 to 0 x00FF) is ISO? 8859-1 on the expansion of ASCII.
     * Unicode in different parts of the same characters are based on existing standards.
     * This is to facilitate the conversion.
     * Greek alphabet x0370 use from 0 to 0 x03FF code,
     * The use of Slavic language x0400 from 0 to 0x04FF code
     * United States use x0530 from 0 to 0x058F code
     * Hebrew x0590 from 0 to 0 x05FF code.
     * China, Japan and South Korea hieroglyphs (known as the CJK) occupiers from 0x3000 to 0x9FFF code.
     * There are four types of Shift-JIS encoding available (although you may need to customize the encoding list to see them all).
     * Japanes(Shift JIS) contains the half-width alphanumeric characters, symbols, and katakana, along with JIS X0208 (JIS 1&2) characters.
     * Japanese (Mac OS) contains all of Japanese (Shift JIS) along with some Mac-specific symbols.
     * Japanese (Windows, DOS) contains all of Japanese (Shift JIS) along with some Windows-specific symbols.
     * Japanese (Shift JIS X0213) contains all of Japanese (Shift JIS) along with kanji specified in JIS 3&4 for a total of 13,000 characters total.
     * Japanese text created in Classic OS is encoded in Japanese (Mac OS), so this is the safest option among the various Shift JIS encodings.
     *
     * Asian characters are Kanji, Hiragana, Katakana
     * (full and half-width), full-width numbers and punctuation,
     * and Chinese and Korean characters. This is often abbreviated as “CJK.”
     *
     *
     * Japanese Language Encoding
     * Encoding		Other Names		Vendor/Standard Body		Other Rosette Names
     * CCSID 1027		EBCDIK			Microsoft & IBM				CCSID-1027, CCSID1027
     * CCSID 290		EBCDIK			Microsoft & IBM				CCSID-290, CCSID290
     * CCSID 930						IBM							CCSID-930, CCSID930
     * CCSID 939						IBM							CCSID-939, CCSID939
     * CCSID 942						Microsoft & IBM				CCSID-942, CCSID942
     * CP10001			Macintosh Japanese	Microsoft & IBM			CP10001
     * CP20290			(full/half width Latin & halfwidth katakana) Microsoft & IBM	CP20290
     * CP21027			(halfwidth Latin, halfwidth katakana&private use)	Microsoft & IBM	CP21027
     * EUC-JP							Unix						EUC-JP, EUC-J
     * EUC-JP-JISROMAN					Unix						EUC-JP-JISROMAN
     * ISO 2022-JP						International or National Standard	ISO-2022-JP
     * JapaneseAutoDetect	For encodings, see JapaneseAutodetect	Rosette Autodetect	JapaneseAutoDetect
     *
     * Encoding		Other Names		Vendor/Standard Body		Other Rosette Names
     * JIS_X_0201		HalfWidthKatakana	International or National Standard	JIS_X_0201, IBM897
     * JIS_X_0208						International or National Standard	JIS_X_0208
     * Shift-JISMS		MS_Kanji, CP932	Microsoft & IBM				Shift-JIS, SJIS
     * Shift_JIS-2004	ShiftJISX0213	Microsoft & IBM				Shift-Jis2004, Shift_JISX0213,Shift-X
     * Shift-JIS78		Shift-JIS without MS/IBM extensions	Unix/Macintosh	Shift-JIS78, SJIS78
     *
     * Chinese Language Encoding
     * Encoding		Other Names		Vendor/Standard Body		Other Rosette Names
     * ChineseAutoDetect	For encodings, see ChineseAutodetect	Rosette Autodetect	ChineseAutoDetect
     * HKSCS							International or National Standard	HKSCS
     * ISO 2022-CN						International or National Standard	ISO-2022-CN
     * GB 18030						International or National Standard  GB18030
     * Chinese, Simplified
     * CCSID 935						IBM							CCSID-935, CCSID935
     * EUC-CN			GB2312, EUC-SC	Unix						GB2312
     * GB2312			EUC-CN, EUC-SC	International or National Standard	GB2312
     * HZ-GB-2312		HZ-GB-2312		International or National Standard	HZ, HZ-GB-2312
     * CP936			GBK				Microsoft & IBM				CP936, GBK
     * MacChineseSimplified			Macintosh					MacChineseSimplified
     * Chinese, Traditional
     * CCSID 937						IBM							CCSID-937, CCSID937
     * CNS-11643-1986	EUC-TW			International or National Standard	CNS-11643-1986
     * CNS-11643-1992	EUC-TW			International or National Standard	CNS-11643, CNS-11643-1992
     * EUC-TW			CNS-11643-1986, CNS-11643-1992 Unix			CNS-11643, CNS-11643-1992
     * GB12345							International or National Standard	GB12345
     * Big5							International or National Standard	Big5
     * Big5+							International or National Standard	Big5+, Big5Plus
     * CP10002			Macintosh Traditional Chinese	Microsoft & IBM		CP10002
     * CP950							Microsoft & IBM				CP950
     * MacChineseTraditional			Macintosh					MacChineseTraditional
     *
     * The CCSID in the iSeries objects must be 935 (Simplified Chinese)
     * The code page in the PC for Simplified Chinese is 1388
     * The Data Type in the DSPF/PRTF must be O (Other)
     *
     * more info:
     * Shift_JIS	DBCS		16-bit Japanese encoding (Note that you must use an underscore character (_), not a hyphen (-) in the name in CFML attributes.)
     * (same as MS932)
     * EUC-KR		DBCS		16-bit Korean encoding
     * UCS-2		DBCS		Two-byte Unicode encoding
     * UTF-8		MBCS		Multibyte Unicode encoding. ASCII is 7-bit; non-ASCII characters used in European and many Middle Eastern languages are two-byte; and most Asian characters are three-byte
     *
     *
     *
     * ExtRst data is defined as:
     *
     * ID (WORD): 0x0001
     * Length (WORD)
     * unknown (WORD)
     * Flag (WORD)
     * Number of relation informations between Katakana and Kanji (WORD)
     * Number of characters #1 (WORD)
     * Number of characters #2 (WORD)
     * >>> always the same ?
     * Katakana characters
     * Relation informations between Katakana and Kanji (6 bytes each)
     *
     *
     *
     * mdImeMode (8 bits): An unsigned integer that specifies the Input Method Editor (IME) mode
     * enforced by this data validation. This value is only used when the input language is one of the
     * following languages:
     *
     * 1.       Chinese Simplified (Locale ID = 2052)
     * 2.       Chinese Traditional (Locale ID = 1028)
     * 3.       Japanese (Locale ID = 1041)
     * 4.       Korean (Locale ID = 1042)
     * The input for the cell can be restricted to specific sets of characters, as specified by the value of
     * mdImeMode. MUST be a value from the following table:
     * Value          Meaning
     * 0x00           No Control
     * 0x01           On
     * 0x02           Off (English)
     * 0x04           Hiragana
     * 0x05           wide katakana
     * 0x06           narrow katakana
     * 0x07           Full-width alphanumeric
     * 0x08           Half-width alphanumeric
     * 0x09           Full-width hangul
     * 0x0A           Half-width hangul
     *
     *
     * ExtRst:
     * Asian phonetic text5 (Ruby) can be used to provide extended phonetic information for specific characters or words. It
     * appears above the regular text (or to the right of vertical text), and can refer to single characters, groups of characters, or
     * entire words.
     * Offset Size Contents
     * 0 2 Unknown identifier 0001H
     * 2 2 Size of the following data (10 + 2∙ln + 6∙np)
     * 4 2 Index to FONT record (➜5.45) used for the Asian phonetic text
     * 6 2 Additional settings for the Asian phonetic text:
     * Bit Mask Contents
     * 1-0 0003H Type of Japanese phonetic text (type):
     * 002 = Katakana (narrow) 102 = Hiragana
     * 012 = Katakana (wide)
     * 3-2 000CH Alignment of all portions of the Asian phonetic text (align):
     * 002 = Not specified (Japanese only) 102 = Centered
     * 012 = Left (Top for vertical text) 112 = Distributed
     * 5-4 0030H 112 (always set)
     * 8 2 Number of portions the Asian phonetic text is broken into (np).
     * If np = 0, the Asian phonetic text refers to the entire cell text.
     * 10 2 Total length of the following Asian phonetic text (number of characters, ln)
     * 12 2 Repeated total length of the text
     * 14 2∙ln or 2 Character array of Asian phonetic text, no Unicode string header, always 16-bit characters.
     * Note: If ln = 0, this field is not empty but contains 0000H.
     * 14+2∙ln 6∙np List of np structures that describe the position of each portion in the main text. Each
     * structure contains the following fields:
     * Offset Size Contents
     * 0 2 First character in the Asian phonetic text of this portion (cpa)
     * 2 2 First character of the main text belonging to this portion (cpm)
     * 4 2 Number of characters in main text belonging to this portion (ccm)
     *
     * Example: Japanese word Tokyo (東京) with added hiragana (とうきょう)6. The following examples show the
     * contents of the important fields of the Asian Phonetic Settings Block.
     * Example 1: Hiragana centered over the entire word:
     * とうきょう
     * 東京
     * type = 102 (hiragana)
     * align = 102 (centered)
     * np = 0 (no portions, hiragana refers to entire text)
     * ln = 5 (length of entire hiragana text)
     * No portion structures
     *
     *
     * Asian phonetic text5 (Ruby) can be used to provide extended phonetic information for specific characters or words. It
     * appears above the regular text (or to the right of vertical text), and can refer to single characters, groups of characters, or
     * entire words.
     * Offset Size Contents
     * 0 2 Unknown identifier 0001H
     * 2 2 Size of the following data (10 + 2∙ln + 6∙np)
     * 4 2 Index to FONT record (➜5.45) used for the Asian phonetic text
     * 6 2 Additional settings for the Asian phonetic text:
     * Bit Mask Contents
     * 1-0 0003H Type of Japanese phonetic text (type):
     * 002 = Katakana (narrow) 102 = Hiragana
     * 012 = Katakana (wide)
     * 3-2 000CH Alignment of all portions of the Asian phonetic text (align):
     * 002 = Not specified (Japanese only) 102 = Centered
     * 012 = Left (Top for vertical text) 112 = Distributed
     * 5-4 0030H 112 (always set)
     * 8 2 Number of portions the Asian phonetic text is broken into (np).
     * If np = 0, the Asian phonetic text refers to the entire cell text.
     * 10 2 Total length of the following Asian phonetic text (number of characters, ln)
     * 12 2 Repeated total length of the text
     * 14 2∙ln or 2 Character array of Asian phonetic text, no Unicode string header, always 16-bit characters.
     * Note: If ln = 0, this field is not empty but contains 0000H.
     * 14+2∙ln 6∙np List of np structures that describe the position of each portion in the main text. Each
     * structure contains the following fields:
     * Offset Size Contents
     * 0 2 First character in the Asian phonetic text of this portion (cpa)
     * 2 2 First character of the main text belonging to this portion (cpm)
     * 4 2 Number of characters in main text belonging to this portion (ccm)
     * 5      */
}
/**
 * string 1=
 * [-115, -1, 111, -1, 107, -1, 102, -1, 114, 0, 108, 0, 100, 0, 33, 0]
 * phonetic
 * [1, 0, 12, 0, 1, 0, 55, 0, 0, 0, 0, 0, 0, 0, 0, 0]
 * new String(stringarray, "UTF-16LE")
 * (java.lang.String) ﾍｯｫｦrld!
 * String 3=
 * [-40, 48, -61, 48, -87, 48, -14, 48, 82, -1, 76, -1, 68, -1, 1, -1]
 * phonetic
 * [1, 0, 12, 0, 1, 0, 55, 0, 0, 0, 0, 0, 0, 0, 0, 0]
 * new String(stringarray, "UTF-16LE")
 * (java.lang.String) ヘッォヲｒｌｄ！
 *
 *
 * Excel 2007 Supported Character Sets:
 * 0x00 Specifies the ANSI character set.
 * 0x01 Specifies the default character set.
 * 0x02 Specifies the Symbol character set.
 * 0x4D Specifies a Macintosh (Standard Roman) character set.
 * 0x80 Specifies the JIS character set.
 * 0x81 Specifies the Hangul character set.
 * 0x82 Specifies a Johab character set.
 * 0x86 Specifies the GB-2312 character set.
 * 0x88 Specifies the Chinese Big Five character set.
 * 0xA1 Specifies a Greek character set.
 * 0xA2 Specifies a Turkish character set.
 * 0xA3 Specifies a Vietnamese character set.
 * 0xB1 Specifies a Hebrew character set.
 * 0xB2 Specifies an Arabic character set.
 * 0xBA Specifies a Baltic character set.
 * 0xCC Specifies a Russian character set.
 * 0xDE Specifies a Thai character set.
 * 0xEE Specifies an Eastern European character set.
 * 0xFF Specifies an OEM character set not defined by this Office Open XML Standard.
 * Any other value Application-defined, may be ignored.
 */
