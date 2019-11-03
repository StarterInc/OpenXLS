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
import io.starter.OpenXLS.WorkBookHandle
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.Serializable
import java.io.UnsupportedEncodingException


/**
 * **Hlink: Hyperlink (1b8h)**<br></br>
 *
 *
 * hyperlink record
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       rwFirst         2       First row of link
 * 6       rwLast          2       Last row of link
 * 8       colFirst        2       First col of link
 * 10      colLast         2       Last col of link
 * 12      rgbHlink        var     HLINK data
 *
 *
</pre> *
 */
class Hlink : XLSRecord() {
    internal var colFirst = -1
    internal var colLast = -1
    internal var rowFirst = -1
    internal var rowLast = -1
    private var linkStruct: HLinkStruct? = null

    /**
     * get the URL for this Hlink
     */
    val url: String
        get() = if (linkStruct == null) "" else linkStruct!!.getUrl()

    /**
     * return the description part of the hyperlink
     *
     * @return
     */
    val description: String
        get() = if (linkStruct == null) "" else linkStruct!!.linkText

    /**
     * @return
     */
    /**
     * @param range
     */
    var range: CellRange? = null

    /**
     * set last/first cols/rows
     */
    fun setRowFirst(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        System.arraycopy(b, 0, dt!!, 0, 2)
        this.rowFirst = c
    }

    fun getRowFirst(): Int {
        return rowFirst
    }

    /**
     * set last/first cols/rows
     */
    fun setRowLast(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        System.arraycopy(b, 0, dt!!, 2, 2)
        this.rowLast = c
    }

    fun getRowLast(): Int {
        return rowLast
    }

    /**
     * set last/first cols/rows
     */
    fun setColFirst(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        System.arraycopy(b, 0, dt!!, 4, 2)
        this.colFirst = c
    }

    fun getColFirst(): Int {
        return colFirst
    }

    /**
     * set last/first cols/rows
     */
    fun setColLast(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        System.arraycopy(b, 0, dt!!, 6, 2)
        this.colLast = c
    }

    fun getColLast(): Int {
        return colLast
    }

    /**
     * set link URL with description and test mark
     * note that either url or text mark must be present ...
     *
     * @param url
     * @param textMark
     * @param desc
     */
    fun setURL(url: String, desc: String, textMark: String) {
        try {
            if (url == "" && textMark == "") {
                Logger.logWarn("HLINK.setURL:  no url or text mark specified")
                return
            }
            linkStruct!!.setUrl(url, desc, textMark)
        } catch (e: Exception) {
            Logger.logWarn("setting URL $url failed: $e")
        }

        val bt = linkStruct!!.bytes
        this.setData(bt)
    }

    fun setFileURL(url: String, desc: String, textMark: String) {
        try {
            linkStruct!!.setFileURL(url, desc, textMark)
        } catch (e: Exception) {
            Logger.logWarn("setting URL $url failed: $e")
        }

        val bt = linkStruct!!.bytes
        this.setData(bt)
    }

    /**
     * returns whether a given col
     * is referenced by this Hyperlink
     */
    fun inrange(x: Int): Boolean {
        return x <= colLast && x >= colFirst
    }

    override fun init() {
        super.init()
        var pos = 0
        rowFirst = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()).toInt()
        rowLast = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()).toInt()
        colFirst = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()).toInt()
        colLast = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()).toInt()
        val ustr = Unicodestring()

        var nm = ""
        if (this.sheet != null) {
            if (this.sheet!!.sheetName != "") nm = this.sheet!!.sheetName + "!"
        }
        nm += io.starter.OpenXLS.ExcelTools.getAlphaVal(colFirst)
        nm += "$rowFirst:"
        nm += io.starter.OpenXLS.ExcelTools.getAlphaVal(colLast)
        nm += rowLast
        try {
            range = CellRange(nm, null)
        } catch (e: Exception) {
            Logger.logWarn("initializing Hlink record failed: $e")
        }

        if (DEBUGLEVEL > 5) Logger.logInfo("Hlink Cells: " + range!!.toString())

        try {
            linkStruct = HLinkStruct(this.getBytesAt(0, this.length))
        } catch (e: Exception) {
            Logger.logWarn("Hyperlink parse failed for Cells " + range!!.toString() + ": " + e)
        }

        if (DEBUGLEVEL > 5) Logger.logInfo("Hlink URL: " + linkStruct!!.getUrl())
    }


    fun initCells(wbook: WorkBookHandle) {
        val cellcoords = IntArray(4)
        cellcoords[0] = this.getRowFirst()
        cellcoords[2] = this.getRowLast()
        cellcoords[1] = this.getColFirst()
        cellcoords[3] = this.getColLast()

        try {
            val cr = CellRange(wbook.getWorkSheet(this.sheet!!.sheetName), cellcoords)
            cr.workBook = wbook
            val ch = cr.cellRecs
            for (t in ch.indices) {
                ch[t].hyperlink = this
            }
        } catch (e: Exception) {
            Logger.logWarn("initializing Hyperlink Cells failed: $e")
        }

    }

    companion object {

        private val serialVersionUID = -4259979643231173799L

        //        retlab.setURL(url, desc, textMark);		200606 KSC: do separately!
        val prototype: XLSRecord?
            get() {
                val retlab = Hlink()
                val rbytes = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, -48, -55, -22, 121, -7, -70, -50, 17, -116, -126, 0, -86, 0, 75, -87, 11, 2, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, -32, -55, -22, 121, -7, -70, -50, 17, -116, -126, 0, -86, 0, 75, -87, 11, -116, 0, 0, 0, 0, 0)

                retlab.opcode = XLSConstants.HLINK
                retlab.length = 94.toShort()
                retlab.setData(rbytes)
                retlab.init()
                return retlab
            }
    }
}

/*
    HLINK Struct
    ??      16
    int?    4
    int?    4
    cch?    4
    var     var
    int?    4
    ?       16
    cch     4
    urlstr  var
*/
internal class HLinkStruct
/**
 * Inner class with no documentation
 */
(barr: ByteArray) : XLSConstants, Serializable {
    var isLink = false
    var isAbsoluteLink = false
    var hasDescription = false
    var hasTextMark = false
    var hasTargetFrame = false
    var isUNCPath = false
    var grbit = ByteArray(4)

    var DEBUG = false

    var url = ""
    /**
     * get the URL link text for this Hlink
     */
    var linkText = ""
    var textMark = ""
    var targetFrame = ""
    var int1 = -1
    var urlPos = -1
    var int4 = -1
    var bytes: ByteArray? = null
        private set

    /**
     * decodes option flag into vars
     * Here is the breakdown on what these mean:
     *
     *
     * standard (non-local) URL:	isLink= true, isAbsoluteLink= true, isUNCPath= false
     * local file:					isLink= true, isUNCPath= false
     * UNC path:					isLink= true, isAbsoluteLink= true, isUNCPath= true
     * link in current workbook:	isLink= false, isAbsoluteLink= false, hasTextMark= true, isUNCPath= false
     *
     * @param grbytes
     */
    fun decodeGrbit(grbytes: ByteArray) {
        if ((grbytes[0] and 0x1).toByte() > 0x0) isLink = true
        if ((grbytes[0] and 0x2).toByte() > 0x0) isAbsoluteLink = true
        // 20060406 KSC: need both bits 2 & 4 to be set for hasDescription
        //		if ((byte)(grbytes[0] & 0x14) > 0x0)hasDescription = true;
        if ((grbytes[0] and 0x14).toByte() >= 0x14) hasDescription = true
        if ((grbytes[0] and 0x8).toByte() > 0x0) hasTextMark = true
        if ((grbytes[0] and 0x80).toByte() > 0x0) hasTargetFrame = true
        if ((grbytes[0] and 0x100).toByte() > 0x0) isUNCPath = true
    }

    fun setGrbit() {
        grbit[0] = 0x0
        if (isLink) grbit[0] = (0x1 or grbit[0]).toByte()
        if (isAbsoluteLink) grbit[0] = (0x2 or grbit[0]).toByte()
        if (hasDescription) grbit[0] = (0x14 or grbit[0]).toByte()
        if (hasTextMark) grbit[0] = (0x8 or grbit[0]).toByte()
        if (hasTargetFrame) grbit[0] = (0x80 or grbit[0]).toByte()
        if (isUNCPath) grbit[0] = (0x100 or grbit[0]).toByte()
        bytes[28] = grbit[0]
    }

    /**
     * get the URL for this Hlink
     */
    fun getUrl(): String {
        return if (textMark == "")
            url
        else
            "$url#$textMark"
    }


    // 20060406 KSC:  mods to setUrl:
    // 	Added ability to set description + modified byte input to work mo' betta ...

    /**
     * set the URL for this Hlink, description= URL <default>
     *
     *
     * Assume link URL i.e. no file URL, UNC path, etc.
    </default> */
    fun setUrl(url: String) {
        setUrl(url, url, "")
    }

    /**
     * set standard link URL for this Hlink and optional description
     * i.e. no file URL, UNC path, text marks ...
     */
    fun setUrl(ur: String, desc: String) {
        setUrl(url, desc, "")
    }

    /**
     * set proper settings for URL and write bytes
     *
     * @param ur
     * @param desc
     * @param textMark
     */
    fun setUrl(ur: String, desc: String, textMark: String) {
        isLink = true
        isAbsoluteLink = true
        isUNCPath = false
        hasDescription = desc.length > 0
        hasTextMark = textMark.length > 0
        setBytes(ur, desc, textMark)
    }

    /**
     * Assume NOT TRUE FILE URL (avoids writing complex dir info + using FILE_GUID)
     * so difference between link URL and file URL is the isAbsoluteLink var ...
     *
     * @param ur
     * @param desc
     * @param textMark
     */
    fun setFileURL(ur: String, desc: String, textMark: String) {
        isLink = true
        isAbsoluteLink = false
        isUNCPath = false
        hasDescription = desc.length > 0
        hasTextMark = textMark.length > 0
        setBytes(ur, desc, textMark)
    }

    fun setBytes(ur: String, desc: String, tm: String) {
        try {
            setGrbit()
            val pos = 32    // start of description/text input

            val blankbytes = ByteArray(2)    // trailing zero word
            var newbytes = ByteArray(pos)
            System.arraycopy(bytes!!, 0, newbytes, 0, pos) // copy old pre-string data into new array

            if (hasDescription) {
                // copy char array length (cch) of description + description bytes
                val descbytes = desc.toByteArray(charset(XLSConstants.UNICODEENCODING))
                val newcch = descbytes.size
                val newcchbytes = ByteTools.cLongToLEBytes(newcch / 2 + 1)

                // copy cch of desc in...
                newbytes = ByteTools.append(newcchbytes, newbytes)

                //copy bytes of description in
                newbytes = ByteTools.append(descbytes, newbytes)
                // copy trailing dumb str bytes in
                newbytes = ByteTools.append(blankbytes, newbytes)
            }

            /* TODO:  Implement target frame right here
			if (hasTargetFrame) {
				// copy targetFrame bytes + length
				// get cch
				byte[] tfbytes= tf.getBytes(UNICODEENCODING);
	            int newcch = tfbytes.length;
				byte[] newcchbytes = ByteTools.cLongToLEBytes(newcch/2 +1);
				// copy cch of tm in...
				newbytes = ByteTools.append(newcchbytes, newbytes);

				//copy bytes of tf in
				newbytes = ByteTools.append(tfbytes, newbytes);

				// copy trailing dumb str bytes in
				newbytes = ByteTools.append(blankbytes, newbytes);
			}
			*/

            /* URL Handling */
            // copy GUID in - ASSUME URL_GUID since aren't supporting relative FILE_GUIDs!
            newbytes = ByteTools.append(URL_GUID, newbytes)

            if (isLink) {
                // copy url bytes + length
                // get cch (which is different alg. from both description + textmark)
                val urlbytes = ur.toByteArray(charset(XLSConstants.UNICODEENCODING))
                val newcch = urlbytes.size

                // copy cch of url in...
                val newcchbytes = ByteTools.cLongToLEBytes(newcch + 2)
                newbytes = ByteTools.append(newcchbytes, newbytes)

                // copy url bytes in
                newbytes = ByteTools.append(urlbytes, newbytes)

                // copy trailing dumb str bytes in
                newbytes = ByteTools.append(blankbytes, newbytes)
            }

            if (hasTextMark) {
                // copy textmark bytes + length
                // get cch
                val tmbytes = tm.toByteArray(charset(XLSConstants.UNICODEENCODING))
                val newcch = tmbytes.size
                val newcchbytes = ByteTools.cLongToLEBytes(newcch / 2 + 1)
                // copy cch of tm in...
                newbytes = ByteTools.append(newcchbytes, newbytes)

                //copy bytes of tm in
                newbytes = ByteTools.append(tmbytes, newbytes)

                // copy trailing dumb str bytes in
                newbytes = ByteTools.append(blankbytes, newbytes)
            }
            this.bytes = newbytes

            this.linkText = desc
            this.textMark = tm
            this.url = ur
        } catch (e: UnsupportedEncodingException) {
            Logger.logWarn("Setting URL failed: $ur: $e")
        }

    }

    init {
        bytes = barr
        var pos = 28

        System.arraycopy(barr, 28, grbit, 0, 4)
        decodeGrbit(grbit)
        pos += 4

        /*
         * This section gets the display string for the Hyperlink, if it exists.
         */
        if (hasDescription) {
            val cch = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])
            if (cch > 0) { // 20070814 KSC: shouldn't be 0 and also hasDescription ...
                try {
                    val descripbytes = ByteArray(cch * 2 - 2)
                    System.arraycopy(barr, pos, descripbytes, 0, cch * 2 - 2)
                    linkText = String(descripbytes, XLSConstants.UNICODEENCODING)
                    pos += cch * 2
                    if (DEBUG) Logger.logInfo("Hlink.hlstruct Display URL: $linkText")
                } catch (e: Exception) {
                    if (DEBUG) Logger.logWarn("decoding Display URL in Hlink: $e")
                }

            }
        }
        /*
         * if it has a target frame, read in
         */
        if (hasTargetFrame) {
            val cch = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])
            if (cch > 0) {
                try {
                    val tfbytes = ByteArray(cch * 2 - 2)
                    System.arraycopy(barr, pos, tfbytes, 0, cch * 2 - 2)
                    targetFrame = String(tfbytes, XLSConstants.UNICODEENCODING)
                    if (DEBUG)
                        Logger.logInfo("Hlink.hlstruct targetFrame: $targetFrame")
                    pos += cch * 2
                } catch (e: Exception) {
                    if (DEBUG)
                        Logger.logWarn("Hlink Decode of targetFrame failed: $e")
                }

            }
        }
        /*
         * URL section:  non-local URL or Link in current file
         */
        if (isLink) {
            val GUID = ByteArray(16)
            System.arraycopy(barr, pos, GUID, 0, 16)
            val bIsCurrentFileRef = java.util.Arrays.equals(GUID, FILE_GUID)
            pos += 16    // skip GUID
            if (!bIsCurrentFileRef) {    // then it's a URL or non-relative file path
                urlPos = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])
                if (urlPos > 0) {
                    try {
                        val urlbytes = ByteArray(urlPos - 2)
                        System.arraycopy(barr, pos, urlbytes, 0, urlPos - 2)
                        url = String(urlbytes, XLSConstants.UNICODEENCODING)
                        if (DEBUG)
                            Logger.logInfo("Hlink.hlstruct URL: $url")
                        pos += urlPos
                    } catch (e: Exception) {
                        if (DEBUG)
                            Logger.logWarn("Hlink Decode of URL failed: $e")
                    }

                }
            } else {    // (appears to be a) current file link (Actuality is different than documentation!)
                val dirUps = ByteTools.readShort(barr[pos++].toInt(), barr[pos++].toInt()).toInt()
                urlPos = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])
                if (urlPos > 0) {
                    try {
                        val urlbytes = ByteArray(urlPos - 1)
                        System.arraycopy(barr, pos, urlbytes, 0, urlPos - 1)
                        url = String(urlbytes, XLSConstants.DEFAULTENCODING)
                        if (DEBUG)
                            Logger.logInfo("Hlink.hlstruct File URL: $url")
                        pos += urlPos + 24    // add char count + avoid the 24 "unknown" bytes
                        val extraInfo = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])
                        if (extraInfo > 0) {
                            pos += extraInfo
                            val sz = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])


                        }
                    } catch (e: Exception) {
                        if (DEBUG)
                            Logger.logWarn("Hlink Decode of File URL failed: $e")
                    }

                }
            }
        }

        if (hasTextMark) {
            val cch = ByteTools.readInt(barr[pos++], barr[pos++], barr[pos++], barr[pos++])
            if (cch > 0) {
                try {
                    val tmbytes = ByteArray(cch * 2 - 2)
                    System.arraycopy(barr, pos, tmbytes, 0, cch * 2 - 2)
                    textMark = String(tmbytes, XLSConstants.UNICODEENCODING)
                    if (DEBUG)
                        Logger.logInfo("Hlink.hlstruct textMark: $textMark")
                    pos += cch * 2
                } catch (e: Exception) {
                    if (DEBUG)
                        Logger.logWarn("Hlink Decode of textmark failed: $e")
                }

            }
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -1915454683496117350L

        /**
         * fills the HLINK bytes based on the settings of
         * isLink
         * isAbsoluteLink
         * hasDescription
         * hasTextMark
         * hasTargetFrame
         * isUNCPath
         *
         *
         * how these are set, before calling this method, determine
         * how the HLINK record bytes are written
         *
         * @param ur    URL string
         * @param desc    optional descrption
         * @param tm    optional text mark text as in:  ...#textmarktext
         */
        val URL_GUID = byteArrayOf(-32, -55, -22, 121, -7, -70, -50, 17, -116, -126, 0, -86, 0, 75, -87, 11)
        val FILE_GUID = byteArrayOf(3, 3, 0, 0, 0, 0, 0, 0, -64, 0, 0, 0, 0, 0, 0, 70)
    }

}

