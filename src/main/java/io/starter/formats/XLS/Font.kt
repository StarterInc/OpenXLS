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

import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.Color
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParser

import java.io.UnsupportedEncodingException


/**
 * **Font: Font Description (231h)**<br></br>
 *
 *
 * Font records describe a font in the workbook
 *
 *
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       dyHeight        2       Height in 1/20 point (twips)
 * 6       grbit           2       attributes
 * grbit Mask Contents
 * 0 0001H 1 = Characters are bold
 * 1 0002H 1 = Characters are italic
 * 2 0004H 1 = Characters are underlined
 * 3 0008H 1 = Characters are struck out
 * 8       icv             2       index to color palette
 * 10      bls             2       bold style (weight)  100-1000 default is 190h norm 2bch bold
 * 12      sss             2       super/sub (0 = none, 1 = super, 2 = sub)
 * 14      uls             1       Underline Style (0 = none, 1 = single, 2 = double, 21h = single acctg, 22h = dble acctg)
 * 15      bFamily         1       Font Family (WinAPI LOGFONT struct)
 * 16      bCharSet        1       Characterset (WinAPI LOGFONT struct)
 * 17      reserved        0
 * 18      cch             1       Length of font name
 * 19      rgch            var     Font name
 *
</pre> *
 *
 *
 *
 * "http://www.extentech.com">Extentech Inc.
 *
 * @see XF
 *
 * @see FORMAT
 */

class Font : io.starter.formats.XLS.XLSRecord, FormatConstants {
    private var grbit: Short = -1
    private var cch: Short = -1
    private var dyHeight: Short = -1
    private var icv: Short = -1
    private var bls: Short = -1
    private var sss: Short = -1
    private var uls: Short = -1
    private var bFamily: Short = -1
    private var bCharSet: Short = 0
    private var fontName = ""
    // OOXML specifics:
    private var customColor: Color? = null // holds custom color (OOXML or other use)
    /**
     * return whether this font is Condensed (OOXML-specific) Macintosh
     * compatibility setting.
     *
     * @return
     */
    /**
     * set whether this font is Condensed (OOXML-specific) Macintosh
     * compatibility setting.
     *
     * @param condensed
     */
    var isCondensed: Boolean = false
    /**
     * return whether this font is Expanded (OOXML-specific) Macintosh
     * compatibility setting.
     *
     * @return expanded
     */
    /**
     * set whether this font is Extended (OOXML-specific) Macintosh
     * compatibility setting.
     *
     * @param expanded
     */
    var isExtended: Boolean = false

    /**
     * Get if the font record is striken our or not
     *
     * @return
     */
    var stricken: Boolean
        get() = grbit and BITMASK_STRIKEOUT == BITMASK_STRIKEOUT
        set(b) {
            if (b)
                grbit = (grbit or BITMASK_STRIKEOUT).toShort()
            else {
                val grbittemp = grbit xor BITMASK_STRIKEOUT
                grbit = (grbittemp and grbit).toShort()
            }
            this.setGrbit()
        }

    /**
     * Get if the font is italic or not
     *
     * @return
     */
    // false: need to account for multiple font formats - add step 2: &
    // original grbit
    var italic: Boolean
        get() {
            val isItalic = grbit and BITMASK_ITALIC
            return isItalic == BITMASK_ITALIC
        }
        set(b) {
            if (b)
                grbit = (grbit or BITMASK_ITALIC).toShort()
            else {
                val grbittemp = grbit xor BITMASK_ITALIC
                grbit = (grbittemp and grbit).toShort()
            }
            this.setGrbit()
        }

    /**
     * Get if the font is underlined or not
     *
     * @return
     */
    // 20070821 KSC: should also set underline
    // style ...
    var underlined: Boolean
        get() = grbit and BITMASK_UNDERLINED == BITMASK_UNDERLINED
        set(b) {
            if (b)
                grbit = (grbit or BITMASK_UNDERLINED).toShort()
            else {
                val grbittemp = grbit xor BITMASK_UNDERLINED
                grbit = (grbittemp and grbit).toShort()
            }
            this.setGrbit()
            setUnderlineStyle(1.toByte())
        }

    /**
     * get if the font is bold or not
     *
     * @return
     */
    /**
     * Set or unset bold attribute of the font record
     *
     * @param b
     */
    var bold: Boolean
        get() = grbit and BITMASK_BOLD == BITMASK_BOLD
        set(b) {
            if (data == null)
                this.setData(this.getData())
            if (b) {
                val boldbytes = ByteTools.shortToLEBytes(0x2bc.toShort())
                System.arraycopy(boldbytes, 0, data!!, 6, 2)
                bls = 0x2bc
                grbit = (grbit or BITMASK_BOLD).toShort()
            } else {
                val boldbytes = ByteTools.shortToLEBytes(0x190.toShort())
                System.arraycopy(boldbytes, 0, data!!, 6, 2)
                bls = 0x190
                val grbittemp = grbit xor BITMASK_BOLD
                grbit = (grbittemp and grbit).toShort()
            }
            this.setGrbit()
        }

    /**
     * @param idx
     */
    var idx = -1

    /**
     * add to Fonts table in Workbook
     */
    override var workBook: WorkBook?
        get
        set(b) {
            super.workBook = b
            if (this.idx == -1) {
                this.idx = this.workBook!!.addFont(this)
            }
        }

    /**
     * Get an int representing the underline style of this record, matches int
     * records in FormatConstants.STYLE_UNDERLINE*****
     *
     * @return
     * @see
     */
    val underlineStyle: Int
        get() = this.getData()!![10].toInt()

    /**
     * returns the super/sub script for the Font
     *
     * @return int (0 = none, 1 = super, 2 = sub)
     */
    /**
     *
     */
    /**
     * Set the super/sub script for the Font
     *
     *
     * super/sub (0 = none, 1 = super, 2 = sub)
     */
    var script: Int
        get() = sss.toInt()
        set(ss) {
            if (data == null)
                this.setData(this.getData())
            val newss = ByteTools.shortToLEBytes(ss.toShort())
            System.arraycopy(newss, 0, data!!, 8, 2)
            sss = ss.toShort()
        }

    /**
     * Set the weight of the font in 1/20 point units 100-1000 range.
     */
    var fontWeight: Int
        get() = this.bls.toInt()
        set(wt) {
            if (data == null)
                this.setData(this.getData())
            val newwt = ByteTools.shortToLEBytes(wt.toShort())
            System.arraycopy(newwt, 0, data!!, 6, 2)
            bls = wt.toShort()
        }

    /**
     * Set the size of the font in 1/20 point units
     */
    var fontHeight: Int
        get() = this.dyHeight.toInt()
        set(ht) {
            if (data == null)
                this.setData(this.getData())
            val newht = ByteTools.shortToLEBytes(ht.toShort())
            System.arraycopy(newht, 0, data!!, 0, 2)
            dyHeight = ht.toShort()
        }

    val fontHeightInPoints: Double
        get() = this.dyHeight / 20.0

    /**
     * Get the color for this Font as a avt.Color
     *
     * @return
     */
    // If icv is System window text color=7FFF, default fg color or default tooltip text color, return black
    /* notes: special icv values:
		0x0040	Default foreground color. This is the window text color in the sheet display.
		0x0041	Default background color. This is the window background color in the sheet display and is the default background color for a cell.
		0x004D	Default chart foreground color. This is the window text color in the chart display.
		0x004E	Default chart background color. This is the window background color in the chart display.
		0x004F	Chart neutral color which is black, an RGB value of (0,0,0).
		0x0051	ToolTip text color. This is the automatic font color for comments.
		0x7FFF	Font automatic color. This is the window text color.
		*/ val colorAsColor: java.awt.Color
        get() {
            if (customColor != null)
                return customColor!!.colorAsColor
            if (this.icv.toInt() == 0x7FFF || this.icv.toInt() == 0x40 || this.icv.toInt() == 0x51) {
                return java.awt.Color.BLACK
            } else if (this.icv > FormatHandle.COLORTABLE.size) {
                return java.awt.Color.BLACK
            }
            return if (this.workBook == null) FormatHandle.COLORTABLE[this.icv] else this.colorTable[this.icv]
        }

    /**
     * Get the color of this Font as a web-compliant Hex String
     */
    // remove "FF" from beginning
    val colorAsHex: String
        get() = if (customColor != null && customColor!!.colorAsOOXMLRBG != null) "#" + customColor!!.colorAsOOXMLRBG!!.substring(2) else FormatHandle.colorToHexString(colorAsColor)

    /**
     * returns the font color as an OOXML-compliant Hex Stringf	 *
     *
     * @return
     */
    val colorAsOOXMLRBG: String
        get() {
            val rgbcolor = colorAsHex
            return "FF" + rgbcolor.substring(1)
        }

    /**
     * gets the color of this font as an index into excel 2003 color table
     *
     * @return int
     */
    val fontColor: Int
        @Deprecated("use getColor()")
        get() = color

    /**
     * gets the color of this font as an index into excel 2003 color table
     *
     * @return int
     */
    /**
     * Set the font color via index into 2003 color table
     */
    // this is a value for system font color, default
    // to black
    // don't do it if the font is already this color
    var color: Int
        get() {
            if (customColor != null)
                return customColor!!.colorInt
            return if (this.icv.toInt() == 32767) 0 else this.icv.toInt()
        }
        set(cl) {
            if (data == null)
                setData(this.getData())
            if (cl != icv.toInt()) {
                val newcl = ByteTools.shortToLEBytes(cl.toShort())
                System.arraycopy(newcl, 0, data!!, 4, 2)
                icv = cl.toShort()
            }
            if (customColor != null)
                customColor!!.colorInt = cl
        }

    /**
     * Get if the font is bold or not
     */
    val isBold: Boolean
        get() = bls > 0x190

    /**
     * @return an XML descriptor for this Font
     */
    // changed from 'getFontInfoXML'
    val xml: String
        get() = getXML(false)

    /**
     * return the OOXML font color element
     *
     * @return
     */
    /**
     * store OOXML font color
     *
     * @param c
     */
    var ooxmlColor: Color?
        get() = customColor
        set(c) {
            if (c != null)
                this.color = c.colorInt
            customColor = c
        }

    /**
     * generate the OOXML to define this Font
     *
     * @return
     */
    // TODO: family, scheme
    // the default
    // KSC: modify due to certain XLS->XLSX issues with automatic color
    // leave automatic "blank"
    // for incremental styles, font size may not be set
    // for incremental styles, font name may
    // not be set
    // TODO: family val= # (see OOXMLConstants)
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<font>")
            if (this.isBold)
                ooxml.append("<b/>")
            if (this.italic)
                ooxml.append("<i/>")
            if (this.underlined) {
                val u = this.underlineStyle
                if (u == 1)
                    ooxml.append("<u/>")
                else if (u == 2)
                    ooxml.append("<u val=\"double\"/>")
                else if (u == 0x21)
                    ooxml.append("<u val=\"singleAccounting\"/>")
                else if (u == 0x22)
                    ooxml.append("<u val=\"doubleAccounting\"/>")
            }
            if (this.stricken)
                ooxml.append("<strike/>")
            if (!this.isCondensed)
                ooxml.append("<condense val=\"0\"/>")
            if (!this.isExtended)
                ooxml.append("<extend val=\"0\"/>")
            val c = this.ooxmlColor
            if (c != null) {
                ooxml.append(c.ooxml)
            } else {
                if (this.icv.toInt() != 9 && this.icv.toInt() != 64) {
                    val cl = this.color
                    if (cl > 0)
                        ooxml.append("<color rgb=\"FF"
                                + FormatHandle.colorToHexString(
                                this.colorTable[cl]).substring(1)
                                + "\"/>")
                }
            }
            val sz = this.fontHeightInPoints
            if (sz > 0)
                ooxml.append("<sz val=\"$sz\"/>")
            val n = this.getFontName()
            if (n != null && n != "")
                ooxml.append("<name val=\"$n\"/>")

            ooxml.append("</font>")
            ooxml.append("\r\n")
            return ooxml.toString()
        }

    /**
     * return the appropriate SVG string to define this font
     *
     * @return
     */
    // sbf.append(" fill='#222222'"); // TODO: get proper text color
    val svg: String
        get() {
            val sbf = StringBuffer("font-family='"
                    + this.getFontName() + "'")
            sbf.append(" font-size='" + this.fontHeightInPoints + "pt'")
            sbf.append(" font-weight='" + this.fontWeight + "'")
            if (this.icv.toInt() != 9)
                sbf.append(" fill='"
                        + FormatHandle.colorToHexString(FormatHandle.getColor(this
                        .color)) + "'")
            else
                sbf.append(" fill='"
                        + FormatHandle.colorToHexString(FormatHandle.getColor(0))
                        + "'")
            return sbf.toString()
        }

    /**
     * EXPERIMENTAL AND MAY NOT BE COMPLETE <br></br>
     * Returns true if this font is a Unicode (non-ascii) Charset
     *
     * @return
     */
    private val isUnicodeCharSet: Boolean
        get() = (bCharSet.toInt() == SHIFTJIS_CHARSET || bCharSet.toInt() == HANGEUL_CHARSET
                || bCharSet.toInt() == HANGUL_CHARSET || bCharSet.toInt() == GB2312_CHARSET
                || bCharSet.toInt() == CHINESEBIG5_CHARSET
                || bCharSet.toInt() == HEBREW_CHARSET || bCharSet.toInt() == ARABIC_CHARSET
                || bCharSet.toInt() == GREEK_CHARSET || bCharSet.toInt() == TURKISH_CHARSET
                || bCharSet.toInt() == VIETNAMESE_CHARSET || bCharSet.toInt() == THAI_CHARSET
                || bCharSet.toInt() == EASTEUROPE_CHARSET
                || bCharSet.toInt() == RUSSIAN_CHARSET || bCharSet.toInt() == BALTIC_CHARSET)

    /**
     * Initialize the font record
     */
    override fun init() {
        super.init()
        dyHeight = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())// Height
        // in
        // 1/20
        // point
        grbit = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())// attributes
        icv = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())// index
        // to
        // color
        // palette
        bls = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())// bold
        // style
        // (weight)
        // 100-1000
        // default
        // is
        // 190h
        // norm
        // 2bch
        // bold
        sss = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())// super/sub
        // (0 =
        // none,
        // 1 =
        // super,
        // 2 =
        // sub)
        uls = this.getByteAt(10).toShort()// Underline Style (0 = none, 1 = single, 2 =
        // double, 21h = single acctg, 22h = dble
        // acctg)
        bFamily = this.getByteAt(11).toShort()// Font Family (WinAPI LOGFONT struct)
        /**
         * lfCharSet
         *
         * The character set. The following values are predefined.
         *
         * ANSI_CHARSET BALTIC_CHARSET xx CHINESEBIG5_CHARSET xx DEFAULT_CHARSET
         * xx EASTEUROPE_CHARSET GB2312_CHARSET xx GREEK_CHARSET xx
         * HANGUL_CHARSET xx MAC_CHARSET xx OEM_CHARSET xx RUSSIAN_CHARSET xx
         * SHIFTJIS_CHARSET xx SYMBOL_CHARSET xx TURKISH_CHARSET xx
         * VIETNAMESE_CHARSET
         *
         *
         * Korean language edition of Windows: JOHAB_CHARSET
         *
         * Middle East language edition of Windows: ARABIC_CHARSET
         * HEBREW_CHARSET
         *
         * Thai language edition of Windows: THAI_CHARSET
         *
         * The OEM_CHARSET value specifies a character set that is
         * operating-system dependent.
         *
         * DEFAULT_CHARSET is set to a value based on the current system locale.
         * For example, when the system locale is English (United States), it is
         * set as ANSI_CHARSET.
         *
         * Fonts with other character sets may exist in the operating system. If
         * an application uses a font with an unknown character set, it should
         * not attempt to translate or interpret strings that are rendered with
         * that font.
         *
         * This parameter is important in the font mapping process. To ensure
         * consistent results, specify a specific character set. If you specify
         * a typeface name in the lfFaceName member, make sure that the
         * lfCharSet value matches the character set of the typeface specified
         * in lfFaceName.
         */
        bCharSet = ByteTools.readUnsignedShort(this.getByteAt(12),
                0.toByte()).toShort()// Characterset (WinAPI LOGFONT struct)

        // this.getData()[13]= 0;// set byte to 0 for comparisons
        // get the Name
        var pos = 14
        cch = this.getByteAt(pos++).toShort()
        var buflen = cch * 2
        pos++
        var compressed = false

        if (buflen + pos >= this.length) {
            buflen = this.length - pos
            compressed = true
        }

        if (buflen < 0) {
            Logger.logWarn("could not parse font: length reported as $buflen")
            return
        }

        val namebytes = this.getBytesAt(pos, buflen)
        if (!compressed) {
            pos = 0
            try {
                fontName = String(namebytes!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logErr("Font name decoding failed.", e)
            }

        } else { // compressed
            fontName = String(namebytes!!)
        }
    }

    /**
     * update the Grbit bytes in the underlying byte stream
     */
    fun setGrbit() {
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(grbit)
        System.arraycopy(b, 0, data!!, 2, 2)
        this.setData(data)

    }

    override fun toString(): String {
        return (this.fontName + "," + this.bls + "," + this.dyHeight + " "
                + this.colorAsColor + " font style:[" + this.bold
                + this.italic + this.stricken + this.underlined
                + this.color + this.underlineStyle + "]")
    }

    fun getFontName(): String? {
        return fontName
    }

    /**
     * Set the underline style of this font recotd
     *
     * @param styl
     */
    fun setUnderlineStyle(styl: Byte) {
        this.uls = styl.toShort()
        this.getData()[10] = styl
    }

    constructor() {}

    /**
     * Create a New Font from the String definition.
     *
     *
     * Roughly matches the functionality of the java.awt.Font class.
     *
     * @param String font name
     * @param int    font style
     * @param int    font size in Points
     */
    constructor(nm: String, stl: Int, sz: Int) {
        val bl = byteArrayOf(-56, 0, 0, 0, -1, 127, -112, 1, 0, 0, 0, 0, 0, 0, 5, 1, 65, 0, 114, 0, 105, 0, 97, 0, 108, 0)
        opcode = XLSConstants.FONT
        length = bl.size.toShort()
        this.setData(bl)
        this.init()
        this.setFontName(nm)
        this.fontWeight = stl
        this.fontHeight = sz
    }

    /**
     * Set the Font name.
     *
     *
     * To be valid, this font name must be available on the client system.
     */
    fun setFontName(fn: String) {
        var namebytes: ByteArray? = null
        try {
            namebytes = fn.toByteArray(charset(XLSConstants.UNICODEENCODING))
        } catch (e: UnsupportedEncodingException) {
            Logger.logWarn("setting Font Name using Default Encoding failed: $e")
            namebytes = fn.toByteArray()
        }

        cch = (namebytes!!.size / 2).toShort()
        fontName = fn
        val newdata = ByteArray(namebytes.size + 16)
        System.arraycopy(this.getBytesAt(0, 13)!!, 0, newdata, 0, 13)// 20061027 KSC: keep 13th byte for sake of comparisons - 20070816 - revert to original
        System.arraycopy(this.getBytesAt(0, 14)!!, 0, newdata, 0, 14)
        newdata[14] = cch.toByte()
        newdata[15] = 1.toByte()
        System.arraycopy(namebytes, 0, newdata, 16, namebytes.size)
        this.setData(newdata)
        this.init()
    }

    /**
     * Sets the font color via java.awt.Color
     */
    fun setColor(color: java.awt.Color) {
        if (customColor != null)
            customColor!!.setColor(color)
        else
            customColor = Color(color, "color", this.workBook!!.theme)
        icv = customColor!!.colorInt.toShort()
        val newcl = ByteTools.shortToLEBytes(icv)
        System.arraycopy(newcl, 0, data!!, 4, 2)
    }

    /**
     * Sets the font color via a web-compliant Hex String
     */
    fun setColor(clr: String) {
        if (customColor != null)
            customColor!!.setColor(clr)
        else
            customColor = Color(clr, "color", this.workBook!!.theme)
        icv = customColor!!.colorInt.toShort()
        val newcl = ByteTools.shortToLEBytes(icv)
        System.arraycopy(newcl, 0, data!!, 4, 2)
    }

    /**
     * return an XML desciptor for this font
     *
     * @param convertToUnicodeFont if true, font family will be changed to ArialUnicodeMS
     * (standard unicode) for non-ascii fonts
     * @return
     */
    fun getXML(convertToUnicodeFont: Boolean): String {
        val sb = StringBuffer()
        if (!convertToUnicodeFont || !isUnicodeCharSet)
            sb.append("name=\""
                    + StringTool.convertXMLChars(this.getFontName()) + "\"")
        else
            sb.append("name=\"ArialUnicodeMS\"")
        sb.append(" size=\"" + this.fontHeightInPoints + "\"")
        sb.append(" color=\""
                + FormatHandle.colorToHexString(this.colorAsColor) + "\"")
        sb.append(" weight=\"" + this.fontWeight + "\"")
        if (this.isBold) {
            sb.append(" bold=\"1\"")
        }
        if (this.underlineStyle != Font.STYLE_UNDERLINE_NONE.toInt())
            sb.append(" underline=\"$underlineStyle\"")
        return sb.toString()
    }

    /**
     * return true if font f matches key attributes of this font
     *
     * @param f
     * @return
     */
    fun matches(f: Font): Boolean {
        return (this.fontName == f.fontName && this.dyHeight == f.dyHeight
                && this.bls == f.bls && this.color == f.color
                && this.sss == f.sss && this.uls == f.uls && this.grbit == f.grbit)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -398444997553403671L

        // grbit flags
        internal val BITMASK_BOLD = 0x0001
        internal val BITMASK_ITALIC = 0x0002
        internal val BITMASK_UNDERLINED = 0x0004
        internal val BITMASK_STRIKEOUT = 0x0008

        // charset values
        internal val ANSI_CHARSET = 0
        internal val DEFAULT_CHARSET = 1
        internal val SYMBOL_CHARSET = 2
        internal val SHIFTJIS_CHARSET = 128
        internal val HANGEUL_CHARSET = 129
        internal val HANGUL_CHARSET = 129
        internal val GB2312_CHARSET = 134
        internal val CHINESEBIG5_CHARSET = 136
        internal val OEM_CHARSET = 255
        internal val JOHAB_CHARSET = 130
        internal val HEBREW_CHARSET = 177
        internal val ARABIC_CHARSET = 178
        internal val GREEK_CHARSET = 161
        internal val TURKISH_CHARSET = 162
        internal val VIETNAMESE_CHARSET = 163
        internal val THAI_CHARSET = 222
        internal val EASTEUROPE_CHARSET = 238
        internal val RUSSIAN_CHARSET = 204
        internal val MAC_CHARSET = 77
        internal val BALTIC_CHARSET = 186

        /**
         * utility to convert points to correct font height
         *
         * @param h
         * @return
         */
        fun PointsToFontHeight(h: Double): Int {
            return (h * 20).toInt()
        }

        fun FontHeightToPoints(h: Int): Double {
            return h / 20.0
        }

        /**
         * parse incoming OOXML into a Font object
         *
         * @param xpp
         * @return
         */
        // TODO: family, scheme
        fun parseOOXML(xpp: XmlPullParser, bk: WorkBookHandle): Font {
            var c: Color? = null
            var sz: String? = null
            var name = ""
            var u: Any? = null
            var b = false
            var strike = false
            var ital = false
            var condense = false
            var expand = false
            try {
                var eventType = xpp.next()
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "sz") {
                            sz = xpp.getAttributeValue(0)
                        } else if (tnm == "name") {
                            name = xpp.getAttributeValue(0)
                        } else if (tnm == "b") {
                            if (xpp.attributeCount == 0)
                                b = true
                            else
                                b = xpp.getAttributeValue(0) == "1"
                        } else if (tnm == "i") {
                            if (xpp.attributeCount == 0)
                                ital = true
                            else
                                ital = xpp.getAttributeValue(0) == "1"
                        } else if (tnm == "u") {
                            if (xpp.attributeCount == 0)
                                u = java.lang.Boolean.valueOf(true)
                            else
                                u = xpp.getAttributeValue(0)
                        } else if (tnm == "strike") {
                            strike = true
                        } else if (tnm == "condense") {
                            condense = false
                        } else if (tnm == "expand") {
                            expand = false
                        } else if (tnm == "color") {
                            c = Color.parseOOXML(xpp, FormatHandle.colorFONT, bk) as Color
                        }
                    } else if (eventType == XmlPullParser.END_TAG && xpp.name == "font")
                        break
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("Font.parseOOXML: $e")
            }

            // for incremental styles, font size may not be set
            val size = if (sz == null)
                -1
            else
                Font.PointsToFontHeight(Double(sz))
            val f = Font(name, 400, size)
            if (c != null)
                f.ooxmlColor = c
            if (u != null) {
                f.underlined = true
                if (u is String) {
                    if (u == "double")
                        f.setUnderlineStyle(2.toByte())
                    else if (u == "singleAccounting")
                        f.setUnderlineStyle(0x21.toByte())
                    else if (u == "doubleAccounting")
                        f.setUnderlineStyle(0x22.toByte())
                }
            }
            if (b)
                f.bold = b
            if (ital)
                f.italic = true
            if (strike)
                f.stricken = true
            if (!condense)
                f.isCondensed = false
            if (!expand)
                f.isExtended = false
            return f
        }
    }
}