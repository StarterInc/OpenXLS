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
import io.starter.formats.OOXML.Fill
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.awt.Color


/**
 * **XF: Extended Format (E0h)**<br></br>
 * The XF record stores formatting properties.
 *
 *
 * If fStyle bit is true, then the XF is a style XF, otherwise
 * it is a BiffRec XF.  Cells and Styles both contain ixfe pointers
 * which correspond to their associated XF record.
 *
 * <pre>
 * BiffRec XF Record
 *
 * offset  Bits   MASK     name        contents
 * ---
 * 0       15-0   0xFFFF   ifnt        Index to the FONT record.
 * 2       15-0   0xFFFF   ifmt        Index to the FORMAT record.
 * 4       0      0x0001   fLocked     =1 if the cell is locked.
 * 1      0x0002   fHidden     =1 if the cell formula is hidden (value still shown)
 * 2      0x0004   fStyle      =0 for cell XF.
 * =1 for style XF.
 *
 * ~~~ additional option flags omitted ~~~
 *
</pre> *
 *
 * @see SST
 *
 * @see LABELSST
 *
 * @see EXTSST
 */


class Xf : io.starter.formats.XLS.XLSRecord {

    var idx = -1
        internal set
    // Shared variables.
    /**
     * @return Returns the ifnt.
     */
    /**
     * @param ifnt The ifnt to set.
     */
    var ifnt: Short = 0
    var ifmt: Short = 0
        private set
    private var fLocked: Short = 0
    private var fHidden: Short = 0
    private var fStyle: Short = 0
    // the records that are initialized to "0" are a 1/0 flag.  Just makin my life easier...
    private var f123Prefix: Short = 0
    private var ixfParent: Short = 0
    private var alc: Short = 0
    private var fWrap: Short = 0
    private var alcV: Short = 0
    // private short fJustLast; Used in Far East version of Excel only!!
    private var cIndent: Short = 0
    private var trot: Short = 0
    // private short cIntednt;
    private var fShrinkToFit: Short = 0
    private var fMergeCell: Short = 0
    private var iReadingOrder: Short = 0
    private var fAtrNum: Short = 0
    private var fAtrFnt: Short = 0
    private var fAtrAlc: Short = 0
    private var fAtrBdr: Short = 0
    private var fAtrPat: Short = 0
    private var fAtrProt: Short = 0
    private var dgLeft: Short = 0
    private var dgRight: Short = 0
    private var dgTop: Short = 0
    private var dgBottom: Short = 0
    private var icvLeft: Short = 0
    private var icvRight: Short = 0
    private var grbitDiag: Short = 0
    private var icvTop: Short = 0
    private var icvBottom: Short = 0
    private var icvDiag: Short = 0
    var diagBorderLineStyle: Short = 0
        private set
    private var fls: Short = 0
    private var icvFore: Short = 0
    private var icvBack: Short = 0
    private var fSxButton: Short = 0
    private var icvColorFlag: Short = 0
    internal var Iflag = 0
    internal var mystery: Byte = 0
    private var pat: String? = null
    /**
     * OOXML fill, if any
     */
    var fill: Fill? = null    // ugly that it's public ...
    // These should only be populated for boundsheet transferral issues.
    private var myFont: Font? = null
    private var myFormat: Format? = null
    /**
     * return # records using this xf
     *
     * @return
     */
    var useCount: Short = 0
        private set    // KSC: added 20121003 to keep track of xf usage by biffrecs

    /**
     * Set the workbook for this XF
     *
     *
     * This can get called multiple times.  This results in a disparity within
     * xf counting in workbook.
     */
    override var workBook: WorkBook?
        get
        set(b) {
            super.workBook = b
        }


    // marginal!
    // 0xf4 ?
    var merged: Boolean
        get() = fMergeCell.toInt() == 1
        set(mgd) {
            val rkdata = this.getData()
            rkdata[9] = 0x78.toByte()
            if (DEBUGLEVEL > 1) Logger.logInfo("Xf The merge style bit is: $fMergeCell")

        }

    /**
     * returns the associated  Font record for this XF
     */
    override val font: Font?
        get() {
            if (myFont != null) return myFont
            myFont = this.workBook!!.getFont(ifnt.toInt())
            return myFont
        }

    /**
     * returns whether this Format is a Date
     *
     *
     * Needs to be revisited.  Currently I am only returning true for the standard "built in" dates
     */
    // Check the format ID against all known date formats. Why do we do
    // this instead of letting it be caught by the string matching below?
    // toLowerCase is a simplistic way to implement the case insensitivity
    // of the pattern tokens. It could cause issues with string literals.
    val isDatePattern: Boolean
        get() {
            for (x in FormatConstants.DATE_FORMATS.indices) {
                val sxt = Integer.parseInt(FormatConstants.DATE_FORMATS[x][1], 16).toShort()
                if (ifmt == sxt)
                    return true
            }

            val fmt = this.workBook!!.getFormat(ifmt.toInt()) ?: return false
            val myfmt = fmt.format!!.toLowerCase()
            return isDatePattern(myfmt)
        }


    /**
     * returns whether this Format is a Currency
     */
    // what up with this?	General?
    // probably a built-in format that is not a currency format
    val isCurrencyPattern: Boolean
        get() {
            if (pat == null) {
                formatPattern = formatPattern
            }
            for (x in FormatConstants.CURRENCY_FORMATS.indices) {
                val cpt = Integer.parseInt(FormatConstants.CURRENCY_FORMATS[x][1], 16).toShort()
                if (ifmt == cpt) {
                    val ptx = FormatConstants.CURRENCY_FORMATS[x][0]
                    return if (cpt.toInt() == 1) {
                        pat == ptx

                    } else {
                        true
                    }
                }
            }
            val fmt = this.workBook!!.getFormat(ifmt.toInt()) ?: return false
            val myfmt = fmt.format
            for (x in FormatConstants.CURRENCY_FORMATS.indices) {
                if (FormatConstants.CURRENCY_FORMATS[x][0] == myfmt) {
                    return true
                }
            }
            return false
        }

    /**
     * get the Pattern Color index for this Cell if Solid Fill, or the Foreground color if no Solid Pattern
     */
    val foregroundColor: Short
        get() = if (fill != null) fill!!.getFgColorAsInt(workBook!!.theme).toShort() else this.icvFore

    /**
     * get the Pattern Color for this Cell if Solid Fill, or the Foreground color if no Solid Pattern, as a Hex Color String
     *
     * @return Hex Color String
     */
    val foregroundColorHEX: String?
        get() = if (fill != null) fill!!.getFgColorAsRGB(workBook!!.theme) else FormatHandle.colorToHexString(FormatHandle.getColor(this.icvFore.toInt()))

    /**
     * get the background Color for this Cell as a Hex Color String
     *
     * @return Hex Color String
     */
    // default background color
    // return white
    val backgroundColorHEX: String?
        get() {
            if (fill != null)
                return fill!!.getBgColorAsRGB(workBook!!.theme)
            return if (this.icvBack.toInt() == 65) "#FFFFFF" else FormatHandle.colorToHexString(FormatHandle.getColor(this.icvBack.toInt()))
        }

    /**
     * get the background Color index for this Cell
     */
    // default background color
    // return white
    val backgroundColor: Short
        get() {
            if (fill != null) return fill!!.getBgColorAsInt(workBook!!.theme).toShort()
            return if (this.icvBack.toInt() == 65) 64 else this.icvBack
        }

    /**
     * get the Formatting for this BiffRec from the pattern
     * match.
     *
     *
     * case insensitive pattern match is performed...
     */
    /**
     * Sets the number format pattern for this format.
     */
    override var formatPattern: String?
        get() {
            if (pat != null) return pat
            val fmts = FormatConstantsImpl.builtinFormats
            for (x in fmts.indices) {
                if (this.ifmt.toInt() == Integer.parseInt(fmts[x][1], 16)) {
                    pat = fmts[x][0]
                    return pat
                }
            }

            val fmt = this.workBook!!.getFormat(ifmt.toInt())
            if (fmt != null) {
                pat = fmt.toString()
                return fmt.format
            }
            return null
        }
        set(pattern) {
            this.pat = pattern

            if (this.workBook == null)
                throw IllegalStateException(
                        "attempting to set format pattern but workbook is null")

            this.setFormat(addFormatPattern(workBook!!, pattern))
        }

    /**
     * get the format pattern for this particular XF
     *
     * @return
     */
    val fillPattern: Int
        get() = if (fill != null) fill!!.fillPatternInt else fls.toInt()


    // BORDER SECTION

    var bottomBorderLineStyle: Short
        get() = this.dgBottom
        set(t) {
            dgBottom = t
            this.updateBorders()
        }

    var topBorderLineStyle: Short
        get() = this.dgTop
        set(t) {
            dgTop = t
            this.updateBorders()
        }

    var leftBorderLineStyle: Short
        get() = this.dgLeft
        set(t) {
            dgLeft = t
            this.updateBorders()
        }

    var rightBorderLineStyle: Short
        get() = this.dgRight
        set(t) {
            dgRight = t
            this.updateBorders()
        }

    /**
     * set the Top Border Color for this Format
     */
    // 20070205 KSC: 64 is automatic border color but should be interpreted as 65
    // 20080118 KSC
    var topBorderColor: Int
        get() = if (icvTop.toInt() == 64) 65 else icvTop.toInt()
        set(t) {
            var t = t
            if (t == 0) t = 64
            icvTop = t.toShort()
            updateBorderColors()
        }

    /**
     * set the Bottom Border Color for this Format
     */
    // 20070205 KSC: 64 is automatic border color but should be interpreted as 65
    // 20080118 KSC
    var bottomBorderColor: Int
        get() = if (icvBottom.toInt() == 64) 65 else icvBottom.toInt()
        set(t) {
            var t = t
            if (t == 0) t = 64
            icvBottom = t.toShort()
            updateBorderColors()
        }

    /**
     * set the Left Border Color for this Format
     */
    // 20070205 KSC: 64 is automatic border color but should be interpreted as 65
    // 20080118 KSC
    var leftBorderColor: Int
        get() = if (icvLeft.toInt() == 64) 65 else icvLeft.toInt()
        set(t) {
            var t = t
            if (t == 0) t = 64
            icvLeft = t.toShort()
            updateBorderColors()
        }

    // 20070205 KSC: 64 is automatic border color but should be interpreted as 65
    val rightBorderColor: Short
        get() = if (icvRight.toInt() == 64) 65 else icvRight

    /**
     * get the diagonal border color
     *
     * @return
     */
    // 20070205 KSC: 64 is automatic border color but should be interpreted as 65
    val diagBorderColor: Short
        get() = if (icvDiag.toInt() == 64) 65 else icvDiag

    val isBackgroundSolid: Boolean
        get() {
            if (fill != null) return fill!!.isBackgroundSolid
            val rkdata = this.getData()
            return rkdata!![17] == PATTERN_SOLID.toByte()
        }

    /**
     * return the indent setting (1=3 spaces)
     *
     * @return
     */
    /**
     * set the indent (1=3 spaces)
     *
     * @param indent
     */
    // indent # = 3 spaces
    // mask = 0xF, 4 bits,
    // 1st 4 bits
    // indent only valid for Left and Right (apparently
    var indent: Int
        get() = cIndent.toInt()
        set(indent) {
            cIndent = indent.toShort()
            var b = (this.getData()!![8] and 0xF0).toByte()
            b = b or cIndent.toByte()
            this.getData()[8] = b
            if (alc.toInt() != FormatConstants.ALIGN_LEFT || alc.toInt() != FormatConstants.ALIGN_RIGHT)
                horizontalAlignment = FormatConstants.ALIGN_LEFT
            setAttributeFlag()
        }

    /**
     * returns true if this style is set to Right-to-Left text direction (reading order)
     *
     * @return
     */
    /**
     * sets the Right to Left Text Direction or reading order of this style
     *
     * @param rtl possible values:
     * <br></br>0=Context Dependent
     * <br></br>1=Left-to-Right
     * <br></br>2=Right-to-Let
     */
    // iReadingOrder= bits 7-6
    // 00= According to Context
    // 01= Left to Right (0x40)
    // 10= Right to Left (0x80)
    var rightToLeftReadingOrder: Int
        get() = iReadingOrder shr 6
        set(rtl) {
            if (rtl == 2)
                iReadingOrder = 0x80
            else if (rtl == 1)
                iReadingOrder = 0x40
            else
                iReadingOrder = 0
            var b = this.getData()!![8]
            b = b or iReadingOrder.toByte()
            this.getData()[8] = b
            this.wkbook!!.updateFormatCache(this)
            setAttributeFlag()
        }

    var horizontalAlignment: Int
        get() = alc.toInt()
        set(hAlign) {
            alc = hAlign.toShort()
            updateAlignment()
            setAttributeFlag()
        }

    var wrapText: Boolean
        get() = fWrap.toInt() == 1
        set(wraptext) {
            if (wraptext) {
                fWrap = 1
            } else {
                fWrap = 0
            }
            updateAlignment()
            setAttributeFlag()
        }

    var verticalAlignment: Int
        get() = alcV.toInt()
        set(vAlign) {
            alcV = vAlign.toShort()
            updateAlignment()
            setAttributeFlag()
        }

    var rotation: Int
        get() = trot.toInt()
        set(rot) {
            trot = rot.toShort()
            updateAlignment()
            setAttributeFlag()
        }

    /**
     * Returns an XML fragment representing the XF backing the format Handle.  The XF record is style information
     * associated with a cell.  Font information/lookup is not included in this output so it can be used as a comparitor
     * style
     */
    // font info...
    // format info, should be expanded prolly
    // 20071218 KSC: Add Fill
    // get the border..
    // get the alignment..
    // get the background color
    val xml: String
        get() {
            val sb = StringBuffer("<XF")
            sb.append(">")
            val myf = this.font
            sb.append("<font name=\"" + myf!!.fontName!!)
            sb.append("\" size=\"" + myf.fontHeightInPoints)
            sb.append("\" color=\"" + FormatHandle.colorToHexString(myf.colorAsColor))
            sb.append("\" weight=\"" + myf.fontWeight)
            if (myf.isBold) {
                sb.append("\" bold=\"1")
            }
            sb.append("\" />")
            sb.append("<format id=\"$ifmt")
            sb.append("\" />")
            sb.append("<fill id=\"$fls")
            sb.append("\" />")
            sb.append("<Borders>")
            if (rightBorderLineStyle.toInt() != 0) {
                sb.append("<Border")
                sb.append(" Position=\"right\"")
                sb.append(" LineStyle=\"" + FormatHandle.BORDER_NAMES[rightBorderLineStyle] + "\"")
                sb.append(" Color=\"" + FormatHandle.colorToHexString(wkbook!!.colorTable[rightBorderColor]) + "\"")
                sb.append("/>")
            }
            if (bottomBorderLineStyle.toInt() != 0) {
                sb.append("<Border")
                sb.append(" Position=\"bottom\"")
                sb.append(" LineStyle=\"" + FormatHandle.BORDER_NAMES[bottomBorderLineStyle] + "\"")
                sb.append(" Color=\"" + FormatHandle.colorToHexString(wkbook!!.colorTable[bottomBorderColor]) + "\"")
                sb.append("/>")
            }
            if (leftBorderLineStyle.toInt() != 0) {
                sb.append("<Border")
                sb.append(" Position=\"left\"")
                sb.append(" LineStyle=\"" + FormatHandle.BORDER_NAMES[leftBorderLineStyle] + "\"")
                sb.append(" Color=\"" + FormatHandle.colorToHexString(wkbook!!.colorTable[leftBorderColor]) + "\"")
                sb.append("/>")
            }
            if (topBorderLineStyle.toInt() != 0) {
                sb.append("<Border")
                sb.append(" Position=\"top\"")
                sb.append(" LineStyle=\"" + FormatHandle.BORDER_NAMES[topBorderLineStyle] + "\"")
                sb.append(" Color=\"" + FormatHandle.colorToHexString(wkbook!!.colorTable[topBorderColor]) + "\"")
                sb.append("/>")
            }
            sb.append("</Borders>")
            sb.append("<Alignment")
            sb.append(" Horizontal=\"" + FormatHandle.HORIZONTAL_ALIGNMENTS[this.horizontalAlignment] + "\"")
            sb.append(" />")
            if (wkbook!!.colorTable[foregroundColor] !== Color.WHITE) {
                sb.append("<Interior Color=\"" +
                        FormatHandle.colorToHexString(wkbook!!.colorTable[foregroundColor]) +
                        "\"/>")
            }

            sb.append("</XF>")
            return sb.toString()
        }

    /**
     * get whether this cell formula is hidden
     *
     * @return
     */
    /**
     * sets the cell formula as hidden
     *
     * @param hd
     */
    var isFormulaHidden: Boolean
        get() = this.fHidden.toInt() == 0x1
        set(hd) {
            if (hd)
                this.fHidden = 0x1
            else
                this.fHidden = 0x0
            updateLockedHidden()
        }

    /**
     * get whether this is a locked Cell
     *
     * @return
     */
    /**
     * sets the cell as locked
     *
     * @param lk
     */
    var isLocked: Boolean
        get() = this.fLocked.toInt() == 0x1
        set(lk) {
            if (lk)
                this.fLocked = 0x1
            else
                this.fLocked = 0x0
            updateLockedHidden()
        }

    /**
     * return whether this cell is set to "shrink to fit"
     *
     * @return
     */
    // turn off bit 4
    // set bit 4
    var isShrinkToFit: Boolean
        get() = this.fShrinkToFit.toInt() == 0x1
        set(b) = if (b) {
            this.fShrinkToFit = 0x1
            this.getData()[9] = this.getData()[9] or 0x10
        } else {
            this.fShrinkToFit = 0x0
            this.getData()[9] = this.getData()[9] and 0xF7
        }


    /**
     * @return
     */
    var stricken: Boolean
        get() = if (myFont != null) myFont!!.stricken else false
        set(b) {
            if (myFont != null) myFont!!.stricken = b
        }


    /**
     * @return
     */
    var italic: Boolean
        get() = if (myFont != null) myFont!!.italic else false
        set(b) {
            if (myFont != null) myFont!!.italic = b
        }


    /**
     * @return
     */
    var underlined: Boolean
        get() = if (myFont != null) myFont!!.underlined else false
        set(b) {
            if (myFont != null) myFont!!.underlined = b
        }


    /**
     * @return
     */
    var bold: Boolean
        get() = if (myFont != null) myFont!!.bold else false
        set(b) {
            if (myFont != null) myFont!!.bold = b
        }

    /**
     * return truth of "this Xf rec is a style xf"
     *
     * @return
     */
    val isStyleXf: Boolean
        get() = fStyle.toInt() == 1

    constructor() {
        //empty constructor
    }

    /**
     * create a new Xf with pointer to its font
     */
    constructor(f: Int) {
        val bl = byteArrayOf(0, 0, 0, 0, 1, 0, 32, 0, 0, 64, 0, 0, 0, 0, 0, 0, 0, 0, -64, 32)
        this.opcode = XLSConstants.XF
        this.length = bl.size.toShort()
        this.setData(bl)
        this.setFont(f)
        this.init()
    }

    /**
     * create a new Xf with pointer to its font and workbook set
     */
    constructor(f: Int, wkbook: WorkBook) {
        val bl = byteArrayOf(0, 0, 0, 0, 1, 0, 32, 0, 0, 64, 0, 0, 0, 0, 0, 0, 0, 0, -64, 32)
        this.opcode = XLSConstants.XF
        this.length = bl.size.toShort()
        this.setData(bl)
        super.workBook = wkbook    // set workbook but don't insert rec or add to xfrecs
        this.setFont(f)
        this.init()
    }

    /**
     * constructor which takes a Font object + a workbook
     * useful for cloning xf's from other workbooks
     *
     * @param f      font
     * @param wkbook
     */
    constructor(f: Font, wkbook: WorkBook) {
        val bl = byteArrayOf(0, 0, 0, 0, 1, 0, 32, 0, 0, 64, 0, 0, 0, 0, 0, 0, 0, 0, -64, 32)
        this.opcode = XLSConstants.XF
        this.length = bl.size.toShort()
        myFont = f
        System.arraycopy(ByteTools.shortToLEBytes(f.idx.toShort()), 0, bl, 0, 2)
        this.setData(bl)
        super.workBook = wkbook    // set workbook but don't insert rec or add to xfrecs
        this.init()
    }

    /**
     * Create a string representation of the Xf
     */
    override fun toString(): String {
        var f: String? = "unknown"        //Handle missing formats
        try {
            f = this.formatPattern
        } catch (e: Exception) {
        }

        val thisToString = " format:" + f + " fill:" + this.fillPattern +
                " fg:" + this.foregroundColor +
                " bg:" + this.backgroundColor +
                " border:[" +
                this.topBorderLineStyle + "-" + this.topBorderColor + ":" +
                this.leftBorderLineStyle + "-" + this.leftBorderColor + ":" +
                this.bottomBorderLineStyle + "-" + this.bottomBorderColor + ":" +
                this.rightBorderLineStyle + "-" + this.rightBorderColor + "]" +
                "W:" + this.wrapText +
                "R:" + this.rotation +
                "H:" + this.horizontalAlignment + "V:" + this.verticalAlignment +
                "I:" + this.indent +
                "L:" + this.isLocked +
                "F:" + this.isFormulaHidden +
                "D:" + this.rightToLeftReadingOrder
        return this.font!!.toString() + thisToString
    }

    /**
     * inc # records using this xf
     */
    fun incUseCount() {
        useCount++
    }

    /**
     * dec # records using this xf
     */
    fun decUseCoount() {
        useCount--
    }

    /**
     * Populates the myFont and myFormat variables to be held onto
     * when the xf record is serialized for boundsheet transfer
     */
    fun populateForTransfer() {
        myFont = this.font
        myFormat = this.workBook!!.getFormat(ifmt.toInt())
        this.getData()
    }

    /**
     * The XF record can either be a style XF or a Cell XF.
     */
    override fun init() {
        super.init()
        ifnt = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        ifnt = (ifnt and 0xffff).toShort()
        ifmt = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        ifmt = (ifmt and 0xffff).toShort()


        val flag = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        // is the cell locked?
        if (flag and 0x1 == 0x1) {
            fLocked = 0x1
        } else {
            fLocked = 0
        }
        // is the cell hidden?
        if (flag and 0x2 == 0x2) {
            fHidden = 0x1
        } else {
            fHidden = 0
        }

        // is it a cell rec or a style rec?
        if (flag and 0x4 == 0x4) {
            fStyle = 1
        } else {
            fStyle = 0
        }
        if (flag and 0x8 == 0x0008)
            f123Prefix = 0x1
        ixfParent = (flag and 0xFFF0 shr 4).toShort()


        initXF()

        pat = null    // ensure reset if xf has changed
        if (DEBUGLEVEL > XLSConstants.DEBUG_LOW)
            Logger.logInfo("Xf.init() ifnt: " + ifnt
                    + " ifmt: " + ifmt + ":" +
                    this.toString()
                    + " border: "
                    + "l:" + this.leftBorderColor + ":"
                    + "b:" + this.bottomBorderColor + ":"
                    + "r:" + this.rightBorderColor + ":"
                    + "t:" + this.topBorderColor + ":")
    }

    /**
     * read and interpret bytes 6-18)
     */
    internal fun initXF() {
        var flag: Short

        // bytes 6, 7: alignment, rotation, text break
        flag = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        alc = (flag and 0x7).toShort()
        if (flag and 0x8 == 0x8) fWrap = 1
        alcV = (flag and 0x70 shr 4).toShort()
        trot = (flag and 0xFF00 shr 8).toShort()

        // byte 8: indent, reading order, shrink
        flag = this.getByteAt(8).toShort()
        cIndent = (flag and 0xF).toShort()
        if (flag and 0x10 == 0x10) fShrinkToFit = 1
        if (flag and 0x20 == 0x20) fMergeCell = 1
        if (DEBUGLEVEL > 5) Logger.logInfo("Xf The merge cell bit is: $fMergeCell and the int is $flag")

        iReadingOrder = (flag and 0xC0).toShort()// >> 6);	// reading order is byte 7-6 mask 0xCO
        // USED_ATTRIB:	 bits 7-2 of byte 9
        flag = this.getByteAt(9).toShort()
        /* for all these flags, a cleared bit means use Parent Style XF attribute
		 if set, means the attributes of THIS xf is used
		bit mask 	meaning
		0 	01H 	Flag for number format
		1 	02H 	Flag for font
		2 	04H 	Flag for horizontal and vertical alignment, text wrap, indentation, orientation, rotation, and
		text direction
		3 	08H 	Flag for border lines
		4 	10H 	Flag for background area style
		5 	20H 	Flag for cell protection (cell locked and formula hidden)
		 */
        if (flag and 0x4 == 0x4) fAtrNum = 1        // number format
        if (flag and 0x8 == 0x8) fAtrFnt = 1        // font
        if (flag and 0x10 == 0x10) fAtrAlc = 1    // alignment (h + v) text wrap rotation direction indent
        if (flag and 0x20 == 0x20) fAtrBdr = 1    // border lines
        if (flag and 0x40 == 0x40) fAtrPat = 1    // background format pattern
        if (flag and 0x80 == 0x80) fAtrProt = 1    // cell protection

        // BORDER Section
        flag = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
        dgLeft = (flag and 0xF).toShort()
        dgRight = (flag and 0xF0 shr 4).toShort()
        dgTop = (flag and 0xF00 shr 8).toShort()
        dgBottom = (flag and 0xF000 shr 12).toShort()

        flag = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        icvLeft = (flag and 0x7f).toShort()
        icvRight = (flag and 0x3F80 shr 7).toShort()
        grbitDiag = (flag and 0xC000 shr 15).toShort()

        // bytes 14-17 color and fill
        Iflag = ByteTools.readInt(this.getByteAt(14), this.getByteAt(15), this.getByteAt(16), this.getByteAt(17))
        icvTop = (Iflag and 0x7F).toShort()
        icvBottom = (Iflag and 0x3F80 shr 7).toShort()
        icvDiag = (Iflag and 0x1FC000 shr 14).toShort()
        diagBorderLineStyle = (Iflag and 0x1E00000 shr 21).toShort()
        mystery = (Iflag and 0x3800000 shr 25).toByte()
        fls = (Iflag and -0x4000000 shr 26).toShort() // fill pattern

        if (DEBUGLEVEL > 5 && icvTop > 0) Logger.logInfo("Xf The cell outline is true")
        // bytes 18, 19: fill pattern colors
        icvColorFlag = ByteTools.readShort(this.getByteAt(18).toInt(), this.getByteAt(19).toInt())
        icvFore = (icvColorFlag and 0x7F).toShort()                // = Pattern Color
        icvBack = (icvColorFlag and 0x3F80 shr 7).toShort()    // = Pattern Background Color
        if (icvColorFlag and 0x4000 == 0x4000) fSxButton = 1

        // Logger.logInfo(io.starter.OpenXLS.ExcelTools.getRecordByteDef(this));
    }

    /**
     * set the pointer to the XF's Format in the WorkBook
     */
    fun setFormat(ifm: Short) {
        ifmt = ifm
        val nef = ByteTools.shortToLEBytes(ifmt)
        this.getData()[2] = nef[0]
        this.getData()[3] = nef[1]
        this.pat = null    // 20080228 KSC: flag to re-input
    }


    /**
     * set the pointer to the XF's Font in the WorkBook
     */
    fun setFont(ifn: Int) {
        ifnt = ifn.toShort()
        val nef = ByteTools.shortToLEBytes(ifnt)
        this.getData()[0] = nef[0]
        this.getData()[1] = nef[1]
        // reset the pointer for xf's and font's brought from other workbooks.
        if (this.workBook != null) {
            myFont = this.workBook!!.getFont(ifn)
        }
    }

    /**
     * Set myFont in XF to the same as Workbook's
     */
    fun setMyFont(f: Font) {
        myFont = f
    }

    /**
     * set the Fill Pattern for this Format
     */
    fun setPattern(t: Int) {
        fls = t.toShort()
        this.updatePattern()
        if (fill != null)
            fill!!.setFillPattern(t)
    }

    /**
     * set the Right Border Color for this Format
     */
    fun setRightBorderColor(t: Int) {
        var t = t
        if (t == 0) t = 64    // 20080118 KSC
        icvRight = t.toShort()
        updateBorderColors()

    }

    /**
     * set the diagonal Border Color for this Format
     */
    fun setDiagBorderColor(t: Int) {
        var t = t
        if (t == 0) t = 64    // 20080118 KSC
        icvDiag = t.toShort()
        updateBorderColors()

    }

    /**
     * set the Left Border Color for this Format
     */
    fun setLeftBorderColor(t: Short) {
        var t = t
        if (t.toInt() == 0) t = 64    // 20080118 KSC
        icvLeft = t
        updateBorderColors()
    }

    /**
     * set the diagonal border for this Format
     */
    fun setBorderDiag(t: Int) {
        Iflag = 0
        Iflag = Iflag or icvTop.toInt()
        Iflag = Iflag or (icvBottom shl 7)
        Iflag = Iflag or (t.toShort() shl 14)
        Iflag = Iflag or (diagBorderLineStyle shl 21)
        Iflag = Iflag or (mystery.toShort() shl 25)
        Iflag = Iflag or (fls shl 26)
        this.updatePattern()
    }

    /**
     * set the border line style for this Format
     */
    fun setBorderLineStyle(t: Short) {
        dgLeft = t
        dgRight = t
        dgTop = t
        dgBottom = t
        this.updateBorders()
    }

    /**
     * set border line styles via array of ints representing border styles
     * order= top, left, bottom, right [diagonal]
     *
     * @param b int[]
     */
    fun setAllBorderLineStyles(b: IntArray) {
        try {
            if (b[0] > -1) dgTop = b[0].toShort()
            if (b[1] > -1) dgLeft = b[1].toShort()
            if (b[2] > -1) dgBottom = b[2].toShort()
            if (b[3] > -1) dgRight = b[3].toShort()
            if (b[4] > -1) diagBorderLineStyle = b[4].toShort()
        } catch (e: ArrayIndexOutOfBoundsException) {
        }

        this.updateBorders()
    }

    /**
     * set all border colors via an array of ints representing border color ints
     * order= top, left, bottom, right, [diagonal]
     *
     * @param b int[]
     */
    fun setAllBorderColors(b: IntArray) {
        try {
            if (b[0] > -1) icvTop = b[0].toShort()
            if (b[1] > -1) icvLeft = b[1].toShort()
            if (b[2] > -1) icvBottom = b[2].toShort()
            if (b[3] > -1) icvRight = b[3].toShort()
            if (b[4] > -1) icvDiag = b[4].toShort()
        } catch (e: ArrayIndexOutOfBoundsException) {
        }

        this.updateBorderColors()
    }

    fun updateBorders() {
        var borderflag: Short = 0
        borderflag = dgLeft
        borderflag = (borderflag or (dgRight shl 4)).toShort()
        borderflag = (borderflag or (dgTop shl 8)).toShort()
        borderflag = (borderflag or (dgBottom shl 12)).toShort()
        //byte[] rkdata = this.getData();
        val bords = ByteTools.shortToLEBytes(borderflag)
        this.getData()[10] = bords[0]
        this.getData()[11] = bords[1]
        setAttributeFlag()
    }

    /**
     * removes all borders for the style
     */
    fun removeBorders() {
        this.dgBottom = 0
        this.dgTop = 0
        this.diagBorderLineStyle = 0
        this.dgLeft = 0
        this.dgRight = 0
        this.dgBottom = 0
        this.updateBorders()
    }

    fun updateBorderColors() {
        setAttributeFlag()
    }

    fun updatePattern() {
        val rkdata = this.getData()
        var thisFlag: Short = 0
        thisFlag = thisFlag or icvLeft
        thisFlag = thisFlag or (icvRight shl 7).toShort()
        thisFlag = thisFlag or (grbitDiag shl 14).toShort()
        val bytes = ByteTools.shortToLEBytes(thisFlag)
        rkdata[12] = bytes[0]
        rkdata[13] = bytes[1]

        Iflag = 0
        Iflag = Iflag or icvTop.toInt()
        Iflag = Iflag or (icvBottom shl 7)
        Iflag = Iflag or (icvDiag shl 14)
        Iflag = Iflag or (diagBorderLineStyle shl 21)
        Iflag = Iflag or (mystery.toShort() shl 25)
        Iflag = Iflag or (fls shl 26)
        val nef = ByteTools.cLongToLEBytes(Iflag)
        rkdata[14] = nef[0]
        rkdata[15] = nef[1]
        rkdata[16] = nef[2]
        rkdata[17] = nef[3]
        // update format cache upon change
        pat = null
        this.wkbook!!.updateFormatCache(this)
    }

    /**
     * set the Foreground Color for this Format
     * THIS SETS THE BACKGROUND COLOR when PATTERN (fls) = PATTERN_SOLID
     * THIS SETS THE PATTERN COLOR when PATTERN (fls) > PATTERN_SOLID
     * <br></br>"If the fill style is solid: When solid is specified, the
     * foreground color (fgColor) is the only color rendered,
     * even when a background color (bgColor) is also specified"
     * icvFore==Pattern Background Color
     *
     * @param clr java.awt.Color or null if use standard Excel 2003 Color Table
     * @param    t    best match index into 2003-style Color tabe
     */
    fun setForeColor(t: Int, clr: Color?) {
        icvColorFlag = 0
        icvColorFlag = icvColorFlag or t.toShort()
        icvColorFlag = icvColorFlag or (icvBack shl 7).toShort()
        if (clr != null) {
            if (clr != FormatHandle.COLORTABLE[t]) { // no exact match for color
                if (fill == null)
                    fill = Fill(fillPattern, t, FormatHandle.colorToHexString(clr), icvBack.toInt(), null, this.workBook!!.theme)
                else
                    fill!!.setFgColor(t, FormatHandle.colorToHexString(clr))
            }
        } else if (fill != null)
            fill!!.setFgColor(t)
        this.updateColors()
    }

    /**
     * set the Background Color for this Format (when PATTERN - fls != PATTERN_SOLID)
     * When PATTERN is PATTERN_SOLID, == 64
     *
     * @param clr java.awt.Color or null if use standard Excel 2003 Color Table
     * @param    t    best-match index into 2003-style Color table
     */
    fun setBackColor(t: Int, clr: Color?) {
        icvColorFlag = 0
        icvColorFlag = icvColorFlag or icvFore
        icvColorFlag = icvColorFlag or (t.toShort() shl 7).toShort()
        if (clr != null) {
            if (clr != FormatHandle.COLORTABLE[t]) { // no exact match for color - store custom color
                if (fill == null)
                    fill = Fill(fillPattern, icvFore.toInt(), null, t, FormatHandle.colorToHexString(clr), this.workBook!!.theme)
                else
                    fill!!.setBgColor(t, FormatHandle.colorToHexString(clr))
            }
        } else if (fill != null)
            fill!!.setBgColor(t)

        this.updateColors()
    }

    internal fun updateColors() {
        val rkdata = this.getData()
        val nef = ByteTools.shortToLEBytes(icvColorFlag)
        rkdata[18] = nef[0]
        rkdata[19] = nef[1]
        icvFore = (icvColorFlag and 0x7F).toShort()
        icvBack = (icvColorFlag and 0x3F80 shr 7).toShort()
        // update format cache upon change
        pat = null
        this.wkbook!!.updateFormatCache(this)
    }

    /**
     * Sets the fill pattern to solid, which renders the background to 64=="the default fg color"
     * "If the fill style is solid: When solid is specified, the
     * foreground color (fgColor) is the only color rendered,
     * even when a background color (bgColor) is also
     * specified"
     */
    fun setBackgroundSolid() {
        setPattern(PATTERN_SOLID)
        setBackColor(64, null)
        if (fill != null) fill!!.setFillPattern(PATTERN_SOLID)
    }


    /**
     * Sets the attribute flags for this xf record.  These flags consist of
     * // bit 8= fAtrProt
     * //     7= fAtrPat
     * //     6= fAtrBdr
     * //     5= fAtrAlc (Alignment)
     * //     4= fAtrFnt
     * //     3= fAtrNum
     */
    private fun setAttributeFlag() {
        setToCellXF()
        val rkdata = this.getData()
        var used_attrib = rkdata!![9]
        val borderFlag = (if (dgBottom > 0 || dgTop > 0 || dgLeft > 0 || dgRight > 0 || diagBorderLineStyle > 0) 1 else 0).toByte()    // if border is set
        if (borderFlag.toInt() == 1)
            used_attrib = (used_attrib or 0x20).toByte()    // set bit # 6
        else
            used_attrib = (used_attrib and 0xDF).toByte()    // clear it
        if (cIndent.toInt() != 0 || iReadingOrder.toInt() != 0 || alc.toInt() != 0 || alcV.toInt() != 0 || fWrap.toInt() != 0 || trot.toInt() != 0)
        // set bit # 5
            used_attrib = (used_attrib or 0x10).toByte()
        else
            used_attrib = (used_attrib and 0xEF).toByte()  // clear it

        rkdata[9] = used_attrib
        fAtrNum = (if (used_attrib and 0x04 == 0x04) 1 else 0).toShort()
        fAtrFnt = (if (used_attrib and 0x08 == 0x08) 1 else 0).toShort()
        fAtrAlc = (if (used_attrib and 0x10 == 0x10) 1 else 0).toShort()
        fAtrBdr = (if (used_attrib and 0x20 == 0x20) 1 else 0).toShort()
        fAtrPat = (if (used_attrib and 0x40 == 0x40) 1 else 0).toShort()
        fAtrProt = (if (used_attrib and 0x80 == 0x80) 1 else 0).toShort()


        // must set color flag for borders or Excel will not like [BugTracker 2861]
        if (dgTop > 0 && icvTop.toInt() == 0)
            icvTop = 64
        if (dgBottom > 0 && icvBottom.toInt() == 0)
            icvBottom = 64
        if (dgRight > 0 && icvRight.toInt() == 0)
            icvRight = 64
        if (dgLeft > 0 && icvLeft.toInt() == 0)
            icvLeft = 64
        if (diagBorderLineStyle > 0 && icvDiag.toInt() == 0)
            icvDiag = 64
        this.updatePattern()
    }

    /**
     * Switch the record to a cell XF record
     */
    fun setToCellXF() {
        if (fStyle.toInt() != 0) {// must set to cell xf (fStyle==0) as changes will not show [BugTracker 2861]
            fStyle = 0
            var flag = fLocked.toByte()
            flag = (flag or (fHidden shl 1)).toByte()
            this.getData()[4] = flag
            this.getData()[5] = 0   // upper bits are style parent rec index
        }
    }

    private fun updateAlignment() {
        //short tempAlc = (short)(alc << 3);
        val tempfWrap = (fWrap shl 3).toShort()
        val tempAlcV = (alcV shl 4).toShort()
        val tempTrot = (trot shl 8).toShort()
        var res: Short = 0x0
        res = (res or alc).toShort()
        res = (res or tempfWrap).toShort()
        res = (res or tempAlcV).toShort()
        res = (res or tempTrot).toShort()
        val rkdata = this.getData()
        val bords = ByteTools.shortToLEBytes(res)
        rkdata[6] = bords[0]
        rkdata[7] = bords[1]
        // update format cache upon change
        this.wkbook!!.updateFormatCache(this)
    }

    /**
     * @return Returns the myFormat.
     */
    fun getFormat(): Format? {
        return myFormat
    }

    /**
     * @param myFormat The myFormat to set.
     */
    fun setFormat(myFormat: Format) {
        this.myFormat = myFormat
    }


    /**
     * 2 2 XF type, cell protection, and parent style XF:
     * Bit Mask Contents
     * 2-0 0007H XF_TYPE_PROT – XF type, cell protection (see above)
     * 15-4 FFF0H Index to parent style XF (always FFFH in style XFs)
     *
     *
     * Bit Mask Contents
     * 0 01H 1 = Cell is locked
     * 1 02H 1 = Formula is hidden
     * 2 04H 0 = Cell XF; 1 = Style XF
     */
    private fun updateLockedHidden() {

        val tempFL = (fLocked shl 0x0).toShort()
        val tempFH = (fHidden shl 0x1).toShort()
        val tempST = (fStyle shl 0x2).toShort()

        var flag: Short = 0x0
        flag = (flag or tempFL).toShort()
        flag = (flag or tempFH).toShort()
        flag = (flag or tempST).toShort()

        val dx = this.getData()
        val nef = ByteTools.shortToLEBytes(flag)
        dx[4] = nef[0]
        dx[5] = nef[1]
        // update format cache upon change
        pat = null
        this.wkbook!!.updateFormatCache(this)
    }

    /**
     * set the OOXML fill for this xf
     *
     * @param f
     */
    fun setFill(f: Fill) {
        this.fill = f.cloneElement() as Fill
        fls = this.fill!!.fillPatternInt.toShort()
        icvFore = this.fill!!.getFgColorAsInt(workBook!!.theme).toShort()
        icvBack = this.fill!!.getBgColorAsInt(workBook!!.theme).toShort()
    }

    /**
     * return the OOXML fill for this xf, if any
     */
    fun getFill(): Fill? {
        return this.fill
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        super.close()
        this.myFont!!.close()
        this.myFormat = null
    }

    companion object {

        private val serialVersionUID = -419388613530529316L
        val NDEFAULTXFS = 20

        fun isDatePattern(myfmt: String): Boolean {
            // Search for the format string in the list of known date formats
            for (x in FormatConstants.DATE_FORMATS.indices) {
                if (FormatConstants.DATE_FORMATS[x][0] == myfmt) {
                    return true
                }
            }

            // check for string patterns that only exist within date records (as far as we know, may need refining)
            return myfmt.indexOf("mm") > -1 || myfmt.indexOf("yy") > -1 || myfmt.indexOf("dd") > -1
        }

        /**
         * Parses an escaped xml format pattern (from ooxml) and returns an io.starter.OpenXLS compatible
         * pattern.
         *
         *
         * This method certainly has weaknesses, but my intention is that if it is not a fairly standard format and/or
         * we are not sure how to parse it we should leave the existent format intact soas to not break read/write operations
         *
         *
         * Oddly enough, excel seems to be able to handle biff8 patterns, in my testing so far there has been no need
         * to reencode, that could obviously change...
         *
         * @param xmlFormatPattern
         * @return compatible biff8 formatPattern
         */
        fun unescapeFormatPattern(xmlFormatPattern: String): String {
            var xmlFormatPattern = xmlFormatPattern
            // strip escaping for currency pattern.  Probably should explore all currency types and do an iteration
            xmlFormatPattern = xmlFormatPattern.replace("\"$\"", "$")

            // separator between positive/negative
            xmlFormatPattern = xmlFormatPattern.replace("_);", ";")

            // unescape parens
            xmlFormatPattern = xmlFormatPattern.replace("\\(", "(")
            xmlFormatPattern = xmlFormatPattern.replace("\\)", ")")
            return xmlFormatPattern
        }

        /**
         * Ensures that the given format pattern exists on the given workbook.
         *
         * @param book    the workbook to which the pattern should belong
         * @param pattern the number format pattern to ensure exists
         * @return the format ID of the given format pattern
         */
        fun addFormatPattern(book: WorkBook, pattern: String?): Short {
            var ifmt: Short = -1

            // Look up the pattern on the workbook
            ifmt = book.getFormatId(pattern!!)

            // If the pattern is unknown, create and add a Format record
            if (ifmt.toInt() == -1) {
                val format = Format(book, pattern)
                ifmt = format.ifmt
            }

            return ifmt
        }

        /**
         * PATTERN_SOLID is a special case where icvFore= the background color and icvBack=64.
         * "If the fill style is solid: When solid is specified, the
         * foreground color (fgColor) is the only color rendered,
         * even when a background color (bgColor) is also
         * specified"
         */
        val PATTERN_SOLID = 1    // was set to 4 but tht's wrong!!

        /**
         * clone the xf and add to streamer
         *
         * @param xf
         * @return
         */
        private fun cloneXf(xf: Xf, wkbook: WorkBook): Xf {
            val clone: Xf
            if (xf.idx > -1) {    // it's in the wb already
                clone = Xf(xf.ifnt.toInt(), wkbook)
                val xfbytes = xf.getBytesAt(0, xf.length - 4)
                clone.setData(xfbytes)
                clone.init()
            } else {    // xf hasn't been added to wb yet, no need to clone
                clone = xf
            }
            clone.fill = xf.fill
            clone.setToCellXF()    // changes will not be seen if fstyle bit is set TODO: is this correct in all cases???
            clone.idx = wkbook.insertXf(clone)
            return clone
        }

        /**
         * if xf parameter doesn't exist, create; if it does, create a new xf based on it
         *
         * @param xf      original xf
         * @param fontIdx font to link xf to
         * @param wkbook
         * @return new xf
         */
        fun updateXf(xf: Xf?, fontIdx: Int, wkbook: WorkBook): Xf? {
            var xf = xf
            if (xf == null) {
                xf = Xf(fontIdx, wkbook)
                xf.idx = wkbook.insertXf(xf)    // insert new xf into stream ...
                return xf
            } else {
                xf = Xf.cloneXf(xf, wkbook)
            }
            return xf
        }
    }

}

