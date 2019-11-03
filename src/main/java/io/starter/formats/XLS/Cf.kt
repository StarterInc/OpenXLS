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
import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.Dxf
import io.starter.formats.OOXML.Fill
import io.starter.formats.XLS.formulas.FormulaParser
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgArray
import io.starter.formats.XLS.formulas.PtgRefN
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList
import java.util.Stack


/**
 * **Cf:  Conditional Formatting Conditions 0x1B1**<br></br>
 *
 *
 * This record stores a conditional formatting condition
 *
 *
 *
 *
 * There are some restrictions in the usage of conditional formattings:
 * In the user interface it is possible to modify the font style (boldness and posture), the text colour, the underline style
 * and the strikeout style. It is not possible to change the used font, the font height, and the escapement style, though it is
 * possible to specify a new font height and escapement style in this record which are correctly displayed
 * It is not possible to change a border line style, but to preserve the line colour, and vice versa. Diagonal lines are not
 * supported at all. The user interface only offers thin line styles, but files containing other line styles work correctly too.
 * It is not possible to set the background pattern colour to No colour (using system window background
 *
 *
 *
 *
 *
 *
 * OFFSET		NAME			SIZE		CONTENTS
 * -----
 * 4				ct					1			Conditional formatting type
 * 5				cp					1			Conditional formatting operator
 * 6				cce1				2			Count of bytes in rgce1	-- size of formula data for first value or formula
 * 8				cce2				2			Count of bytes in rgce2	-- size of formula data for first value or formula: used for second part of 'between' and 'not between' comparison, else 0
 * 10									4			Option flags (see below)
 * 2			Not used
 * 16 								118 		(optional, only if font = 1, see option flags) Font formatting block, see below
 * var 								8			(optional, only if bord = 1, see option flags) Border formatting block, see below
 * var								4 			(optional, only if patt = 1, see option flags) Pattern formatting block, see below
 * var			rgbdxf			var		Conditional format to apply
 * var			rgce1			var		First formula for this condition (RPN token array without size field, ?4)
 * var			rgce2			var		Second formula for this condition (RPN token array without size field, ?4)
 *
 *
 * Conditional formatting operator:
 * 00H = No comparison (only valid for formula type, see above)
 * 01H = Between 05H = Greater than
 * 02H = Not between 06H = Less than
 * 03H = Equal 07H = Greater or equal
 * 04H = Not equal 08H = Less or equal
 *
 *
 * Option Flags
 * If none of the formatting attributes is set, the option flags field contains 00000000H.
 * The following table assumes that
 * the conditional formatting contains at least one modified formatting attribute
 * (it will occur at least one of the formatting
 * information blocks in the record). In difference to the first case some of the
 * bits are always set now.
 * All flags specifying that an attribute is modified are 02, if the conditional formatting
 * changes the respective attribute,
 * and 12, if the original cell formatting is preserved. The flags for modified font
 * attributes are not contained in this
 * option flags field, but in the font formatting block itself. !
 *
 *
 * Bit Mask Contents
 * 9-0 000003FFH Always 11.1111.11112 (but not used)
 * 10 00000400H 0 = Left border style and colour modified (bord - left )
 * 11 00000800H 0 = Right border style and colour modified (bord - right )
 * 12 00001000H 0 = Top border style and colour modified (bord - top )
 * 13 00002000H 0 = Bottom border style and colour modified (bord - bot )
 * 15-14 0000C000H Always 112 (but not used)
 * 16 00010000H 0 = Pattern style modified (patt - style )
 * 17 00020000H 0 = Pattern colour modified (patt - col )
 * 18 00040000H 0 = Pattern background colour modified (patt - bgcol )
 * 21-19 00380000H Always 1112 (but not used)
 * 26 04000000H 1 = Record contains font formatting block (font)
 * 28 10000000H 1 = Record contains border formatting block (bord)
 * 29 20000000H 1 = Record contains pattern formatting block (patt)
 *
 * @see Condfmt
 */
class Cf
//    0x1,0x1,0x3,0x0,0x3,0x0,0xf,0xc,0x3f,0x90,0x2,0x80,0x11,0x11,0x14,0xa,0x14,0xa,0x0,0x0,0x1e,0x7b,0x0,0x1e,0xea,0x0};


/**
 * default constructor
 */
() : io.starter.formats.XLS.XLSRecord() {
    /**
     * return the type of this Cf Rule as an int
     *
     * 1= Cell is ("Cell value is")
     *
     * 2= expression ("formula value is")
     *
     * @return
     */
    var type: Short = 0
        internal set        //		Conditional formatting type
    /**
     * returns the operator or qualifier of this Cf Rue as an int value
     * <br></br>possible values:
     *
     * 0= no comparison
     *
     * 1= between
     *
     * 2= not between
     *
     * 3= equal
     *
     * 4= not equal
     *
     * 5= greater than
     *
     * 6= less than
     *
     * 7= greater than or equal
     *
     * 8= less than or equal
     * ********************
     * 2007-specific operators
     *
     * 9= begins With
     *
     * 10= ends With
     *
     * 11= contains Text
     *
     * 12= not contains
     *
     * @return
     */
    var operator: Short = 0
        internal set        //		Conditional formatting operator
    internal var cce1 = 0        //		Count of bytes in rgce1
    internal var cce2 = 0        //		Count of bytes in rgce2
    internal var rgbdxf = ""        //		Conditional format to apply
    internal var rgce1 = ""            //		First formula for this condition
    internal var rgce2 = ""            // 		Second formula for this condition
    internal var flags = 0                    // option flags 20080303 KSC:
    internal var bHasFontBlock = false        // ""
    internal var bHasBorderBlock = false        // ""
    internal var bHasPatternBlock = false    // ""

    private val prefHolder: Ptg? = null                // this is a placeholder ptg for the conditional format expression
    private var expression1: Stack<*>? = Stack()    //
    private var expression2: Stack<*>? = Stack()    //
    private var formula1: Formula? = null                //
    private var formula2: Formula? = null                //
    private var containsText: String? = null            // OOXML-specific if type==containsText, this will hold the actual text to test

    /**
     * @return Returns the condfmt.
     */
    /**
     * @param condfmt The condfmt to set.
     */
    var condfmt: Condfmt? = null
    private var patternFillStyle = -1
    private var patternFillColorsFlag = 0
    private var patternFillColor = 0
    private var patternFillColorBack = -1
    /**
     * return the 2007v Fill element, or null if not set
     */
    var fill: Fill? = null
        private set
    private var font: Font? = null

    // Offset Size Contents
    // 0 64 Not used
    // 64 4 Font height (in twips = 1/20 of a point); or FFFFFFFFH to preserve the cell font height
    private var fontHeight = -1
    private var fontOptsFlag = -1

    // 72 2 Font weight (100-1000, only if font - style = 0).
    private var fontWeight = -1
    private var fontEscapementFlag = -1
    private var fontUnderlineStyle = -1

    private var fontColorIndex = -1
    private var fontModifiedOptionsFlag = -1
    private var fontEscapementFlagModifiedFlag = -1
    private var fontUnderlineModifiedFlag = -1
    private var borderLineStylesFlag: Short = 0
    private var borderLineStylesLeft = -1
    private var borderLineStylesRight = -1
    private var borderLineStylesTop = -1
    private var borderLineStylesBottom = -1
    private var borderLineColorsFlag = 0
    private var borderLineColorLeft = 0
    private var borderLineColorRight = 0
    private var borderLineColorTop = 0
    private var borderLineColorBottom = 0

    // TODO: finish Cf prototype bytes + formatting options such as font block
    //private byte[] PROTOTYPE_BYTES = {1, 5, 3, 0, 0, 0, -1, -1, 59, -92, 2, -128, 0, 0, 2, 0, 2, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 4, 0, 6, -36, 0, -76, 5, 83, -17, -1, -1, -1, -1, 0, 0, 0, 0, -1, -1, -1, -1, -1, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, -97, -1, -1, -1, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, -64, 22, 30, 123, 0};
    // a very basic cf which only contains a pattern fill block and function compares <= 100
    //    private byte[] PROTOTYPE_BYTES= {1, 8, 3, 0, 0, 0, -1, -1, 59, -96, 2, -128, 0, 0, -64, 26, 30, 100, 0};
    private val PROTOTYPE_BYTES = byteArrayOf(1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)//,  -64, 26, 30, 100, 0};

    /**
     * reutrn the string representation of this Cf Rule type
     *
     * @return
     */
    val typeString: String
        get() = Cf.translateType(this.type.toInt())

    /**
     * @return Returns the fontEscapementFlag.
     */
    /**
     * @param fontEscapementFlag The fontEscapementFlag to set.
     */
    // set modified
    //		if (font!=null)
    //			font.setScript(ss);
    var fontEscapement: Int
        get() = fontEscapementFlag
        set(fontEscapementFlag) {
            this.fontEscapementFlag = fontEscapementFlag
            this.fontEscapementFlagModifiedFlag = 0
            bHasFontBlock = true
        }


    /**
     * returns the pattern color, if any, as an HTML color String.  Includes custom OOXML colors.
     *
     * @return String HTML Color String
     */
    val patternFgColor: String?
        get() {
            if (fill != null)
                return fill!!.getFgColorAsRGB(workBook!!.theme)
            return if (patternFillColor != -1) FormatHandle.colorToHexString(this.colorTable[patternFillColor]) else null
        }

    /**
     * returns the pattern background color, if any, as an HTML color String.  Includes custom OOXML colors.
     *
     * @return String HTML Color String
     */
    val patternBgColor: String?
        get() {
            if (fill != null)
                return if (patternFillStyle == 1)
                    fill!!.getFgColorAsRGB(workBook!!.theme)
                else
                    fill!!.getBgColorAsRGB(workBook!!.theme)
            return if (patternFillColorBack != -1) FormatHandle.colorToHexString(this.colorTable[patternFillColorBack]) else null
        }

    val borderColors: Array<java.awt.Color>?
        get() = if (!this.bHasBorderBlock) null else arrayOf(colorTable[getBorderLineColorTop()], colorTable[getBorderLineColorLeft()], colorTable[getBorderLineColorBottom()], colorTable[getBorderLineColorRight()])


    val allBorderColors: IntArray?
        get() = if (!this.bHasBorderBlock) null else intArrayOf(getBorderLineColorTop(), getBorderLineColorLeft(), getBorderLineColorBottom(), getBorderLineColorRight())

    /**
     * order= top, left, bottom, right
     *
     * @return
     */
    // diag
    val borderStyles: IntArray?
        get() = if (!this.bHasBorderBlock) null else intArrayOf(getBorderLineStylesTop(), getBorderLineStylesLeft(), getBorderLineStylesBottom(), getBorderLineStylesRight(), -1)

    val borderSizes: IntArray?
        get() {
            if (!this.bHasBorderBlock)
                return null
            val hasTop = this.getBorderLineStylesTop() > 0
            val hasLeft = this.getBorderLineStylesLeft() > 0
            val hasBottom = this.getBorderLineStylesBottom() > 0
            val hasRight = this.getBorderLineStylesRight() > 0
            return intArrayOf(if (hasTop) 1 else 0, if (hasLeft) 1 else 0, if (hasBottom) 1 else 0, if (hasRight) 1 else 0, 0)
        }

    val foregroundColor: Int
        get() {
            if (!this.bHasPatternBlock)
                return -1
            if (fill != null)
                return fill!!.getFgColorAsInt(workBook!!.theme)
            return if (this.patternFillStyle == 1)
                this.patternFillColorBack
            else
                this.patternFillColor
        }


    /**
     * @return Returns the fontOptsPosture.
     */
    /**
     * @param fontOptsPosture The fontOptsPosture to set.
     */
    var fontOptsPosture: Int
        get() = (this.fontOptsFlag and FONT_OPTIONS_POSTURE).toShort().toInt()
        set(fontOptsPosture) {
            this.fontOptsFlag = (this.fontOptsFlag and FONT_OPTIONS_POSTURE).toShort().toInt()
            bHasFontBlock = true
        }

    /**
     * @return Returns the fontOptsCancellation.
     */
    val fontOptsCancellation: Int
        get() = (this.fontOptsFlag and FONT_OPTIONS_CANCELLATION shr 7).toShort().toInt()

    /**
     * @return Returns the fontOptsItalic.
     */
    /**
     * @param fontOptsItalic The fontOptsItalic to set.
     */
    // 1 if italic
    // set italic
    // todo: is below correct?
    // clear italic bit
    var fontItalic: Boolean
        get() = fontOptsFlag and 0x2 == FONT_OPTIONS_POSTURE_ITALIC
        set(italic) {
            if (italic) {
                this.fontOptsFlag = this.fontOptsFlag or 0x2
                this.fontModifiedOptionsFlag = this.fontModifiedOptionsFlag or 0x2
                bHasFontBlock = true
            } else {
                this.fontOptsFlag = this.fontOptsFlag xor 0x2
                this.fontModifiedOptionsFlag = this.fontModifiedOptionsFlag xor 0x2
            }
            if (font != null) font!!.italic = italic
        }

    // todo: is below correct?
    // turn off
    var fontStriken: Boolean
        get() = fontOptsFlag and 0x80 == FONT_OPTIONS_CANCELLATION_ON
        set(bStriken) {
            if (bStriken) {
                fontOptsFlag = fontOptsFlag or 0x80
                fontModifiedOptionsFlag = fontModifiedOptionsFlag or 0x80
                bHasFontBlock = true
            } else {
                fontOptsFlag = fontOptsFlag xor 0x80
                fontModifiedOptionsFlag = fontModifiedOptionsFlag xor 0x80
            }
            if (font != null) font!!.stricken = bStriken
        }

    /**
     * @return Returns true if font escapement is superscript
     */
    val fontEscapementSuper: Boolean
        get() = fontEscapementFlag == FONT_ESCAPEMENT_SUPER

    /**
     * @return Returns true if font escapement is subscript
     */
    val fontEscapementSub: Boolean
        get() = fontEscapementFlag == FONT_ESCAPEMENT_SUB

    private val refPos = -1


    /**
     * Returns the human-readable Condition type
     *
     * @return the condition type for this rule
     */
    //    	 handle evaluated condition
    // No comparison (only valid for formula type, see above)
    // okay annoying, but apparenlty there is a ptgRef to A1 in these that should
    // be replaced with our ptg... whatever!!
    // Between
    // expression2 for the other bounds ...
    // Greater than
    // Not between
    // Less than
    // Equal
    // Greater or equal
    // Not equal
    // Less or equal
    // 2007-Specific
    val conditionString: String
        get() {
            when (this.operator) {
                0x0 -> return this.expression1!!.toString() + this.expression2!!.toString()

                1 -> return "Between"

                0x5 -> return "Greater Than"

                0x2 -> return "Not Between"

                0x6 -> return "Less Than"

                0x3 -> return "Equals"

                0x7 -> return "Greater Than or Equal"

                0x4 -> return "Not Equal"

                0x8 -> return "Less Than or Equal"
                0x9 -> return "Begins With"
                0xA -> return "Ends With"
                0xB -> return "Contains Text"
                0xC -> return "Not Contains"
                else -> return "Unknown"
            }
        }

    /**
     * returns EXML (XMLSS) for the Conditional Format Rule
     *
     *
     *
     * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
     * <Range>R12C2:R16C2</Range>
     * <Condition>
     * <Qualifier>Between</Qualifier>
     * <Value1>2</Value1>
     * <Value2>4</Value2>
     * <Format Style='color:#002060;font-weight:700;text-line-through:none;
    border:.5pt solid windowtext;background:#00B0F0'></Format>
    </Condition> *
    </ConditionalFormatting> *
     *
     * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
     * <Range>R6C2</Range>
     * <Condition>
     * <Value1>NOT(ISERROR(SEARCH(&quot;yes&quot;,RC)))</Value1>
     * <Format Style='color:#9C0006;background:#FFC7CE'></Format>
    </Condition> *
    </ConditionalFormatting> *
     *
     * @return
     */
    val xml: String
        get() = getXML(this)

    /**
     * constructor which takes the cellrange of cells
     * and the Condfmt that reference this conditional format rule
     *
     * @param f the condfmt
     * @param r the cellrange
     */
    constructor(f: Condfmt) : this() {
        setData(this.PROTOTYPE_BYTES)
        this.condfmt = f
    }


    /** Pattern Formatting Block  */
    /**
     * initialize the Cf record
     */
    override fun init() {
        super.init()
        data = this.getData()
        type = this.getByteAt(0).toShort()
        operator = this.getByteAt(1).toShort()
        cce1 = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
        cce2 = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()
        //parsing of formula refs
        flags = ByteTools.readInt(this.getByteAt(6), this.getByteAt(7), this.getByteAt(8), this.getByteAt(9))
        // Font formatting Block
        bHasFontBlock = flags and 0x04000000 == 0x04000000
        // Border Formatting Block
        bHasBorderBlock = flags and 0x10000000 == 0x10000000
        // Pattern Formating Block
        bHasPatternBlock = flags and 0x20000000 == 0x20000000
        var pos = 12

        if (bHasFontBlock) { // handle Font formatting section
            pos += 64 // 1st 64 bits of font block is unused
            // 64 4 Font height (in twips = 1/20 of a point); or FFFFFFFFH to preserve the cell font height
            fontHeight = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))

            // 68 4 Font options:
            // Bit Mask Contents
            // 1 00000002H Posture: 0 = Normal; 1 = Italic (only if font - style = 0)
            // 7 00000080H Cancellation: 0 = Off; 1 = On (only if font - canc = 0)
            fontOptsFlag = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))

            // 72 2 Font weight (100-1000, only if font - style = 0).
            // Standard values are 0190H (400) for normal text and 02BCH (700) for bold text.
            fontWeight = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt()

            // 74 2 Escapement type (only if font - esc = 0):
            // 0000H = None; 0001H = Superscript; 0002H = Subscript
            // FONT_ESCAPEMENT_NONE 	= 0x0;
            // FONT_ESCAPEMENT_SUPER 	= 0x1;
            // FONT_ESCAPEMENT_SUB 	= 0x2;
            fontEscapementFlag = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt()


            // 76 1 Underline type (only if font - underl = 0):
            // 00H = None
            // 01H = Single 21H = Single accounting
            // 02H = Double 22H = Double accounting
            // FONT_UNDERLINE_NONE 	= 0x0;
            // FONT_UNDERLINE_SINGLE 	= 0x1;
            // FONT_UNDERLINE_DOUBLE 	= 0x2;
            fontUnderlineStyle = getByteAt(pos++).toInt()

            // 77 3 Not used
            pos += 3

            // 80 4 Font colour index (?6.70); or FFFFFFFFH to preserve the cell font colour
            fontColorIndex = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))

            // 84 4 Not used
            pos += 4

            // 88 4 Option flags for modified font attributes:
            // Bit Mask Contents
            // 1 00000002H 0 = Font style (posture or boldness) modified (font - style )
            // 4-3 00000018H Always 112 (but not used)
            // 7 00000080H 0 = Font cancellation modified (font - canc )
            fontModifiedOptionsFlag = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))

            // 92 4 0 = Escapement type modified (font - esc )
            fontEscapementFlagModifiedFlag = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))

            // 96 4 0 = Underline type modified (font - underl )
            fontUnderlineModifiedFlag = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))
            // 100 16 Not used
            pos += 16
            // 116 2 0001H
            pos += 2
            if (pos != 130) {
                Logger.logWarn("Cf font block parsing pos mismatch$pos")
            }
            getFont()
        }
        if (bHasBorderBlock) {

            //	Offset Size Contents
            //	0 2 Border line styles:
            //	Bit Mask Contents
            //	3-0 000FH Left line style (only if bord - left = 0, ?3.10)
            //	7-4 00F0H Right line style (only if bord - right = 0, ?3.10)
            //	11-8 0F00H Top line style (only if bord - top = 0, ?3.10)
            //	15-12 F000H Bottom line style (only if bord - bot = 0, ?3.10)
            // BORDER_LINESTYLE_LEFT 	= 0x000F;
            // BORDER_LINESTYLE_RIGHT 	= 0x00F0;
            // BORDER_LINESTYLE_TOP 	= 0x0F00;
            // BORDER_LINESTYLE_BOTTOM = 0xF000;
            borderLineStylesFlag = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt())
            updateBorderLineStyles()

            // 2 4 Border line colour indexes:
            // 	Bit Mask Contents
            //	6-0 0000007FH Colour index (?6.70) for left line (only if bord - left = 0)
            //	13-7 00003F80H Colour index (?6.70) for right line (only if bord - right = 0)
            //	22-16 007F0000H Colour index (?6.70) for top line (only if bord - top = 0)
            //	29-23 3F800000H Colour index (?6.70) for bottom line (only if bord - bot = 0)
            // BORDER_LINECOLOR_LEFT 	= 0x0000007F;
            // BORDER_LINECOLOR_RIGHT 	= 0x00003F80;
            // BORDER_LINECOLOR_TOP 	= 0x007F0000;
            // BORDER_LINECOLOR_BOTTOM = 0x3F800000;
            borderLineColorsFlag = ByteTools.readInt(getByteAt(pos++), getByteAt(pos++), getByteAt(pos++), getByteAt(pos++))
            updateBorderLineColors()

            //6 2 Not used
            pos += 2
        }
        if (bHasPatternBlock) {
            // Offset Size Contents

            // 0 2 Fill pattern style:
            // Bit Mask Contents
            // 15-10 FC00H Fill pattern style (only if patt - style = 0, ?3.11)
            // PATTERN_FILL_STYLE = 0xFC00;
            patternFillStyle = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()) shr 9
            /* very strangely, patternFillStyle appears to be 0 when it should be 1 (solid)
             * found in testing through our test suite ...
             * according to documetation this algorithm *appears* ok but the doc is really confusing on cf's
             */

            // 2 2 Fill pattern colour indexes:

            // Bit Mask Contents
            // 6-0 007FH Colour index (?6.70) for pattern (only if patt - col = 0)
            // 13-7 3F80H Colour index (?6.70) for pattern background (only if patt - bgcol = 0)
            // PATTERN_FILL_COLOR 		= 0x007F;
            // PATTERN_FILL_BACK_COLOR 	= 0x3F80;
            patternFillColorsFlag = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt()

            // update the current vals
            updatePatternFillColors()

        }

        var postest = 12
        if (bHasFontBlock)
            postest += 118
        if (bHasBorderBlock)
            postest += 8
        if (bHasPatternBlock)
            postest += 4

        if (postest != pos) {
            Logger.logWarn("Cf bad pos offset during init().")
            pos = postest
        }
        // 1st formula data= pos->cce1
        var function = this.getBytesAt(pos, cce1)
        try {
            expression1 = ExpressionParser.parseExpression(function!!, this, cce1)
        } catch (e: Exception) {
            Logger.logErr("Initializing expression1 for Cf failed: " + String(function!!))
        }

        pos += cce1
        // 2nd formula data= pos+cce1->cce2
        function = this.getBytesAt(pos, cce2)
        if (cce2 > 0) {
            try {
                expression2 = ExpressionParser.parseExpression(function!!, this, cce2)
            } catch (e: Exception) {
                Logger.logErr("Initializing expression2 for Cf failed: " + String(function!!))
            }

        }
        if (DEBUGLEVEL > XLSConstants.DEBUG_LOW)
            Logger.logInfo("Cf record encountered.")
    }

    /**
     * take current state of Cf and update record
     */
    private fun updateRecord() {
        var newdata = ByteArray(12) // enough room for basics; formatting blocks will be appended
        newdata[0] = type.toByte()
        newdata[1] = operator.toByte()
        flags = 0x3FF or 0xC000 or 0x380000    // set required and unused bits of flag
        if (bHasFontBlock)
            flags = flags or 0x04000000
        if (bHasBorderBlock) {
            flags = flags or 0x10000000
            flags = flags or 0x400    // left border mod
            flags = flags or 0x800    // right border mod
            flags = flags or 0x1000    // top border mod
            flags = flags or 0x2000    // bottom border mod
        }
        if (bHasPatternBlock) {
            flags = flags or 0x20000000
            flags = flags or 0x10000    // patt style mod
            flags = flags or 0x20000    // patt fg color (pattern color) mod
            flags = flags or 0x40000    // pat bg color mod
        }

        var b = ByteTools.cLongToLEBytes(flags)
        System.arraycopy(b, 0, newdata, 6, 4)

        var pos = 12
        if (bHasFontBlock) { // update font section
            newdata = ByteTools.append(ByteArray(118), newdata)
            pos += 64 // 1st 64 bits of font block is unused
            b = ByteTools.cLongToLEBytes(fontHeight)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            b = ByteTools.cLongToLEBytes(fontOptsFlag)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            b = ByteTools.shortToLEBytes(fontWeight.toShort())
            System.arraycopy(b, 0, newdata, pos, 2)
            pos += 2
            b = ByteTools.shortToLEBytes(fontEscapementFlag.toShort())
            System.arraycopy(b, 0, newdata, pos, 2)
            pos += 2
            b = ByteTools.shortToLEBytes(fontUnderlineStyle.toShort())
            System.arraycopy(b, 0, newdata, pos, 1)
            pos += 1
            // 77 3 Not used
            pos += 3
            b = ByteTools.cLongToLEBytes(fontColorIndex)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            // 84 4 Not used
            pos += 4            // 88:
            b = ByteTools.cLongToLEBytes(fontModifiedOptionsFlag)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            b = ByteTools.cLongToLEBytes(fontEscapementFlagModifiedFlag)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            b = ByteTools.cLongToLEBytes(fontUnderlineModifiedFlag)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            // 100 16 Not used
            pos += 16
            // 116 2 0001H
            pos++
            newdata[pos++] = 1
        }
        if (bHasBorderBlock) {
            newdata = ByteTools.append(ByteArray(8), newdata)
            b = ByteTools.shortToLEBytes(borderLineStylesFlag)
            System.arraycopy(b, 0, newdata, pos, 2)
            pos += 2
            b = ByteTools.cLongToLEBytes(borderLineColorsFlag)
            System.arraycopy(b, 0, newdata, pos, 4)
            pos += 4
            //6 2 Not used
            pos += 2
        }
        if (bHasPatternBlock) {
            newdata = ByteTools.append(ByteArray(4), newdata)
            b = ByteTools.shortToLEBytes((patternFillStyle shl 9).toShort())
            System.arraycopy(b, 0, newdata, pos, 2)
            pos += 2
            b = ByteTools.shortToLEBytes(patternFillColorsFlag.toShort())
            System.arraycopy(b, 0, newdata, pos, 2)
            pos += 2
        }

        if (formula1 != null) {
            val function = getFormulaExpression(formula1!!)
            newdata = ByteTools.append(function, newdata)
            cce1 = function.size
            b = ByteTools.shortToLEBytes(cce1.toShort())
            newdata[2] = b[0]
            newdata[3] = b[1]
        }
        if (formula2 != null) {
            val function = getFormulaExpression(formula2!!)
            newdata = ByteTools.append(function, newdata)
            cce2 = function.size
            b = ByteTools.shortToLEBytes(cce2.toShort())
            newdata[4] = b[0]
            newdata[5] = b[1]
        }
        this.setData(newdata)
        //        this.init(); DO NOT DO AS can overwrite OOXML-specifics
    }

    /**
     * return the expression bytes of the specified formula
     *
     * @param f
     * @return
     */
    private fun getFormulaExpression(f: Formula): ByteArray {
        var hasArray = false
        var expbytes = ByteArray(0)
        var arraybytes: ByteArray? = null
        val expression = f.expression
        for (i in expression!!.indices) {
            val o = expression.elementAt(i)
            val ptg = o as Ptg
            val b: ByteArray
            if (o is PtgArray) {
                b = o.preRecord
                arraybytes = ByteTools.append(o.postRecord, arraybytes)
                hasArray = true
            } else {
                b = ptg.record
            }
            expbytes = ByteTools.append(b, expbytes)
        }
        if (hasArray) {
            expbytes = ByteTools.append(arraybytes, expbytes)
        }
        return expbytes
    }

    /**
     * sets the operator for this Cf Rule from a String
     *
     * @param s
     */
    fun setOperator(s: String) {
        this.operator = getConditionFromString(s).toShort()
        if (operator.toInt() == 0x0) {    // no comparison
            expression1 = ExpressionParser.parseExpression(s.toByteArray(), this)
        } else {
            // val1 = cfx.getFormula1().calculateFormula();
        }
    }

    /**
     * sets the Operator of this Cf Rule
     * <br></br>possible values:
     *
     * 0= no comparison
     *
     * 1= between
     *
     * 2= not between
     *
     * 3= equal
     *
     * 4= not equal
     *
     * 5= greater than
     *
     * 6= less than
     *
     * 7= greater than or equal
     *
     * 8= less than or equal
     * ********************
     * 2007-specific operators
     *
     * 9= begins With
     *
     * 10= ends With
     *
     * 11= contains Text
     *
     * 12= not contains
     *
     * @param int qualifier - constant value as above
     */
    fun setOperator(op: Int) {
        this.operator = op.toShort()
    }

    /**
     * sets the type of this Cf rule from the int value
     *
     * 1= Cell is ("Cell value is")
     *
     * 2= expression ("formula value is")
     *
     * @param int type as above
     */
    fun setType(type: Int) {
        this.type = type.toShort()
    }

    /**
     * sets the first condition from a String
     *
     * @param s
     */
    fun setCondition1(s: String) {
        var s = s
        // takes "123" aka Value1 and converts to formula expression =123
        s = Cf.unescapeFormulaString(s)
        if (s.indexOf("=") != 0) {
            s = "=$s"
            try {
                val r = intArrayOf(0, 0)
                formula1 = FormulaParser.getFormulaFromString(s, this.sheet, r)
                formula1!!.workBook = this.workBook
                expression1 = formula1!!.expression
            } catch (e: Exception) {
            }

        }
    }

    /**
     * sets the second condition from a String
     *
     * @param s
     */
    fun setCondition2(s: String) {
        var s = s
        if (s == "") return  // none

        s = Cf.unescapeFormulaString(s)
        // takes "123" aka Value1 and converts to formula expression =123
        if (operator.toInt() == 0x0) {    // no comparison
            expression2 = ExpressionParser.parseExpression(s.toByteArray(), this)
        } else {
            if (s.indexOf("=") != 0) {
                s = "=$s"
                try {
                    val r = intArrayOf(0, 0)
                    formula2 = FormulaParser.getFormulaFromString(s, this.sheet, r)
                    expression2 = formula1!!.expression
                } catch (e: Exception) {
                }

            }
        }
        this.cce2 = s.length
    }

    /**
     * parse cf font into it's member elements
     */
    private fun parseFont(font: Font) {
        var z = font.fontHeight
        if (z > 0)
            setFontHeight(z)
        z = font.fontWeight
        if (z > 0)
            setFontWeight(z)
        //		} else if (n.equalsIgnoreCase("font-EscapementFlag")) {
        //			cx.setFontEscapement(s);
        fontStriken = font.stricken
        fontItalic = font.italic
        z = font.color
        if (z > -1)
            setFontColorIndex(z)
        z = font.underlineStyle
        if (z > -1)
            setFontUnderlineStyle(z)
    }


    // Start Font Formatting Block

    /**
     * updates the values for the current border colors from the flag
     * -----------------------------------------------------------
     */
    internal fun updateBorderLineStyles() {
        if (flags and BORDER_MODIFIED_LEFT == 0)
            this.borderLineStylesLeft = (this.borderLineStylesFlag and BORDER_LINESTYLE_LEFT).toShort().toInt()
        if (flags and BORDER_MODIFIED_RIGHT == 0)
            this.borderLineStylesRight = (this.borderLineStylesFlag and BORDER_LINESTYLE_RIGHT shr 4).toShort().toInt()
        if (flags and BORDER_MODIFIED_TOP == 0)
            this.borderLineStylesTop = (this.borderLineStylesFlag and BORDER_LINESTYLE_TOP shr 8).toShort().toInt()
        if (flags and BORDER_MODIFIED_BOTTOM == 0)
            this.borderLineStylesBottom = (this.borderLineStylesFlag and BORDER_LINESTYLE_BOTTOM shr 12).toShort().toInt()
        bHasBorderBlock = true
    }

    /**
     * Apply all the borderLineStyle fields into the current border line styles flag
     */
    fun updateBorderLineStylesFlag() {
        flags = flags or 0x10000000        // set flags to denote has a border block
        bHasBorderBlock = flags and 0x10000000 == 0x10000000

        borderLineStylesFlag = 0
        if (borderLineStylesLeft >= 0) borderLineStylesFlag = borderLineStylesLeft.toShort()
        if (borderLineStylesRight >= 0)
            borderLineStylesFlag = (borderLineStylesFlag or (borderLineStylesRight shl 4)).toShort()
        if (borderLineStylesTop >= 0)
            borderLineStylesFlag = (borderLineStylesFlag or (borderLineStylesTop shl 8)).toShort()
        if (borderLineStylesBottom >= 0)
            borderLineStylesFlag = (borderLineStylesFlag or (borderLineStylesBottom shl 12)).toShort()
        /*
        byte[] data = this.getData();
        byte[] updated = ByteTools.shortToLEBytes(borderLineStylesFlag);
        int pos= 12;
        if (bHasFontBlock)
        	pos+=118;
        if (data.length< (pos+2)) {
        	byte[] tmp= new byte[pos+2];
        	System.arraycopy(data, 0, tmp, 0, data.length);
        	this.setData(tmp);
        }
        data[pos++] = updated[0];
        data[pos++] = updated[1];
*/
    }


    //6 2 Not used

    /**
     * updates the values for the current border colors from the flag
     */
    internal fun updateBorderLineColors() {
        if (flags and BORDER_MODIFIED_LEFT == 0)
            this.borderLineColorLeft = (this.borderLineColorsFlag and BORDER_LINECOLOR_LEFT).toShort().toInt()
        if (flags and BORDER_MODIFIED_RIGHT == 0)
            this.borderLineColorRight = (this.borderLineColorsFlag and BORDER_LINECOLOR_RIGHT shr 7).toShort().toInt()
        if (flags and BORDER_MODIFIED_TOP == 0)
            this.borderLineColorTop = (this.borderLineColorsFlag and BORDER_LINECOLOR_TOP shr 16).toShort().toInt()
        if (flags and BORDER_MODIFIED_BOTTOM == 0)
            this.borderLineColorBottom = (this.borderLineColorsFlag and BORDER_LINECOLOR_BOTTOM shr 23).toShort().toInt()
    }

    /**
     * Apply all the borderLineColor fields into the current border line colors flag
     */
    fun updateBorderLineColorsFlag() {
        flags = flags or 0x10000000        // set flags to denote has a border block
        bHasBorderBlock = flags and 0x10000000 == 0x10000000    // = true

        borderLineColorsFlag = 0
        borderLineColorsFlag = borderLineColorsFlag or borderLineColorLeft
        borderLineColorsFlag = borderLineColorsFlag or (borderLineColorRight shl 7)
        borderLineColorsFlag = borderLineColorsFlag or (borderLineColorTop shl 16)
        borderLineColorsFlag = borderLineColorsFlag or (borderLineColorBottom shl 23)
        /*
        byte[] data = this.getData();
        byte[] updated = ByteTools.cLongToLEBytes( borderLineColorsFlag);
        int pos= 12;
        if (bHasFontBlock)
        	pos+=118;
        pos+=2;	// skip border line style
        if (data.length<(pos+4)) {
        	byte[] tmp= new byte[pos+4];
        	System.arraycopy(data, 0, tmp, 0, data.length);
        	this.setData(tmp);
        }
        data[pos++] = updated[0];
        data[pos++] = updated[1];
        data[pos++] = updated[2];
        data[pos++] = updated[3];
        // no need updateBorderLineColors();
*/
    }

    /**
     * updates the values for the current pattern fill colors from the flag
     *
     *
     * // Bit Mask Contents
     * // 15-10 FC00H Fill pattern style (only if patt - style = 0, ?3.11)
     * // PATTERN_FILL_STYLE = 0xFC00;
     */
    internal fun updatePatternFillColors() {
        /* 		below appears correct in testing		this.patternFillColor = (short)(this.patternFillColorsFlag & PATTERN_FILL_COLOR);
		this.patternFillColorBack = (short)((this.patternFillColorsFlag & PATTERN_FILL_BACK_COLOR) >> 7);
*/
        this.patternFillColor = (this.patternFillColorsFlag and PATTERN_FILL_BACK_COLOR shr 7).toShort().toInt()
        this.patternFillColorBack = (this.patternFillColorsFlag and PATTERN_FILL_COLOR).toShort().toInt()
        bHasPatternBlock = true
    }

    /**
     * Apply all the patternFillColor fields into the current pattern colors flag
     */
    fun updatePatternFillColorsFlag() {
        flags = flags or 0x20000000        // set flags to denote has a pattern block
        bHasPatternBlock = true
        patternFillColorsFlag = patternFillColorBack.toShort().toInt()
        patternFillColorsFlag = (patternFillColorsFlag or (patternFillColor shl 7)).toShort().toInt()
    }

    override fun getFont(): Font? {
        if (!this.bHasFontBlock)
            return null

        if (font != null)
            return font
        var t = this.fontHeight
        var x = this.fontWeight
        if (t == -1)
            t = 180
        else
            t *= 20
        if (x == -1)
            x = Font.PLAIN
        font = Font("Arial", x, t)
        if (fontColorIndex > -1)
            font!!.color = fontColorIndex
        return font
    }

    /**
     * sets the font escapement for this conditional format to superscript
     */
    fun setFontEscapementSuper() {
        this.fontEscapementFlag = FONT_ESCAPEMENT_SUPER
        this.fontEscapementFlagModifiedFlag = 0    //
        this.bHasFontBlock = true
    }

    /**
     * sets the font escapement for this conditional format to subscript
     */
    fun setFontEscapementSub() {
        this.fontEscapementFlag = FONT_ESCAPEMENT_SUB
        this.fontEscapementFlagModifiedFlag = 0    //
        this.bHasFontBlock = true
    }

    /**
     * @return Returns the borderLineStylesLeft.
     */
    fun getBorderLineStylesLeft(): Int {
        return borderLineStylesLeft
    }

    /**
     * @param borderLineStylesLeft The borderLineStylesLeft to set.
     */
    fun setBorderLineStylesLeft(b: Int) {
        this.borderLineStylesLeft = b
        this.updateBorderLineStylesFlag()
    }

    /**
     * @return Returns the borderLineStylesRight.
     */
    fun getBorderLineStylesRight(): Int {
        return borderLineStylesRight
    }

    /**
     * @param borderLineStylesRight The borderLineStylesRight to set.
     */
    fun setBorderLineStylesRight(b: Int) {
        this.borderLineStylesRight = b
        this.updateBorderLineStylesFlag()
    }

    /**
     * @return Returns the borderLineStylesTop.
     */
    fun getBorderLineStylesTop(): Int {
        return borderLineStylesTop
    }

    /**
     * @param borderLineStylesTop The borderLineStylesTop to set.
     */
    fun setBorderLineStylesTop(b: Int) {
        this.borderLineStylesTop = b
        this.updateBorderLineStylesFlag()
    }

    /**
     * @return Returns the borderLineStylesBottom.
     */
    fun getBorderLineStylesBottom(): Int {
        return borderLineStylesBottom
    }

    /**
     * @param borderLineStylesBottom The borderLineStylesBottom to set.
     */
    fun setBorderLineStylesBottom(b: Int) {
        this.borderLineStylesBottom = b
        this.updateBorderLineStylesFlag()
    }

    /**
     * @return Returns the borderLineColorLeft.
     */
    fun getBorderLineColorLeft(): Int {
        if (borderLineColorLeft > this.colorTable.size)
            return 0
        return if (borderLineColorLeft < 0) 0 else borderLineColorLeft
    }

    /**
     * @param borderLineColorLeft The borderLineColorLeft to set.
     */
    fun setBorderLineColorLeft(borderLineColorLeft: Int) {
        this.borderLineColorLeft = borderLineColorLeft
        // 20091028 KSC: insure flag denotes borderlinecolor is modified
        flags = flags and BORDER_MODIFIED_LEFT - 1    // set flags to denote border top is modified (set bit=0)
        this.updateBorderLineColorsFlag()
    }

    /**
     * @return Returns the borderLineColorRight.
     */
    fun getBorderLineColorRight(): Int {
        if (borderLineColorRight > this.colorTable.size)
            return 0
        return if (borderLineColorRight < 0) 0 else borderLineColorRight
    }

    /**
     * @param borderLineColorRight The borderLineColorRight to set.
     */
    fun setBorderLineColorRight(borderLineColorRight: Int) {
        this.borderLineColorRight = borderLineColorRight
        // 20091028 KSC: insure flag denotes borderlinecolor is modified
        flags = flags and BORDER_MODIFIED_RIGHT - 1    // set flags to denote border top is modified (set bit=0)
        this.updateBorderLineColorsFlag()
    }

    /**
     * @return Returns the borderLineColorTop.
     */
    fun getBorderLineColorTop(): Int {
        if (borderLineColorTop > this.colorTable.size)
            return 0
        return if (borderLineColorTop < 0) 0 else borderLineColorTop
    }

    /**
     * @param borderLineColorTop The borderLineColorTop to set.
     */
    fun setBorderLineColorTop(b: Int) {
        this.borderLineColorTop = b
        // 20091028 KSC: insure flag denotes borderlinecolor is modified
        flags = flags and BORDER_MODIFIED_TOP - 1    // set flags to denote border top is modified (set bit=0)
        this.updateBorderLineColorsFlag()
        if (this.borderLineColorTop != b)
            Logger.logWarn("setBorderLineColorTop failed")
    }

    /**
     * @return Returns the borderLineColorBottom.
     */
    fun getBorderLineColorBottom(): Int {
        if (borderLineColorBottom > this.colorTable.size)
            return 0
        return if (borderLineColorBottom < 0) 0 else borderLineColorBottom
    }

    /**
     * @param borderLineColorBottom The borderLineColorBottom to set.
     */
    fun setBorderLineColorBottom(b: Int) {
        this.borderLineColorBottom = b
        // 20091028 KSC: insure flag denotes borderlinecolor is modified
        flags = flags and BORDER_MODIFIED_BOTTOM - 1    // set flags to denote border top is modified (set bit=0)
        this.updateBorderLineColorsFlag()
        if (this.borderLineColorBottom != b)
            Logger.logWarn("borderLineColorBottom failed")
    }

    /**
     * Returns the patternFillStyle.
     * <br></br>NOTE in 2003-ver patternFillStyle is valid 1->
     *
     * @return Returns the patternFillStyle.
     */
    fun getPatternFillStyle(): Int {
        return patternFillStyle
    }

    /**
     * @param patternFillStyle The patternFillStyle to set.
     */
    fun setPatternFillStyle(p: Int) {
        this.patternFillStyle = p
        bHasPatternBlock = true
    }

    /**
     * @return Returns the patternFillColor.
     */
    fun getPatternFillColor(): Int {
        return if (fill != null) fill!!.getFgColorAsInt(workBook!!.theme) else patternFillColor
    }

    /**
     * @param patternFillColor The patternFillColor to set.
     */
    fun setPatternFillColor(p: Int, custom: String?) {
        this.patternFillColor = p
        if (fill != null) fill!!.setFgColor(p)
        this.updatePatternFillColorsFlag()
    }

    /**
     * @return Returns the patternFillColorBack.
     */
    fun getPatternFillColorBack(): Int {
        return if (fill != null) fill!!.getBgColorAsInt(workBook!!.theme) else patternFillColorBack
    }

    /**
     * @param patternFillColorBack The patternFillColorBack to set.
     */
    fun setPatternFillColorBack(p: Int) {
        this.patternFillColorBack = p
        if (fill != null) fill!!.setBgColor(p)
        this.updatePatternFillColorsFlag()
    }

    /**
     * @return Returns the fontHeight.
     */
    fun getFontHeight(): Int {
        return fontHeight
    }

    /**
     * @param fontHeight The fontHeight to set.
     */
    fun setFontHeight(fontHeight: Int) {
        this.fontHeight = fontHeight
        bHasFontBlock = true
        if (font != null) font!!.fontHeight = fontHeight
    }

    /**
     * @return Returns the fontWeight.
     */
    fun getFontWeight(): Int {
        return fontWeight
    }

    /**
     * @param fontWeight The fontWeight to set.
     */
    fun setFontWeight(f: Int) {
        this.fontWeight = f
        this.fontOptsFlag = this.fontOptsFlag and 0xFD    // turn off bit 1 = style bit
        bHasFontBlock = true
        if (font != null) font!!.fontWeight = f
    }

    /**
     * @return Returns the fontUnderlineStyle.
     */
    fun getFontUnderlineStyle(): Int {
        return fontUnderlineStyle
    }

    /**
     * @param fontUnderlineStyle The fontUnderlineStyle to set.
     */
    fun setFontUnderlineStyle(fontUnderlineStyle: Int) {
        this.fontUnderlineStyle = fontUnderlineStyle
        this.fontUnderlineModifiedFlag = 0    // set modified flag
        bHasFontBlock = true
        if (font != null) font!!.setUnderlineStyle(fontUnderlineStyle.toByte())
    }

    /**
     * @return Returns the fontColorIndex.
     */
    fun getFontColorIndex(): Int {
        return fontColorIndex
    }

    /**
     * @param fontColorIndex The fontColorIndex to set.
     */
    fun setFontColorIndex(fontColorIndex: Int) {
        this.fontColorIndex = fontColorIndex
        bHasFontBlock = true
        if (font != null) font!!.color = fontColorIndex
    }

    /**
     * Reset the ptgRef to A1 in these that is replaced with
     * current Ptg
     */
    private fun resetFormulaRef() {
        // TODO: test what happens when A1 is a valid part of the expression
        val expr = this.getFormula1()!!.expression
        val itx = expr!!.iterator()
        if (refPos > -1) {
            expr.insertElementAt(prefHolder, refPos)
        }
    }

    /**
     * There is a ptgRef to A1 in these that is replaced with
     * current Ptg
     */
    @Throws(FormulaNotFoundException::class)
    private fun setFormulaRef(refcell: Ptg) {
        // TODO: test what happens when A1 is a valid part of the expression
        val expr = this.getFormula1()!!.expression
        val itx = expr!!.iterator()

        val rc = refcell.intLocation
        if (refPos == -1) {
            while (itx.hasNext()) {
                val prex = itx.next() as Ptg
                if (prex is PtgRefN) {
                    prex.setFormulaRow(rc[0])
                    prex.setFormulaCol(rc[1])
                }
            }
        } else {
            expr.removeAt(refPos)
        }
        if (refPos > -1) {
            expr.remove(prefHolder)
            expr.insertElementAt(refcell, refPos)
        }

    }

    /**
     * pass in the referenced cell and attempt to
     * create a valid formula for this thing and
     * calculate whether the criteria passes
     *
     *
     * returns true or false
     *
     * @param the reference to evaluate
     * @return boolean passes
     */
    fun evaluate(refcell: Ptg): Boolean {
        try {
            var val2: Any? = null
            var val1: Any? = null

            if (this.operator.toInt() != 0x0)
            // calcs later
                val1 = this.getFormula1()!!.calculateFormula()

            if (this.cce2 > 0)
                val2 = this.getFormula2()!!.calculateFormula()

            val valX = refcell.value

            // cast to double then compare
            var d1 = 0.0
            var d2 = 0.0 // second val from expression2

            var dX = 0.0 // the reference val

            try {
                d1 = Double(val1!!.toString())
                dX = Double(valX.toString())
                if (this.cce2 > 0) { // we have a second value
                    d2 = Double(val2!!.toString())
                }

            } catch (e: Exception) {
                // not numeric
            }

            // handle evaluated condition
            when (this.operator) {
                0x0    // No comparison (only valid for formula type, see above)
                -> {
                    setFormulaRef(refcell)
                    val1 = this.getFormula1()!!.calculateFormula()
                    return (val1 as Boolean).booleanValue()
                }

                1    // Between
                ->
                    // expression2 for the other bounds ...
                    return dX >= d1 && dX <= d2

                0x5    // Greater than
                -> return dX > d1

                0x2    // Not between
                -> {
                    if (dX < d1)
                        return false
                    else if (dX > d1)
                    // hmmm... d2 is where? an array?
                        return false
                    return dX < d1
                }

                0x6    // Less than
                -> return dX < d1

                0x3    // Equal
                -> return dX == d1

                0x7    // Greater or equal
                -> return dX >= d1

                0x4    // Not equal
                -> return dX != d1

                0x8    // Less or equal
                -> return dX <= d1

                // 2007-specific operators
                0x9    // begins With
                    , 0xA    // ends With
                    , 0xB    // contains text
                    , 0xC    // not contains
                -> return false

                else -> return false
            }
        } catch (ex: Exception) {
            // Logger.logWarn("CF condition "+this.formula1.getFormulaString()+" evaluation failed for : " + refcell.toString());
            return false
        }

    }


    /**
     * return the first formula referenced by the Conditional Format
     *
     * @return Formula
     */
    fun getFormula1(): Formula? {
        if (formula1 == null) { // hasn't been set
            formula1 = Formula()
            formula1!!.workBook = this.workBook
            if (this.sheet == null)
                this.setSheet(this.condfmt!!.sheet) // help!
            formula1!!.setSheet(this.sheet)
            formula1!!.expression = expression1
        }
        // 20101216 KSC: WHY????    	formula1.setCachedValue(null);
        return formula1
    }

    /**
     * return the second formula referenced by the Conditional Format
     *
     * @return Formula
     */
    fun getFormula2(): Formula? {
        if (formula2 == null && cce2 > 0) { // hasn't been set
            formula2 = Formula()
            formula2!!.workBook = this.workBook
            if (this.sheet == null)
                this.setSheet(this.condfmt!!.sheet) // help!
            formula2!!.setSheet(this.sheet)
            formula2!!.expression = expression2
        }
        if (formula2 != null) {
            formula2!!.setSheet(this.sheet)
            // 20101216 KSC: WHY???			formula2.setCachedValue(null);
        }
        return formula2
    }

    /**
     * OOXML-specific, in a containsText-type condition,
     * this field is the comparitor
     *
     * @param s
     */
    fun setContainsText(s: String) {
        containsText = s
    }

    /**
     * generate OOXML for this Cf (BIFF8->OOXML)
     * attributes:   type, dxfId, priority (REQ), stopIfTrue, aboveAverage,
     * percent, bottom, operator, text, timePeriod, rank, stdDev, equalAverage
     * children:		SEQ: formula (0-3), colorScale, dataBar, iconSet
     *
     * @return
     */
    fun getOOXML(bk: WorkBookHandle, priority: Int, dxfs: ArrayList<*>): String {
        val ooxml = StringBuffer()

        // first deal with dfx's (differential xf's) - part of styles.xml; here we need to add dxf element to dxf's plus trap dxfId
        val dxf = Dxf()
        if (this.bHasFontBlock) {
            if (font != null)
                dxf.font = font
            else
                dxf.createFont(this.fontWeight, this.fontItalic, this.fontUnderlineStyle, this.fontColorIndex, this.fontHeight)
        }
        if (this.bHasPatternBlock) {
            if (fill != null)
                dxf.fill = fill!!
            else
                dxf.createFill(this.patternFillStyle, this.patternFillColor, this.patternFillColorBack, bk)
        }
        if (this.bHasBorderBlock) {
            dxf.createBorder(bk, this.borderStyles, intArrayOf(this.getBorderLineColorTop(), this.getBorderLineColorLeft(), this.getBorderLineColorBottom(), this.getBorderLineColorRight()))
        }
        // TODO: check if this dxf already exists ****************************************************************************
        dxfs.add(dxf)    // save newly created dxf (differential xf) to workbook store
        val dxfId = dxfs.size - 1    // link this cf to it's dxf  NOTE: one of the ONLY OOXML id's that is 0-based ...

        ooxml.append("<cfRule dxfId=\"$dxfId\"")
        /*
         * attributes:   type, dxfId, priority (REQ), stopIfTrue, aboveAverage,
         * 				percent, bottom, operator, text, timePeriod, rank, stdDev, equalAverage
         * children:		SEQ: formula (0-3), colorScale, dataBar, iconSet
         */
        // TODO: ct==0 translates to ??
        // NOTE: types 3 and above are 2007 version (OOXML)-specific
        when (this.type) {
            1 -> ooxml.append(" type=\"cellIs\"")
            2 -> ooxml.append(" type=\"expression\"")
            3 -> ooxml.append(" type=\"containsText\"")
            4 -> ooxml.append(" type=\"aboveAverage\"")
            5 -> ooxml.append(" type=\"beginsWith\"")
            6 -> ooxml.append(" type=\"colorScale\"")
            7 -> ooxml.append(" type=\"containsBlanks\"")
            8 -> ooxml.append(" type=\"containsErrors\"")
            9 -> ooxml.append(" type=\"dataBar\"")
            10 -> ooxml.append(" type=\"duplicateValues\"")
            11 -> ooxml.append(" type=\"endsWith\"")
            12 -> ooxml.append(" type=\"iconSet\"")
            13 -> ooxml.append(" type=\"notContainsBlanks\"")
            14 -> ooxml.append(" type=\"notContainsErrors\"")
            15 -> ooxml.append(" type=\"notContainsText\"")
            16 -> ooxml.append(" type=\"timePeriod\"")
            17 -> ooxml.append(" type=\"top10\"")
            18 -> ooxml.append(" type=\"uniqueValues\"")
        }
        if (this.type.toInt() == 3 && containsText != null)
        // containsText	- shouldn't be null!
            ooxml.append(" text=\"$containsText\"")

        // operator
        when (this.operator) {
            1    // Between
            -> ooxml.append(" operator=\"between\"")
            0x5    // Greater than
            -> ooxml.append(" operator=\"greaterThan\"")
            0x2    // Not between
            -> ooxml.append(" operator=\"notBetween\"")
            0x6    // Less than
            -> ooxml.append(" operator=\"lessThan\"")
            0x3    // Equal
            -> ooxml.append(" operator=\"equal\"")
            0x7    // Greater or equal
            -> ooxml.append(" operator=\"greaterThanOrEqual\"")
            0x4    // Not equal
            -> ooxml.append(" operator=\"notEqual\"")
            0x8    // Less or equal
            -> ooxml.append(" operator=\"lessThanOrEqual\"")
            // 2007-specific
            0x9    // begins With
            -> ooxml.append(" operator=\"beginsWith\"")
            0xA    // ends With
            -> ooxml.append(" operator=\"endsWith\"")
            0xB    // begins With
            -> ooxml.append(" operator=\"containsText\"")
            0xC    // begins With
            -> ooxml.append(" operator=\"notContains\"")
        }
        // priority
        ooxml.append(" priority=\"$priority\"")
        // stopIfTrue == looks like this is set by default
        ooxml.append(" stopIfTrue=\"1\"")
        ooxml.append(">")
        if (this.getFormula1() != null)
            ooxml.append("<formula>" + OOXMLAdapter.stripNonAsciiRetainQuote(this.getFormula1()!!.formulaString).substring(1) + "</formula>")
        if (this.getFormula2() != null)
            ooxml.append("<formula>" + OOXMLAdapter.stripNonAsciiRetainQuote(this.getFormula2()!!.formulaString).substring(1) + "</formula>")

        // TODO: finish children dataBar, colorScale, iconSet, aboveAverage, bottom, equalAverage,
        ooxml.append("</cfRule>")
        return ooxml.toString()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5624169378370505532L
        // Offset Size Contents
        // 0 2 Fill pattern style:
        // Bit Mask Contents
        // 15-10 FC00H Fill pattern style (only if patt - style = 0, ?3.11)
        val PATTERN_FILL_STYLE = 0xFC00

        // 2 2 Fill pattern colour indexes:
        // Bit Mask Contents
        // 6-0 007FH Colour index (?6.70) for pattern (only if patt - col = 0)
        // 13-7 3F80H Colour index (?6.70) for pattern background (only if patt - bgcol = 0)
        val PATTERN_FILL_COLOR = 0x007F
        val PATTERN_FILL_BACK_COLOR = 0x3F80

        // 68 4 Font options:
        // Bit Mask Contents
        // 1 00000002H Posture: 0 = Normal; 1 = Italic (only if font - style = 0)
        // 7 00000080H Cancellation: 0 = Off; 1 = On (only if font - canc = 0)
        val FONT_OPTIONS_POSTURE = 0x2
        val FONT_OPTIONS_CANCELLATION = 0x80
        val FONT_OPTIONS_POSTURE_NORMAL = 0
        val FONT_OPTIONS_POSTURE_ITALIC = 1
        val FONT_OPTIONS_CANCELLATION_OFF = 0
        val FONT_OPTIONS_CANCELLATION_ON = 1

        // Standard values are 0190H (400) for normal text and 02BCH (700) for bold text.

        // 74 2 Escapement type (only if font - esc = 0):
        // 0000H = None; 0001H = Superscript; 0002H = Subscript
        val FONT_ESCAPEMENT_NONE = 0x0
        val FONT_ESCAPEMENT_SUPER = 0x1
        val FONT_ESCAPEMENT_SUB = 0x2

        // 76 1 Underline type (only if font - underl = 0):
        val FONT_UNDERLINE_NONE = 0x0
        val FONT_UNDERLINE_SINGLE = 0x1
        val FONT_UNDERLINE_DOUBLE = 0x2
        val FONT_UNDERLINE_SINGLEACCOUNTING = 0x21
        val FONT_UNDERLINE_DOUBLEACCOUNTING = 0x22

        val FONT_MODIFIED_OPTIONS_STYLE = 0x00000002
        val FONT_MODIFIED_OPTIONS_CANCELLATIONS = 0x00000080

        /**
         * Border Formatting Block
         */
        //	Offset Size Contents
        //	0 2 Border line styles:
        //	Bit Mask Contents
        //	3-0 000FH Left line style (only if bord - left = 0, ?3.10)
        //	7-4 00F0H Right line style (only if bord - right = 0, ?3.10)
        //	11-8 0F00H Top line style (only if bord - top = 0, ?3.10)
        //	15-12 F000H Bottom line style (only if bord - bot = 0, ?3.10)
        val BORDER_LINESTYLE_LEFT = 0x000F
        val BORDER_LINESTYLE_RIGHT = 0x00F0
        val BORDER_LINESTYLE_TOP = 0x0F00
        val BORDER_LINESTYLE_BOTTOM = 0xF000
        // if flags & BORDER_MODIFIED_XX == 0 means that this particular border has been modified
        val BORDER_MODIFIED_LEFT = 0x0400
        val BORDER_MODIFIED_RIGHT = 0x0800
        val BORDER_MODIFIED_TOP = 0x1000
        val BORDER_MODIFIED_BOTTOM = 0x2000

        // 2 4 Border line colour indexes:
        // 	Bit Mask Contents
        //	6-0 0000007FH Colour index (?6.70) for left line (only if bord - left = 0)
        //	13-7 00003F80H Colour index (?6.70) for right line (only if bord - right = 0)
        //	22-16 007F0000H Colour index (?6.70) for top line (only if bord - top = 0)
        //	29-23 3F800000H Colour index (?6.70) for bottom line (only if bord - bot = 0)
        val BORDER_LINECOLOR_LEFT = 0x0000007F
        val BORDER_LINECOLOR_RIGHT = 0x00003F80
        val BORDER_LINECOLOR_TOP = 0x007F0000
        val BORDER_LINECOLOR_BOTTOM = 0x3F800000


        /**
         * helper method to set the style values on this object from a style string
         * <br></br>each name/value pair is delimited by ;
         * <br></br>possible tokens:
         * <br></br>pattern		pattern fill #
         * <br></br>color     	pattern fg color
         * <br></br>patterncolor     pattern bg color
         * <br></br>vertical	vertical alignnment
         * <br></br>horizontal	horizontal alignment
         * <br></br>border		sub-tokens: border-top, border-left, border-bottom, border-top
         *
         * @param style
         * @param cx
         */
        fun setStylePropsFromString(style: String, cx: Cf) {
            val toks = StringTool.getTokensUsingDelim(style, ";")
            // iterate styles, set values
            for (t in toks.indices) {
                val pos = toks[t].indexOf(":")
                var n = ""
                // TODO: border line sizes, interpret border line styles
                // TODO: handle vertical, horizontal alignment, number format ...
                if (pos > 0)
                    n = toks[t].substring(0, pos)

                var v = toks[t].substring(toks[t].indexOf(":") + 1)
                n = StringTool.strip(n, '"')
                v = StringTool.strip(v, '#')
                n = StringTool.allTrim(n)
                v = StringTool.allTrim(v)
                n = n.toLowerCase()
                if (n.indexOf("border") == 0) { // parse the border settings
                    val vs = StringTool.getTokensUsingDelim(v, " ")
                    var sz = -1
                    var cz = -1
                    var stl = -1
                    var clr: String? = null
                    try {
                        sz = Integer.parseInt(vs[0])    // size
                        clr = vs[2]                    // color String
                        // TODO: no way to set border size as of yet

                        // interpret style string into #:
                        for (i in FormatConstants.BORDER_NAMES.indices) {
                            if (FormatConstants.BORDER_NAMES[i] == vs[1]) {
                                stl = i
                                break
                            }
                        }

                        // HexStringToColorInt
                        // java.awt.Color c = FormatHandle.HexStringToColor(clr);
                        cz = FormatHandle.HexStringToColorInt(clr!!, FormatHandle.colorBACKGROUND)
                        //	Logger.logInfo("ix " + sz + " " + clr + " " + stl );

                    } catch (ex: Exception) {
                    }

                    if (n.indexOf("border-top") == 0) {
                        if (clr != null) { // set color
                            cx.setBorderLineColorTop(cz)
                        }
                        if (stl > -1) {
                            cx.setBorderLineStylesTop(stl)
                        }
                    } else if (n.indexOf("border-left") == 0) {
                        if (clr != null) { // set color
                            cx.setBorderLineColorLeft(cz)
                        }
                        if (stl > -1) {
                            cx.setBorderLineStylesLeft(stl)
                        }
                    } else if (n.indexOf("border-bottom") == 0) {
                        if (clr != null) { // set color
                            cx.setBorderLineColorBottom(cz)
                        }
                        if (stl > -1) {
                            cx.setBorderLineStylesBottom(stl)
                        }
                    } else if (n.indexOf("border-right") == 0) {
                        if (clr != null) { // set color
                            cx.setBorderLineColorRight(cz)
                        }
                        if (stl > -1) {
                            cx.setBorderLineStylesRight(stl)
                        }
                    }

                } else if (n.equals("text-line-through", ignoreCase = true)) {
                    if (v.equals("none", ignoreCase = true)) {
                        cx.fontStriken = false
                    } else {
                        cx.fontStriken = true
                    }
                } else if (n.equals("fg", ignoreCase = true)) {    // font color
                    // set the color
                    val cl = FormatHandle.HexStringToColorInt(v, FormatHandle.colorFOREGROUND)
                    cx.setFontColorIndex(cl)
                } else if (n.equals("pattern", ignoreCase = true)) {    //fill pattern
                    val vv = Integer.parseInt(v)
                    cx.setPatternFillStyle(vv)
                } else if (n.equals("color", ignoreCase = true) || n.equals("patterncolor", ignoreCase = true)) {    // fill (pattern or fg) color
                    val cl = FormatHandle.HexStringToColorInt(v, FormatHandle.colorFONT)    // finds best match
                    cx.setPatternFillColor(cl, v)
                } else if (n.equals("background", ignoreCase = true)) {        // fill bg color
                    val cl = FormatHandle.HexStringToColorInt(v, FormatHandle.colorBACKGROUND)    // finds best match
                    cx.setPatternFillColorBack(cl)
                    // ALIGNMENT
                } else if (n.equals("alignment-horizontal", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    // TODO: set alignment
                } else if (n.equals("alignment-vertical", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    // TODO: set alignment
                } else if (n.equals("numberformat", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    // TODO: set number format
                    // FONT options
                } else if (n.equals("font-height", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    cx.setFontHeight(s)
                } else if (n.equals("font-weight", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    cx.setFontWeight(s)
                } else if (n.equals("font-EscapementFlag", ignoreCase = true)) {
                    val s = Integer.parseInt(v)    // 0= none, 1= superscript, 2= subscript
                    cx.fontEscapement = s
                } else if (n.equals("font-striken", ignoreCase = true)) {
                    val b = java.lang.Boolean.valueOf(v).booleanValue()
                    cx.fontStriken = b
                } else if (n.equals("font-italic", ignoreCase = true)) {
                    val b = java.lang.Boolean.valueOf(v).booleanValue()
                    cx.fontItalic = b
                } else if (n.equals("font-ColorIndex", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    cx.setFontColorIndex(s)
                } else if (n.equals("font-UnderlineStyle", ignoreCase = true)) {
                    val s = Integer.parseInt(v)
                    cx.setFontUnderlineStyle(s)
                }
                /*    			TODO: handle
 				fontName?
 * 				fontOptsFlag
    			fontModifiedOptionsFlag
    			fontUnderlineModified
*/
            }
            cx.updateRecord()
        }


        /**
         * set the conditional format format settings from the OOXML dxf, or differential xf, format settings
         *
         * @param dxf
         * @param cf
         */
        fun setStylePropsFromDxf(dxf: Dxf, cf: Cf) {
            val borderStyles = dxf.borderStyles
            if (borderStyles != null) {    // then should have every other element
                val borderColors = dxf.borderColors
                val borderSizes = dxf.borderSizes
                cf.setBorderLineColorTop(borderColors!![0])
                cf.setBorderLineStylesTop(borderStyles[0])
                cf.setBorderLineColorLeft(borderColors[1])
                cf.setBorderLineStylesLeft(borderStyles[1])
                cf.setBorderLineColorBottom(borderColors[2])
                cf.setBorderLineStylesBottom(borderStyles[2])
                cf.setBorderLineColorRight(borderColors[3])
                cf.setBorderLineStylesRight(borderStyles[3])
                // TODO border size??
            }

            // FILL ****************
            val fls = dxf.fillPatternInt
            if (fls >= 0) {
                cf.setPatternFillStyle(fls)
                val fg = dxf.fg
                val bg = dxf.bg
                if (fg != -1)
                    cf.setPatternFillColor(fg, null)
                if (bg != -1)
                    cf.setPatternFillColorBack(bg)
            }
            cf.fill = dxf.fill    // save OOXML var as color settings sometimes cannot be interpreted in 2003 v.

            var z: Int
            // ALIGNMENT ******************
            var s = dxf.horizontalAlign
            if (s != null) {
                z = Integer.parseInt(s)
                //TODO: set alignment
            }
            s = dxf.verticalAlign
            if (s != null) {
                z = Integer.parseInt(s)
                //TODO: set alignment
            }

            // Number Format ***************
            s = dxf.numberFormat
            if (s != null) {
                z = Integer.parseInt(s)
                //TODO: set number format
            }


            // FONT ************************
            if (dxf.font != null) {
                cf.parseFont(dxf.font!!)
                cf.font = dxf.font
            }


            /*    			TODO: handle
				fontName?
* 				fontOptsFlag
			fontModifiedOptionsFlag
			fontUnderlineModified
*/
            cf.updateRecord()
        }


        /**
         * Create a Cf record & populate with prototype bytes
         *
         * @return TODO: NOT FINISHED
         */
        val prototype: XLSRecord?
            get() {
                val cf = Cf()
                cf.opcode = XLSConstants.CF
                cf.setData(cf.PROTOTYPE_BYTES)
                cf.init()
                return cf
            }


        /**
         * Returns the byte Condition type from a human-readable string
         *
         * @return the condition type for this rule
         */
        fun getConditionFromString(cx: String): Byte {
            var cx = cx
            cx = StringTool.allTrim(cx) // trim
            cx = cx.toUpperCase()
            //    	 handle evaluated condition
            if (cx == "BETWEEN")
                return 0x1    // Between
            if (cx == "GREATER THAN")
                return 0x5    // Greater than
            if (cx == "NOT BETWEEN")
                return 0x2    // Not between
            if (cx == "LESS THAN")
                return 0x6    // Less than
            if (cx == "EQUALS")
                return 0x3    // Equal
            if (cx == "GREATER THAN OR EQUAL")
                return 0x7    // Greater or equal
            if (cx == "NOT EQUAL")
                return 0x4    // Not equal
            return if (cx == "LESS THAN OR EQUAL")
                0x8    // Less or equal
            else
                0x0    // No comparison (only valid for formula type, see above)
        }

        /**
         * restore Formula strings from XML serialization
         *
         * @param fmx
         * @return
         */
        private fun unescapeFormulaString(fmx: String): String {
            var fmx = fmx
            fmx = StringTool.replaceText(fmx, "&quot;", "\"")
            // fmx = StringTool.replaceText(fmx,"&amp;","%");
            fmx = StringTool.replaceText(fmx, "&lt;", "<")
            fmx = StringTool.replaceText(fmx, "&gt;", ">")
            return fmx
        }

        /**
         * taks a String representing the Operator for this Cf Rule
         * and translates it to an int
         * <br></br>NOTE: 2003 versions do not use types
         * 0x9, 0xA, 0xB or 0xC
         *
         * @param String operator - String CfRule operator attribute
         * @return int representing Cf operator value
         */
        fun translateOperator(operator: String?): Int {
            if (operator == null)
            // type is not cellIs
                return 0
            if (operator == "between")
                return 0x1
            else if (operator == "greaterThan")
                return 0x5
            else if (operator == "notBetween")
                return 0x2
            else if (operator == "lessThan")
                return 0x6
            else if (operator == "equal")
                return 0x3
            else if (operator == "greaterThanOrEqual")
                return 0x7
            else if (operator == "notEqual")
                return 0x4
            else if (operator == "lessThanOrEqual")
                return 0x8
            else if (operator == "beginsWith")
                return 0x9
            else if (operator == "endsWith")
                return 0xA
            else if (operator == "containsText")
                return 0xB
            else if (operator == "notContains")
                return 0xC// NO EQUIVALENT IN 2003:  	beginsWith, containsText, endsWith, notContains
            return 0
        }

        /**
         * Given an int type, return it's string representation
         * <br></br>NOTE: if type is between 3 and 18,
         * it is an OOXML-specific type.
         *
         * @param type type integer
         * @return String reprentation
         * @see translateType
         */
        protected fun translateType(type: Int): String {
            when (type) {
                1 -> return "Cell Is"
                2 -> return "expression"
                3 -> return "containsText"
                4 -> return "aboveAverage"
                5 -> return "beginsWith"
                6 -> return "colorScale"
                7 -> return "containsBlanks"
                8 -> return "containsErrors"
                9 -> return "dataBar"
                10 -> return "duplicateValues"
                11 -> return "endsWith"
                12 -> return "iconSet"
                13 -> return "notContainsBlanks"
                14 -> return "notContainsErrors"
                15 -> return "notContainsText"
                16 -> return "timePeriod"
                17 -> return "top10"
                18 -> return "uniqueValues"
                else -> return "Unknnown"
            }
        }

        /**
         * takes a String representing the type attribute and translates it to
         * the corresponding integer representation
         * <br></br>IMPORTANT NOTE: OOXML-specific types are converted to an integer that is not valid in 2003 versions
         *
         * @param String type - OOXML CfRule type attribute
         * @return int representing Cf type value
         */
        fun translateOOXMLType(type: String): Int {
            if (type == "cellIs")
                return 1
            else if (type == "expression")
                return 2
            else if (type == "containsText")
                return 3
            else if (type == "aboveAverage")
                return 4
            else if (type == "beginsWith")
                return 5
            else if (type == "colorScale")
                return 6
            else if (type == "containsBlanks")
                return 7
            else if (type == "containsErrors")
                return 8
            else if (type == "dataBar")
                return 9
            else if (type == "duplicateValues")
                return 10
            else if (type == "endsWith")
                return 11
            else if (type == "iconSet")
                return 12
            else if (type == "notContainsBlanks")
                return 13
            else if (type == "notContainsErrors")
                return 14
            else if (type == "notContainsText")
                return 15
            else if (type == "timePeriod")
                return 16
            else if (type == "top10")
                return 17
            else if (type == "uniqueValues")
                return 18// no equivalent in 2003: but must track for 2007 uses
            return 1    // default to cellIs ????
        }

        /**
         * prepare Formula strings for XML serialization
         *
         * @param fmx
         * @return
         */
        private fun escapeFormulaString(fmx: String): String {
            var fmx = fmx
            fmx = StringTool.replaceText(fmx, "\"", "&quot;")
            if (fmx.indexOf("=") == 0) {
                fmx = fmx.substring(1)
            }
            // fmx = StringTool.replaceText(fmx,"%","&amp;");
            fmx = StringTool.replaceText(fmx, "<", "&lt;")
            fmx = StringTool.replaceText(fmx, ">", "&gt;")
            return fmx

        }

        // helper inner method allows it to be static and reused...
        private fun getXML(cfx: Cf): String {

            // reset the placeholder formula reference
            if (cfx.refPos > -1)
                cfx.resetFormulaRef()

            val xml = StringBuffer()

            xml.append("<Range>")
            val rn = CellRange(cfx.condfmt!!.boundingRange)//getConditionalFormatRange();
            if (rn != null)
                xml.append(rn.r1C1Range)

            xml.append("</Range>")
            xml.append("<Condition>")
            if (cfx.operator.toInt() != 0x0) { // calcer
                xml.append("<Qualifier>")
                xml.append(cfx.conditionString)
                xml.append("</Qualifier>")
            }
            val val1: Any?
            val val2: Any?
            if (cfx.operator.toInt() != 0x0) { // calcer
                xml.append("<Value1>")
                val1 = cfx.getFormula1()!!.calculateFormula()
                if (val1 == null)
                // a new formula + calc_explicit
                {
                } else
                    xml.append(val1.toString())
                xml.append("</Value1>")
            } else {
                xml.append("<Value1>")
                var fmx = cfx.getFormula1()!!.formulaString
                fmx = Cf.escapeFormulaString(fmx)
                xml.append(fmx)
                xml.append("</Value1>")
            }

            if (cfx.cce2 > 0) {
                val2 = cfx.getFormula2()!!.calculateFormula()
                xml.append("<Value2>")
                if (val2 == null)
                // a new formula + calc_explicit
                {
                } else
                    xml.append(val2.toString())
                xml.append("</Value2>")
            }

            xml.append("<Format Style='")
            // attributes
            var cfi = cfx.getPatternFillColor()
            if (cfi > -1) {
                val fsi = FormatHandle.colorToHexString(FormatHandle.getColor(cfi))
                xml.append("color:$fsi;")
            }

            if (cfx.getFontWeight() > -1)
                xml.append("font-weight:" + cfx.getFontWeight() + ";")

            if (cfx.fontOptsCancellation > -1) {
                if (cfx.fontOptsCancellation == 0) {
                    xml.append("text-line-through:none;")
                } else {
                    xml.append("text-line-through:" + cfx.fontOptsCancellation + ";")
                }
            }
            if (cfx.borderSizes != null) {
                try {
                    xml.append("border-top:" + cfx.borderSizes!![0] + " " + FormatHandle.BORDER_NAMES[cfx.borderStyles!![0] + 1] + " " + FormatHandle.colorToHexString(cfx.borderColors!![0]) + ";") // .5pt solid windowtext;
                    xml.append("border-left:" + cfx.borderSizes!![1] + " " + FormatHandle.BORDER_NAMES[cfx.borderStyles!![1] + 1] + " " + FormatHandle.colorToHexString(cfx.borderColors!![1]) + ";") // .5pt solid windowtext;
                    xml.append("border-bottom:" + cfx.borderSizes!![2] + " " + FormatHandle.BORDER_NAMES[cfx.borderStyles!![2] + 1] + " " + FormatHandle.colorToHexString(cfx.borderColors!![2]) + ";") // .5pt solid windowtext;
                    xml.append("border-right:" + cfx.borderSizes!![3] + " " + FormatHandle.BORDER_NAMES[cfx.borderStyles!![3] + 1] + " " + FormatHandle.colorToHexString(cfx.borderColors!![3]) + ";") // .5pt solid windowtext;
                } catch (e: ArrayIndexOutOfBoundsException) {
                }

            }
            cfi = cfx.getPatternFillColorBack()
            if (cfi > -1) {
                val fsi = FormatHandle.colorToHexString(FormatHandle.getColor(cfi))
                xml.append("background:$fsi;")
            }

            xml.append("'/>")
            xml.append("</Condition>")

            return xml.toString()
        }

        /**
         * creates a Cf record from the EXML nodes
         *
         *
         * <Condition>
         * <Qualifier>Between</Qualifier>
         * <Value1>2</Value1>
         * <Value2>4</Value2>
         * <Format Style='color:#002060;font-weight:700;text-line-through:none;
        border:.5pt solid windowtext;background:#00B0F0'></Format>
        </Condition> *
         *
         * @param xpp
         * @return
         */
        fun parseXML(xpp: XmlPullParser, cfx: Condfmt, bs: Boundsheet): Cf? {
            val oe = bs.createCf(cfx)
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "Condition") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (n == "Qualifier") {
                                    oe.setOperator(v)
                                } else if (n == "Value1") {
                                    oe.setCondition1(v)
                                } else if (n == "Value2") {
                                    oe.setCondition2(v)
                                } else if (n == "Format") {
                                    setStylePropsFromString(v, oe)
                                }
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "Condition") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("Cf.parseXML: $e")
            }

            return oe
        }
    }

}
