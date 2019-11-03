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
package io.starter.OpenXLS

import io.starter.formats.OOXML.Fill
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.*
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONException
import org.json.JSONObject

import java.awt.Color
import java.util.*

/**
 * Provides methods for querying and changing cell formatting information. Cell
 * formating includes fonts, borders, text alignment, background colors, cell
 * protection (locking), etc.
 *
 *
 * The mutator methods of a FormatHandle object directly and immediately change
 * the formatting of the cell(s) to which it is applied. Under no circumstances
 * will a FormatHandle alter the formatting of cells to which it is not applied.
 * The list of cells to which a FormatHandle applies may be queried with the
 * [.getCells] method. Additional cells may be added with [.addCell]
 * and related methods.
 *
 *
 * A FormatHandle may also be obtained for a row or a column. In such a case the
 * FormatHandle sets the default formats for the row or column. The default
 * format for a row or column is the format which appears when no other formats
 * for the cell are specified, such as for newly created cells.
 */
class FormatHandle : Handle, FormatConstants {

    private var mycells = CompatibleVector() // all the Cells sharing this format e.g. CellRange
    /**
     * For internal usage only, return the internal XF record that
     * represents this FormatHandle
     *
     * @return the myxf
     */
    /**
     * Sets the internal format to the FormatHandle. For internal use only, not
     * supported.
     */
    var xf: Xf? = null
    /**
     * The format ID allows setting of the format for a Cell without adding it
     * to the Format's Cell Collection.
     *
     *
     * Use to decrease memory requirements when dealing with large collections
     * of cells.
     *
     *
     * Usage:
     *
     *
     * mycell.setFormatId(myformat.getFormatId());
     *
     * @return the format ID for this Format
     */
    /**
     * set the pointer to the XFE or Conditional format
     *
     * @param xfe
     */
    var formatId: Int = 0
    // see Xf.usedCount instead	private boolean canModify = false;
    private var mycol: ColHandle? = null
    private var myrow: RowHandle? = null
    private val writeImmediate = false
    internal var underlined = false

    // 20060412 KSC: added for access
    /**
     * Set the workbook for this FormatHandle
     *
     * @param bk
     */
    var workBook: io.starter.formats.XLS.WorkBook? = null
    private var wbh: io.starter.OpenXLS.WorkBook? = null

    /**
     * returns whether this FormatHandle is set to hide formula strings
     *
     * @return true if the formula strings are hidden, false otherwise
     */
    /**
     * sets the cell attached to this FormatHandle to hide or show formula
     * strings;
     *
     * @param boolean b- true if formulas should be hidden for this FormatHandle
     */
    var isFormulaHidden: Boolean
        get() = xf!!.isFormulaHidden
        set(b) {
            val xf = cloneXf(this.xf!!)
            xf.isFormulaHidden = b
            updateXf(xf)
        }

    /**
     * returns whether this Format Handle specifies that cells are locked for
     * changing
     *
     * @return true if cells are locked
     */
    /**
     * Locks the cell attached to this FormatHandle for editing (makes
     * read-only) lock cell and make read-only
     *
     * @param boolean locked - true if cells should be locked for this
     * FormatHandle
     */
    var isLocked: Boolean
        get() = xf!!.isLocked
        set(locked) {
            val xf = cloneXf(this.xf!!)
            xf.isLocked = locked
            updateXf(xf)
        }

    /**
     * provides a mapping between Excel formats and Java formats <br></br>
     * see: [http://java.sun.com/docs/books/tutorial/i18n/format
 * /decimalFormat.html](tutorial)
     *
     *
     * Note there are slight Excel-specific differences in the format strings
     * returned. Several numeric and currency formats in excel have different
     * formatting for postive and negative numbers. In these cases, the java
     * format string is split by semicolons and may contain text [Red] which is
     * to specify the negative number should be displayed in red. Remove this
     * from the string before passing into the Format class;
     *
     * <pre>
     * G  Era designator  Text  AD
     * y  Year  Year  1996; 96
     * M  Month in year  Month  July; Jul; 07
     * w  Week in year  Number  27
     * W  Week in month  Number  2
     * D  Day in year  Number  189
     * d  Day in month  Number  10
     * F  Day of week in month  Number  2
     * E  Day in week  Text  Tuesday; Tue
     * a  Am/pm marker  Text  PM
     * H  Hour in day (0-23)  Number  0
     * k  Hour in day (1-24)  Number  24
     * K  Hour in am/pm (0-11)  Number  0
     * h  Hour in am/pm (1-12)  Number  12
     * m  Minute in hour  Number  30
     * s  Second in minute  Number  55
     * S  Millisecond  Number  978
     * z  Time zone  General time zone  Pacific Standard Time; PST; GMT-08:00
     * Z  Time zone  RFC 822 time zone  -0800
    </pre> *
     *
     * @return String the formatting pattern for the cell
     */
    // toLowerCase is a simplistic way to implement the case insensitivity
    // of the pattern tokens. It could cause issues with string literals.
    /*
         * If we reached here, we don't have a mapping for this particular
         * format. Send a warning to the system then make sure the pattern we
         * are sending back is valid. Many excel patterns have 4 patterns
         * separated by semicolons. We only can pass 2 into the formatter
         * (positive and negative). This usually works out to be the first two
         * patterns in the string.
         */// yet another hackaround -jm
    val javaFormatString: String?
        get() {
            var pat = formatPattern ?: return null
            var patty: Any? = convertFormatString(pat!!.toLowerCase())
            if (patty != null)
                patty = StringTool.qualifyPatternString(patty.toString())
            if (xf!!.isDatePattern) {
                return if (patty != null)
                    patty as String?
                else
                    "M/d/yy h:mm"
            }
            if (patty != null)
                return patty as String?
            pat = StringTool.qualifyPatternString(pat)
            val firstParens = pat!!.indexOf(";")
            if (firstParens != -1) {
                val secondParens = pat!!.indexOf(";", firstParens + 1)
                if (secondParens != -1) {
                    pat = pat!!.substring(0, secondParens)
                } else {
                    pat = pat!!.substring(firstParens + 2, pat!!.length - 1)
                }
            }
            return pat
        }

    /**
     * Get the Right border color
     *
     * @return color constant
     */
    // black i'm afraid
    val borderRightColor: Color?
        get() {
            if (xf!!.rightBorderLineStyle.toInt() == 0)
                return null
            val x = xf!!.rightBorderColor.toInt()
            return if (x < this.workBook!!.colorTable.size)
                this.workBook!!.colorTable[x]
            else
                this.workBook!!.colorTable[0]
        }

    /**
     * returns Border Colors of Cell ie: top, left, bottom, right
     *
     *
     * returns null or 1 color for each of 4 sides
     *
     *
     * 1,1,1,1 represents a border all around the cell 1,1,0,0 represents on the
     * top left edge of the cell
     *
     * @return int array representing Cell borders
     */
    /*
     * FIXME: Border Issues (marker)
     *
     * The methods to set individual border line styles do not follow a
     * consistent naming convention with the rest of the border methods.
     *
     * There are no methods for getting or setting the inside borders. I don't
     * know whether this is a problem or a design choice. It depends on whether
     * the inside borders are stored in the format record or as separate cell
     * formats.
     *
     * There are no methods to get the diagonal borders. There is a method to
     * set the diagonal border line style, but it affects both diagonals.
     */

    /**
     * sets the border color for all borders (top, left, bottom and right) from
     * a Color array
     *
     *
     * NOTE: this setting will affect every cell which refers to this
     * FormatHandle
     *
     * @param java .awt.Color array - 4-element array of desired border colors
     * [T, L, B, R]
     */
    var borderColors: Array<Color>
        get() {
            val colors = arrayOfNulls<Color>(4)
            colors[0] = borderTopColor
            colors[1] = borderLeftColor
            colors[2] = borderBottomColor
            colors[3] = borderRightColor
            return colors
        }
        set(bordercolors) {
            val xf = cloneXf(this.xf!!)
            if (bordercolors[0] != null)
                xf.topBorderColor = getColorInt(bordercolors[0])
            if (bordercolors[1] != null)
                xf.leftBorderColor = getColorInt(bordercolors[1])
            if (bordercolors[2] != null)
                xf.bottomBorderColor = getColorInt(bordercolors[2])
            if (bordercolors[3] != null)
                xf.setRightBorderColor(getColorInt(bordercolors[3]))
            updateXf(xf)
        }

    /**
     * Get the Left border color
     *
     * @return color constant
     */
    val borderLeftColor: Color?
        get() = if (xf!!.leftBorderLineStyle.toInt() == 0) null else this.workBook!!.colorTable[xf!!.leftBorderColor]

    /**
     * Get the Top border color
     *
     * @return color constant
     */
    // guards
    val borderTopColor: Color?
        get() {
            if (xf!!.topBorderLineStyle.toInt() == 0)
                return null
            var xt = xf!!.topBorderColor
            if (xt > this.workBook!!.colorTable.size)
                xt = 0
            return this.workBook!!.colorTable[xt]
        }

    /**
     * Returns true if the value should be red due to a combination of a format
     * pattern and a negative number
     *
     * @return
     */
    val isRedWhenNegative: Boolean
        get() {
            val pattern = xf!!.formatPattern
            return pattern!!.indexOf("Red") > -1
        }

    /**
     * Get the Right border color
     *
     * @return color constant
     */
    // black i'm afraid
    val borderBottomColor: Color?
        get() {
            if (xf!!.bottomBorderLineStyle.toInt() == 0)
                return null
            val x = xf!!.bottomBorderColor
            return if (x < this.workBook!!.colorTable.size)
                this.workBook!!.colorTable[x]
            else
                this.workBook!!.colorTable[0]

        }

    /**
     * Get the border line style
     *
     * @return line style constant
     */
    /**
     * Set the border line style
     *
     * @param line style constant
     */
    var topBorderLineStyle: Int
        get() = xf!!.topBorderLineStyle.toInt()
        set(x) {
            val xf = cloneXf(this.xf!!)
            xf.topBorderLineStyle = x.toShort()
            updateXf(xf)
        }

    /**
     * Get the border line style
     *
     * @return line style constant
     */
    /**
     * Set the border line style
     *
     * @param line style constant
     */
    var bottomBorderLineStyle: Int
        get() = xf!!.bottomBorderLineStyle.toInt()
        set(x) {
            val xf = cloneXf(this.xf!!)
            xf.bottomBorderLineStyle = x.toShort()
            updateXf(xf)
        }

    /**
     * Get the border line style
     *
     * @return line style constant
     */
    /**
     * Set the border line style
     *
     * @param line style constant
     */
    var leftBorderLineStyle: Int
        get() = xf!!.leftBorderLineStyle.toInt()
        set(x) {
            val xf = cloneXf(this.xf!!)
            xf.leftBorderLineStyle = x.toShort()
            updateXf(xf)
        }

    /**
     * Get the border line style
     *
     * @return line style constant
     */
    /**
     * Set the border line style
     *
     * @param line style constant
     */
    var rightBorderLineStyle: Int
        get() = xf!!.rightBorderLineStyle.toInt()
        set(x) {
            val xf = cloneXf(this.xf!!)
            xf.rightBorderLineStyle = x.toShort()
            updateXf(xf)
        }

    /**
     * return the 5 border lines styles (l, r, t, b, diag)
     *
     * @return
     */
    val allBorderLineStyles: IntArray
        get() {
            val ret = IntArray(5)
            ret[0] = xf!!.leftBorderLineStyle.toInt()
            ret[1] = xf!!.rightBorderLineStyle.toInt()
            ret[2] = xf!!.topBorderLineStyle.toInt()
            ret[3] = xf!!.bottomBorderLineStyle.toInt()
            ret[4] = xf!!.diagBorderLineStyle.toInt()
            return ret
        }

    /**
     * return the 5 border line colors (l, r, t, b, diag)
     *
     * @return
     */
    val allBorderColors: IntArray
        get() {
            val ret = IntArray(5)
            ret[0] = xf!!.leftBorderColor
            ret[1] = xf!!.rightBorderColor.toInt()
            ret[2] = xf!!.topBorderColor
            ret[3] = xf!!.bottomBorderColor
            ret[4] = xf!!.diagBorderColor.toInt()
            return ret
        }

/**
 * Get the font weight the weight of the font is in 1/20 point units
 *
 * @return
 */
/**
 * /** Set the weight of the font in 1/20 point units 100-1000 range. 400 is
 * normal, 700 is bold.
 *
 * @param wt
*/
var fontWeight:Int
get() {
return xf!!.font!!.fontWeight
}
set(wt) {
val f = cloneFont(xf!!.font!!)
f.fontWeight = wt
updateFont(f)
}

/**
 * returns whether this Format is formatted as a Date
 *
 * @return boolean true if this Format is formatted as a Date
*/
val isDate:Boolean
get() {
if (xf == null)
return false
return xf!!.isDatePattern
}

/**
 * returns whether this Format is formatted as a Currency
 *
 * @return boolean true if this Format is formatted as a currency
*/
val isCurrency:Boolean
get() {
if (xf == null)
return false
return xf!!.isCurrencyPattern
}

/**
 * Get the Font foreground (text) color as a java.awt.Color
 *
 * @return
*/
val fontColor:Color
get() {
return xf!!.font!!.colorAsColor
}


val fontColorAsHex:String
get() {
return xf!!.font!!.colorAsHex
}

/**
 * Get the Pattern Background Color for this Format Pattern
 *
 * @return the Excel color constant
*/
/**
 * set the background color for this Format NOTE: Foreground color = the
 * CELL BACKGROUND color color for all patterns and Background color= the
 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
 * color=CELL BACKGROUND color and Background Color=64 (white).
 *
 * @param t Excel color constant
*/
var backgroundColor:Int
get() {
return xf!!.backgroundColor.toInt()
}
set(t) {
val xf = cloneXf(this.xf!!)
if (xf.fillPattern == Xf.PATTERN_SOLID)
xf.setForeColor(t, null)
else
xf.setBackColor(t, null)

updateXf(xf)
}

/**
 * get the Pattern Background Color for this Format Pattern as a hex string
 *
 * @return Hex Color String
*/
val backgroundColorAsHex:String?
get() {
return xf!!.backgroundColorHEX
}

/**
 * get the Pattern Background Color for this Format Pattern as an awt.Color
 *
 * @return background Color
*/
val backgroundColorAsColor:java.awt.Color
get() {
return HexStringToColor(this.backgroundColorAsHex!!)
}


/**
 * returns the foreground color setting regardless of format pattern (which
 * can switch fg and bg)
 *
 * @return
*/
// 20080814 KSC: getForegroundColor() does the swapping so use base method
val trueForegroundColor:Int
get() {
return xf!!.foregroundColor.toInt()
}

/**
 * get the Pattern Background Color for this Formatted Cell
 *
 *
 * This method handles display of conditional formats for the cell
 *
 *
 * checks for conditional format, then applies it if conditions are true.
 *
 * @return the Excel color constant
*/
// this.getForegroundColor() does
// the swapping so use base
// method
val cellBackgroundColor:Int
get() {
val fp = fillPattern

if (fp == Xf.PATTERN_SOLID)
return xf!!.foregroundColor.toInt()
else
return xf!!.backgroundColor.toInt()
}

/**
 * get the Pattern Background Color for this Format as a Hex Color String
 *
 * @return Hex Color String
*/
// this.getForegroundColor() does
val cellBackgroundColorAsHex:String?
get() {
val fp = fillPattern

if (fp == Xf.PATTERN_SOLID)
return xf!!.foregroundColorHEX
else
return xf!!.backgroundColorHEX
}

/**
 * get the Pattern Background Color for this Format as an awt.Color
 *
 * @return cell background color
*/
val cellBackgroundColorAsColor:java.awt.Color
get() {
return HexStringToColor(this.cellBackgroundColorAsHex!!)
}

/**
 * get the Background Color for this Format NOTE: Foreground color = the
 * CELL BACKGROUND color color for all patterns and Background color= the
 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
 * color=CELL BACKGROUND color and Background Color=64 (white).
 *
 * @return the Excel color constant
*/
/**
 * set the Foreground Color for this Format NOTE: Foreground color = the
 * CELL BACKGROUND color color for all patterns and Background color= the
 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
 * color=CELL BACKGROUND color and Background Color=64 (white).
 *
 * @param int Excel color constant
*/
// if it's SOLID pattern, fg/bg are swapped
var foregroundColor:Int
get() {
if (fillPattern == Xf.PATTERN_SOLID)
return xf!!.backgroundColor.toInt()

return xf!!.foregroundColor.toInt()
}
set(t) {
val xf = cloneXf(this.xf!!)
xf.setForeColor(t, null)
updateXf(xf)
}

/**
 * get the Background Color for this Format as a Hex Color String NOTE:
 * Foreground color = the CELL BACKGROUND color color for all patterns and
 * Background color= the PATTERN color for all patterns != Solid For
 * PATTERN_SOLID, Foreground color=CELL BACKGROUND color and Background
 * Color=64 (white).
 *
 * @return Hex Color String
*/
// if it's SOLID pattern,
val foregroundColorAsHex:String?
get() {
if (fillPattern == Xf.PATTERN_SOLID)
return xf!!.backgroundColorHEX

return xf!!.foregroundColorHEX
}

/**
 * get the Background Color for this Format as a Color NOTE:
 * Foreground color = the CELL BACKGROUND color color for all patterns and
 * Background color= the PATTERN color for all patterns != Solid For
 * PATTERN_SOLID, Foreground color=CELL BACKGROUND color and Background
 * Color=64 (white).
 *
 * @return Hex Color String
*/
val foregroundColorAsColor:Color
get() {
return HexStringToColor(this.foregroundColorAsHex!!)
}

/**
 * Get if this format is bold or not
 *
 * @return boolean whether the cell font format is bold
*/
val isBold:Boolean
get() {
return xf!!.font!!.isBold
}

/**
 * Return an int representing the underline style
 *
 *
 * These map to the STYLE_UNDERLINE static integers *
 *
 * @return int underline style
*/
/**
 * set the underline style for this font
 *
 * @param int u underline style one of the Font.STYLE_UNDERLINE constants
*/
var underlineStyle:Int
get() {
return xf!!.font!!.underlineStyle
}
set(u) {
val f = cloneFont(xf!!.font!!)
f.setUnderlineStyle(u.toByte())
updateFont(f)
}

/**
 * Get the font height in points
 *
 * @return font height
*/
val fontHeightInPoints:Double
get() {
return font!!.fontHeightInPoints
}

/**
 * Returns the Font's height in 1/20th point increment
 *
 * @return font height
*/
/**
 * Set the Font's height in 1/20th point increment
 *
 * @param new font height
*/
var fontHeight:Int
get() {
return font!!.fontHeight
}
set(fontHeight) {
val f = cloneFont(xf!!.font!!)
f.fontHeight = fontHeight
updateFont(f)
}

/**
 * Returns the Font's name
 *
 * @return font name
*/
/**
 * Set the Font's name
 *
 *
 * To be valid, this font name must be available on the client system.
 *
 * @param font name
*/
var fontName:String?
get() {
return font!!.getFontName()
}
set(fontName) {
val f = cloneFont(xf!!.font!!)
f.setFontName(fontName)
updateFont(f)
}

/**
 * Determine if the format handle refers to a font stricken out
 *
 * @return boolean representing if the FormatHandle is striking out a cell.
*/
/**
 * Set if the format handle is stricken out
 *
 * @param isStricken boolean representing if the formatted cell should be stricken
 * out.
*/
var stricken:Boolean
get() {
if (xf!!.font == null)
return false
return xf!!.font!!.stricken
}
set(isStricken) {
val f = cloneFont(xf!!.font!!)
f.stricken = isStricken
updateFont(f)
}

/**
 * Get if the font is italic
 *
 * @return boolean representing if the formatted cell is italic.
*/
/**
 * Set if the font is italic
 *
 * @param isItalic boolean representing if the formatted cell should be italic.
*/
var italic:Boolean
get() {
if (xf!!.font == null)
return false
return xf!!.font!!.italic
}
set(isItalic) {
val f = cloneFont(xf!!.font!!)
f.italic = isItalic
updateFont(f)
}

/**
 * Get if the font is bold
 *
 * @return boolean representing if the formatted cell is bold
*/
/**
 * Set the format handle to use standard bold text
 *
 * @param boolean isBold
*/
var bold:Boolean
get() {
return xf!!.font!!.bold
}
set(isBold) {
val f = cloneFont(xf!!.font!!)
f.bold = isBold
updateFont(f)
}

/**
 * returns the existing font record for this Format
 *
 *
 * Font is an internal record and should not be accessed by end users
 *
 * @return the XLS Font record associated with this Format
*/
/**
 * Set the Font for this Format.
 *
 *
 * As adding a new Font and format increases the file size, try using this
 * once for each distinct font used in the file, then use 'setFormatId' to
 * share this font with other Cells.
 *
 *
 * Roughly matches the functionality of the java.awt.Font class. Currently
 * the style parameter is only useful for bold/normal weights. Italics,
 * underlines, etc must be modified elsewhere.
 *
 *
 * Note that in order to maintain java.awt.Font compatibility for
 * bold/normal styles, defaults for weight/style have been mapped to 0 =
 * normal (excel 200 weight) and 1 = bold (excel 700 weight)
 *
 * @param String font name
 * @param int    font style
 * @param int    font size
*/
// should this be a protected method?
// shouldn't!
var font:io.starter.formats.XLS.Font?
get() {
if (xf != null)
return xf!!.font
return null
}
set(f) {
setXFToFont(f)
}

/**
 * Gets the number format pattern for this format, if set. For more
 * information on number format patterns see [Microsoft KB264372](http://support.microsoft.com/kb/264372).
 *
 * @return the Excel number format pattern for this cell or
 * `null` if none is applied
*/
/**
 * Sets the number format pattern for this format. All Excel built-in number
 * formats are supported. Custom formats will not be applied by OpenXLS
 * (e.g. [CellHandle.getFormattedStringVal]) but they will be written
 * correctly to the output file. For more information on number format
 * patterns see [Microsoft
 * KB264372](http://support.microsoft.com/kb/264372).
 *
 * @param pat the Excel number format pattern to apply
 * @see FormatConstantsImpl.getBuiltinFormats
*/
var formatPattern:String?
get() {
return xf!!.formatPattern
}
set(pat) {
val xf = cloneXf(this.xf!!)
xf.formatPattern = pat
updateXf(xf)
}

/**
 * get the fill pattern for this format
*/
val fillPattern:Int
get() {
return xf!!.fillPattern
}

/**
 * Returns an int representing the current horizontal alignment in this
 * FormatHandle. These values are mapped to the FormatHandle.ALIGN*** static
 * int's
*/
/**
 * Set the horizontal alignment for this FormatHandle
 *
 * @param align - an int representing the alignment. Please review the
 * FormatHandle.ALIGN*** static int's
*/
var horizontalAlignment:Int
get() {
return xf!!.horizontalAlignment
}
set(align) {
val xf = cloneXf(this.xf!!)
xf.horizontalAlignment = align
updateXf(xf)
}

/**
 * return indent (1 = 3 spaces)
 *
 * @return
*/
/**
 * set indent (1= 3 spaces)
 *
 * @param indent
*/
var indent:Int
get() {
return xf!!.indent
}
set(indent) {
val xf = cloneXf(this.xf!!)
xf.indent = indent
updateXf(xf)
}

/**
 * returns true if this style is set to Right-to-Left text direction
 * (reading order)
 *
 * @return
*/
val rightToLetReadingOrder:Int
get() {
return xf!!.rightToLeftReadingOrder
}

/**
 * Returns an int representing the current Vertical alignment in this
 * FormatHandle. These values are mapped to the FormatHandle.ALIGN*** static
 * int's
*/
/**
 * Set the Vertical alignment for this FormatHandle
 *
 * @param align - an int representing the alignment. Please review the
 * FormatHandle.ALIGN*** static int's
*/
var verticalAlignment:Int
get() {
return xf!!.verticalAlignment
}
set(align) {
val xf = cloneXf(this.xf!!)
xf.verticalAlignment = align
updateXf(xf)
}

/**
 * Get the cell wrapping behavior for this FormatHandle. Default is false
*/
/**
 * Set the cell wrapping behavior for this FormatHandle. Default is false
*/
var wrapText:Boolean
get() {
return xf!!.wrapText
}
set(wrapit) {
val xf = cloneXf(this.xf!!)
xf.wrapText = wrapit
updateXf(xf)
}

/**
 * Get the rotation of the cell. Value 0 means no rotation (horizontal
 * text). Values 1-90 mean rotation up (couter-clockwise) by 1-90 degrees.
 * Values 91-180 mean rotation down (clockwise) by 1-90 degrees. Value 255
 * means vertical text.
*/
/**
 * Set the rotation of the cell in degrees. Values 0-90 represent rotation
 * up, 0-90degrees. Values 91-180 represent rotation down, 0-90 degrees.
 * Value 255 is vertical
 *
 * @param align - an int representing the rotation.
*/
var cellRotation:Int
get() {
return xf!!.rotation
}
set(align) {
val xf = cloneXf(this.xf!!)
xf.rotation = align
updateXf(xf)
}

/**
 * Gets the Excel format ID for this format's number format pattern.
 *
 * @return the Excel format identifier number for the number format pattern
*/
/**
 * Sets the number format pattern based on the format ID number. This method
 * is recommended for advanced users only. In most cases you should use
 * [.setFormatPattern] instead.
 *
 * @param fmt the format ID number for the desired number format pattern
*/
var formatPatternId:Int
get() {
return xf!!.ifmt.toInt()
}
set(fmt) {
xf!!.setFormat(fmt.toShort())
}

/**
 * return truth of "this Xf rec is a style xf"
 *
 * @return
*/
val isStyleXf:Boolean
get() {
return xf!!.isStyleXf
}

/**
 * Nullary constructor for use in bean context. **This is not part of the
 * public API and should not be called.**
*/
constructor() {}

/**
 * Constructs a FormatHandle for the given WorkBook and format record.
 * **This is not part of the public API and should not be called.**
*/
constructor(book:io.starter.OpenXLS.WorkBook?, xfr:Xf) {
xf = xfr
formatId = xf!!.idx
wbh = book
if (book != null)
{
workBook = book!!.workBook
}
else
{
if (xfr.workBook != null)
{
workBook = xfr.workBook
}
else
{
Logger.logErr("FormatHandle constructed with null WorkBook.")
}
}
}

/**
 * Constructs a FormatHandle for the given WorkBook's default format.
*/
constructor(book:io.starter.OpenXLS.WorkBook) : this(book, book.workBook.defaultIxfe) {}

/**
 * Constructs a FormatHandle for the given format ID and WorkBook.
 * This is useful for creating a FormatHandle with the same parameters as
 * a cell that does not refer to the cell. For example:
 * <pre>
 * CellHandle cell = &lt;get cell here&gt;;
 * FormatHandle format = new FormatHandle(
 * cell.getWorkBook(), cell.getFormatId() );
</pre> *
 *
 * @param book  the WorkBook from which the format should be retrieved
 * @param xfnum the ID of the format
*/
constructor(book:io.starter.formats.XLS.WorkBook, xfnum:Int) {
workBook = book
if (xfnum > -1 && xfnum < workBook!!.numXfs)
{
xf = workBook!!.getXf(xfnum)
formatId = xf!!.idx
}
if (xf == null)
{ // add new xf if necessary
xf = duplicateXf(null)
}
xf!!.font // will set to default (0th) font if not already set
}

/** Constructs a FormatHandle for the given format ID and WorkBook.
 * This is useful for creating a FormatHandle with the same parameters as
 * a cell that does not refer to the cell. For example:
 * <pre>
 * CellHandle cell = &lt;get cell here&gt;;
 * FormatHandle format = new FormatHandle(
 * cell.getWorkBook(), cell.getFormatId() );
</pre> *
 *
 * @param book
 * the WorkBook from which the format should be retrieved
 * @param xfnum
 * the ID of the format public FormatHandle(WorkBook book, int
 * xfnum ,int x) { this(book.getWorkBook(),xfnum); wbh = book; }
*/

/**
 * Constructs a FormatHandle for the given format index and WorkBook.
 * **This is not part of the public API and should not be called.**
*/
/*
 * This constructor is just used from XML parsing, due to some dedupe
 * errors. possibly errors in .xlsx too?
*/
protected constructor(book:io.starter.OpenXLS.WorkBook, xfnum:Int,
dedupe:Boolean) : this(book, xfnum) {
writeImmediate = dedupe
}

/**
 * Constructs a dummy FormatHandle for the given conditional format. **This
 * is not part of the public API and should not be called.**
 *
 *
 * This unique flavor of FormatHandle is only used to display the formatting
 * values of a Conditional format record.
 *
 *
 * Creates a dummy Xf to store values, otherwise has no effect on the
 * WorkBook record stream which are manipulated through CellHandle.
 *
 * @param book containing the conditional formats
 * @param the  index to the conditional format in the book collection
*/
constructor(cx:Condfmt, book:io.starter.OpenXLS.WorkBook,
xfnum:Int, c:CellHandle?) {
var c = c
cx.formatHandle = this
formatId = xfnum
wbh = book
workBook = book.workBook
if (c == null)
{
// ok, this is a horrible hack, as its only correct if the top left
// cell of the range has the same background format
// as the cell a user is hitting. Lame, but i've been handed this at
// the last moment and am patching things. weak effort guys.
try
{
val rc = cx.encompassingRange
c = CellHandle(cx.sheet!!.getCell(rc[0], rc[1]),
book)
}
catch (e:Exception) {}

}
xf = duplicateXf(book.workBook.getXf(c!!.formatId))

// set the format from the cf
val lx = cx.rules
val itx = lx.iterator()
while (itx.hasNext())
{
val format = itx.next() as Cf
this.updateFromCF(format, book)
}
}

/**
 * updates this format handle via a Cf rule
 *
 * @param cf   Cf rule
 * @param book workbook
*/
fun updateFromCF(cf:Cf, book:io.starter.OpenXLS.WorkBook) {
// border colors
val clr = cf.borderColors
if (clr != null)
borderColors = clr

// line style
val xs = cf.borderStyles
if (xs != null)
this.setBorderLineStyle(xs)

/*
 * // cf.getBorderSizes() int[] b = cf.getBorderSizes(); if(b!=null)
 * setBorderLineStyle(b);
*/
// cf.getFont()
val f = cf.getFont()
if (f != null)
font = f
else
this.fontHeight = 180 // why????????

if (cf.fontItalic)
this.italic = true

if (cf.fontStriken)
this.stricken = true

val fsup = cf.fontEscapement
// super/sub (0 = none, 1 = super, 2 = sub)
if (fsup > -1)
this.setScript(fsup)

// handle underlines
val us = cf.getFontUnderlineStyle()

if (us > -1)
{
this.underlineStyle = us
this.setUnderlined(true)
}

// number cf
if (cf.formatPattern != null)
formatPattern = cf.formatPattern

if (cf.fill != null)
{
this.setFill(cf.fill)
}
else
{
val fill = cf.getPatternFillStyle() // Now -1 is a valid entry:
if (fill > -1)
{
/* If the fill style is solid: When solid is specified, the
foreground color (fgColor) is the only color rendered,
even when a background color (bgColor) is also
specified. */
val bg = cf.getPatternFillColorBack()
val fg = cf.getPatternFillColor()
this.setFill(fill, fg, bg)
}
else
{
val fg = cf.foregroundColor
if (fg > -1)
foregroundColor = fg
}
}
}

/**
 * Creates a FormatHandle for the given cell. **This is not part of the
 * public API and should not be called.** Customers should use
 * [CellHandle.getFormatHandle] instead.
*/
constructor(c:CellHandle) {
workBook = c.cell!!.workBook
if (c.cell!!.xfRec != null)
{
xf = c.cell!!.xfRec
formatId = xf!!.idx // update the pointer - 20071010 KSC
}
else
{ // ?? create new
// 20090512 KSC: Shigeo NPE error formaterror646694, create new
// rather than outputting warning
// Logger.logWarn("No XF for cell " + c.toString());
xf = workBook!!.getXf(c.cell!!.ixfe)
formatId = xf!!.idx
}
// 20101201 KSC: only add to cache when adding xf's addToCache();
}

/**
 * overrides the equals method to perform equality based on format
 * properties ------------------------------------------------------------
 *
 * @param Object another - the FormatHandle to compare with this FormatHandle
 * @return true if this FormatHandle equals another
*/
public override fun equals(another:Any?):Boolean {
return another!!.toString() == toString()
}

/**
 * remove borders for this format
*/
fun removeBorders() {
val xf = cloneXf(this.xf!!)
xf.removeBorders()
updateXf(xf)
}

/**
 * sets the border color for all borders (top, left, bottom and right) from
 * an int array containing color constants
 *
 *
 *
 * NOTE: this setting will affect every cell which refers to this
 * FormatHandle
 *
 * @param int[] bordercolors - 4-element array of desired border color
 * constants [T, L, B, R]
 * @see FormatHandle.COLOR_* constants
*/
fun setBorderColors(bordercolors:IntArray) {
val xf = cloneXf(this.xf!!)
xf.topBorderColor = bordercolors[0]
xf.leftBorderColor = bordercolors[1]
xf.bottomBorderColor = bordercolors[2]
xf.setRightBorderColor(bordercolors[3])
updateXf(xf)

}

/**
 * set the border color for all borders (top, left, bottom, and right) to
 * one color via color constant
 *
 *
 *
 * NOTE: this setting will affect every cell which refers to this
 * FormatHandle
 *
 * @param int x - color constant which represents the color to set all
 * border sides
 * @see FormatHandle.COLOR_* constants
*/
fun setBorderColor(x:Int) {
val xf = cloneXf(this.xf!!)
xf.setRightBorderColor(x.toShort().toInt())
xf.setLeftBorderColor(x.toShort())
xf.topBorderColor = x.toShort()
xf.bottomBorderColor = x.toShort()
updateXf(xf)

}

/**
 * set the border color for all borders (top, left, bottom, and right) to
 * one java.awt.Color
 *
 *
 * NOTE: this setting will affect every cell which refers to this
 * FormatHandle
 *
 * @param Color col - color to set all border sides
*/
fun setBorderColor(col:Color) {
val xf = cloneXf(this.xf!!)
val x = getColorInt(col).toShort()
xf.setRightBorderColor(x.toInt())
xf.setLeftBorderColor(x)
xf.topBorderColor = x
xf.bottomBorderColor = x
updateXf(xf)

}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderRightColor(x:Int) {
val xf = cloneXf(this.xf!!)
xf.setRightBorderColor(x.toShort().toInt())
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderRightColor(x:Color) {
val xf = cloneXf(this.xf!!)
xf.setRightBorderColor(getColorInt(x).toShort().toInt())
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderLeftColor(x:Int) {
val xf = cloneXf(this.xf!!)
xf.setLeftBorderColor(x.toShort())
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderLeftColor(x:Color) {
val xf = cloneXf(this.xf!!)
xf.setLeftBorderColor(getColorInt(x).toShort())
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderTopColor(x:Int) {
val xf = cloneXf(this.xf!!)
xf.topBorderColor = x
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderTopColor(x:Color) {
val xf = cloneXf(this.xf!!)
xf.topBorderColor = getColorInt(x)
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderBottomColor(x:Int) {
val xf = cloneXf(this.xf!!)
xf.bottomBorderColor = x
updateXf(xf)
}

/**
 * Set the top border color
 *
 * @param color constant
*/
fun setBorderBottomColor(x:Color) {
val xf = cloneXf(this.xf!!)
xf.bottomBorderColor = getColorInt(x)
updateXf(xf)
}

/**
 * Sets the border line style using static BORDER_ shorts within
 * FormatHandle
 *
 * @param line style constant
*/
fun setBorderLineStyle(x:Int) {
val xf = cloneXf(this.xf!!)
xf.setBorderLineStyle(x.toShort())
updateXf(xf)
}

/**
 * set border line styles via array of ints representing border styles
 * order= left, right, top, bottom, [diagonal]
 *
 * @param b int[]
*/
fun setBorderLineStyle(b:IntArray) {
xf!!.setAllBorderLineStyles(b)
}

/**
 * Set the border line style
 *
 * @param line style constant
*/
fun setBorderLineStyle(x:Short) {
val xf = cloneXf(this.xf!!)
xf.setBorderLineStyle(x)
updateXf(xf)
}

/**
 * Set the border line style
 *
 * @param line style constant
*/
fun setBorderDiagonal(x:Int) {
val xf = cloneXf(this.xf!!)
xf.setBorderDiag(x)
updateXf(xf)
}

/**
 * Set a column handle on this format handle, so all changes applied to this
 * format will be applied to the entire column
 *
 * @param c
*/
fun setColHandle(c:ColHandle) {
mycol = c
}

/**
 * Set a row handle on this format handle, so all changes applied to this
 * format will be applied to the entire row
 *
 * @param c
*/
fun setRowHandle(c:RowHandle) {
myrow = c
}

/**
 * Create a copy of this FormatHandle with its own Xf
 *
 * @return the copied FormatHandle
*/
fun clone():Any {
var ret:FormatHandle? = null
if (wbh == null)
{ // who knew???
workBook = xf!!.workBook // Changed to myxf since myfont is no
// longer
ret = FormatHandle()
ret!!.xf = xf
ret!!.formatId = xf!!.idx
ret!!.workBook = workBook
}
else
{
ret = FormatHandle(wbh, xf) // no need to duplicate it - just
// use all formatting of
// original xf
}

return ret
}

public override fun toString():String {
return xf!!.toString()
}

constructor(book:io.starter.OpenXLS.WorkBook, fontname:String,
fontstyle:Int, fontsize:Int) : this(book) {
setFont(fontname, fontstyle, fontsize.toDouble())
}

/**
 * Jan 27, 2011
 *
 * @param workBook
 * @param i
*/
constructor(workBook:io.starter.OpenXLS.WorkBook, i:Int) : this(workBook.workBook, i) {
wbh = workBook
}

/**
 * Set new font to XF, handling duplication and caching ...
 *
 * @param Font f
*/
private fun setXFToFont(f:Font) {
val fti = addFontIfNecessary(f)
if (xf == null)
{ // shouldn't!!
xf = duplicateXf(null)
xf!!.setFont(fti)
}
else if (xf!!.ifnt.toInt() != fti)
{ // if not using font already,
// duplicate xf and set to new font
val xf = cloneXf(this.xf!!)
xf.setFont(fti)
updateXf(xf)
}

}

/**
 * Sets this format handle to a font
 *
 * @param fn  font name e.g. 'Arial'
 * @param stl font style either Font.PLAIN or Font.BOLD
 * @param sz  font size or height in 1/20 point units
*/
fun setFont(fn:String, stl:Int, sz:Double) {
var stl = stl
var sz = sz
sz *= 20.0
if (stl == 0)
stl = 200
if (stl == 1)
stl = 700
val f = Font(fn, stl, sz.toInt())
setXFToFont(f)
}

/**
 * add font to record streamer if cant find exact font already in there
 *
 * @param f
 * @return
*/
private fun addFontIfNecessary(f:Font?):Int {
if (workBook == null)
{
Logger.logErr("AddFontIfNecessary: workbook is null")
return -1
}
var fti = workBook!!.getFontIdx(f)
// don't use the built-ins.
if (fti == 3)
fti = 0 // use initial default font instead of last ...
if (fti == -1)
{ // font doesn't exist yet, add to streamer
f!!.idx = -1 // flag to insert
fti = workBook!!.insertFont(f) + 1
}
else
f!!.idx = fti
return fti
}

/**
 * Adds a font internally to the workbook
 *
 * @param f
 *
 * private void addFont(Font f) {
 *
 * }
*/

/**
 * Apply this Format to a Range of Cells
 *
 * @param CellRange to apply the format to
*/
fun addCellRange(cr:CellRange) {
val crcells = cr.getCells()
for (t in crcells!!.indices)
{
addCell(crcells!![t])
}
}

/**
 * Apply this Format to a Range of Cells
 *
 * @param CellHandle array to apply the format to
*/
fun addCellArray(crcells:Array<CellHandle>) {
for (t in crcells.indices)
{
addCell(crcells[t])
}
}

/**
 * add a Cell to this FormatHandle thus applying the Format to the Cell
 *
 * @param CellHandle to apply the format to
*/
fun addCell(c:CellHandle) {
c.formatHandle = this
}

/**
 * Add a List of Cells to this FormatHandle
 *
 * @param cx
*/
fun addCells(cx:List<*>) {
val itx = cx.iterator()
while (itx.hasNext())
{
addCell(itx.next() as BiffRec)
mycells.add(itx.next())
}
}

internal fun addCell(c:BiffRec) {
if (xf != null)
{
c.setXFRecord(xf!!.idx)
}
else
{
Logger.logWarn("FormatHandle.addCell() - You MUST call setFont() to initialize the FormatHandle's font before adding Cells.")
}
mycells.add(c)
}

/**
 * Applies the format to a cell without establishing a relationship. The
 * format represented by this `FormatHandle` will be applied to
 * the cell but it will not be updated with any future changes. If you want
 * that behavior use [addCell][.addCell] instead.
*/
fun stamp(cell:CellHandle) {
cell.formatId = formatId
}

/**
 * Applies the format to a cell range without establishing a relationship.
 * The format represented by this `FormatHandle` will be applied
 * to the cells but they will not be updated with any future changes. If you
 * want that behavior use [addCell][.addCell] instead.
*/
fun stamp(range:CellRange) {
try
{
range.setFormatID(formatId)
}
catch (e:Exception) {
// This can't actually happen
}

}

/**
 * set the Background Pattern for this Format
 *
 * @param int Excel color constant
*/
fun setBackgroundPattern(t:Int) {
val xf = cloneXf(this.xf!!)
// 20080103 KSC: handle solid (=filled) backgrounds, in which Excel
// switches fg and bg colors (!!!)
if (t != FormatConstants.PATTERN_FILLED)
xf.setPattern(t)
else
{
val bg = xf.backgroundColor.toInt()
xf.setBackgroundSolid()
xf.setForeColor(bg, null)
}
updateXf(xf)
}

/**
 * super/sub (0 = none, 1 = super, 2 = sub)
 *
 * @param int script type for Format Font
*/
fun setScript(ss:Int) {
var ss = ss
if (ss > 2)
ss = 2 // deal with invalid numbers
if (ss < 0)
ss = 0 // deal with invalid numbers
xf!!.font!!.script = ss
}

/**
 * set the Font Color for this Format via indexed color constant
 *
 * @param int Excel color constant
*/
fun setFontColor(t:Int) {
val f = cloneFont(xf!!.font!!)
f.color = t
updateFont(f)
}

/**
 * set the Font Color for this Format
 *
 * @param AWT Color color constant
*/
fun setFontColor(colr:Color) {
val f = cloneFont(xf!!.font!!)
f.setColor(colr)
updateFont(f)
}

/**
 * sets the Font color for this Format via web Hex String
 *
 * @param clr
*/
fun setFontColor(clr:String) {
val f = cloneFont(xf!!.font!!)
f.setColor(clr)
updateFont(f)
}

/**
 * set the foreground Color for this Format NOTE: Foreground color = the
 * CELL BACKGROUND color color for all patterns and Background color= the
 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
 * color=CELL BACKGROUND color and Background Color=64 (white).
 *
 * @param AWT Color constant
*/
fun setForegroundColor(colr:Color) {
val xf = cloneXf(this.xf!!)
val clrz = getColorInt(colr)
xf.setForeColor(clrz, colr)
updateXf(xf)
}

/**
 * set the Cell Background Color for this Format
 *
 *
 * NOTE: Foreground color = the CELL BACKGROUND color color for all patterns
 * and Background color= the PATTERN color for all patterns != Solid For
 * PATTERN_SOLID, Foreground color=CELL BACKGROUND color and Background
 * Color=64 (white).
 *
 * @param awt Color color constant
*/
fun setCellBackgroundColor(colr:Color) {
val clrz = getColorInt(colr)
setCellBackgroundColor(clrz)

}

/**
 * makes the cell a solid pattern background if no pattern was already
 * present NOTE: Foreground color = the CELL BACKGROUND color color for all
 * patterns and Background color= the PATTERN color for all patterns !=
 * Solid For PATTERN_SOLID, Foreground color=CELL BACKGROUND color and
 * Background Color=64 (white).
 *
 * @param int Excel color constant
*/
fun setCellBackgroundColor(t:Int) {
val xf = cloneXf(this.xf!!)

if (xf.fillPattern == 0)
xf.setBackgroundSolid()

if (xf.fillPattern == Xf.PATTERN_SOLID)
xf.setForeColor(t, null)
else
xf.setBackColor(t, null)
updateXf(xf)
}

/**
 * sets this fill pattern from an existing OOXML (2007v) fill element
 *
 * @param f
*/
protected fun setFill(f:Fill?) {
val xf = cloneXf(this.xf!!)
xf.setFill(f!!)
updateXf(xf)
}

/**
 * sets the fill for this format handle if fill==Xf.PATTERN_SOLID then fg is
 * the PATTERN color i.e the CELL BG COLOR
 *
 * @param fillpattern
 * @param fg
 * @param bg
*/
fun setFill(fillpattern:Int, fg:Int, bg:Int) {
val xf = cloneXf(this.xf!!)

xf.setPattern(fillpattern)

/**
 * If the fill style is solid: When solid is specified, the foreground
 * color (fgColor) is the only color rendered, even when a background
 * color (bgColor) is also specified.
*/
if (xf.fillPattern == Xf.PATTERN_SOLID)
{ // is reversed
xf.setForeColor(bg, null)
xf.setBackColor(64, null)
}
else
{
/**
 * or cell fills with patterns specified, then the cell fill color
 * is specified by the bgColor element
*/
xf.setForeColor(fg, null)
xf.setBackColor(bg, null)
}
updateXf(xf)
}

/**
 * set the Background Color for this Format NOTE: Foreground color = the
 * CELL BACKGROUND color color for all patterns and Background color= the
 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
 * color=CELL BACKGROUND color and Background Color=64 (white).
 *
 * @param awt Color color constant
*/
fun setBackgroundColor(colr:Color) {
val xf = cloneXf(this.xf!!)
val clrz = getColorInt(colr)
xf.setBackColor(clrz, colr)
updateXf(xf)
}

/**
 * Get if the font is underlined
 *
 * @return boolean representing if the formatted cell is underlined.
*/
fun getUnderlined():Boolean {
if (xf!!.font == null)
return underlined
return xf!!.font!!.underlined
}

/**
 * Set underline attribute on the font
 *
 * @param isUnderlined boolean representing if the formatted cell should be
 * underlined
*/
fun setUnderlined(isUnderlined:Boolean) {
val f = cloneFont(xf!!.font!!)
f.underlined = isUnderlined
updateFont(f)

}

fun setPattern(pat:Int) {
val xf = cloneXf(this.xf!!)
xf.setPattern(pat)
updateXf(xf)
}

/**
 * increase or decrease the precision of this numeric or curerncy format pattern
 * <br></br>If the format pattern is "General", converts to a basic number pattern
 * <br></br>If the precision is already 0 and !increase, this method does nothing
 *
 * @param increase true if increase the precsion (number of decimals to display)
*/
fun adjustPrecision(increase:Boolean) {
// TODO:  if decimal is contained within quotes ...
var pat = formatPattern
if (pat == "General" && increase)
{
pat = "0.0"    // the most basic numeric pattern
this.formatPattern = pat
return
}

try
{
// split pattern and deal with positive, negative, zero and text separately
// for each, find decimal place and increment/decrement; if not found, find last digit placeholder
val pats = pat!!.split((";").toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
var newPat = ""
for (i in pats.indices)
{
if (i > 0) newPat += ';'.toString()
var z = pats[i].indexOf('.')    // position of decimal
var foundit = false
if (z != -1)
{    // found decimal place
z++
while (z < pats[i].length)
{
val c = pats[i].get(z)
if ((c == '0' || c == '#' || c == '?'))
// numeric placeholders
foundit = true
else if (foundit && !(c == '0' || c == '#' || c == '?'))
// numeric placeholders. if hit last one, either inc or dec
break
z++
}
if (increase)
newPat += StringBuffer(pats[i]).insert(z, "0").toString()    //pats[i].substring(0, z) + "0" + pats[i].substring(z+1);
else
{
if (pats[i].get(z - 2) != '.')
{
newPat += StringBuffer(pats[i]).deleteCharAt(z - 1).toString()    //  .pats[i].substring(0, z-1) + pats[i].substring(z+1);
}
else
{
newPat += StringBuffer(pats[i]).delete(z - 2, z).toString()
}
}

}
else if (increase)
{ // no decimal yet.  If decrease, ignore.  if increase, add
z = pats[i].length - 1
while (z >= 0)
{
val c = pats[i].get(z)
if ((c == '0' || c == '#' || c == '?'))
{    // found last numeric placeholder
foundit = true
break
}
z--
}
if (foundit)
// if had ANY numeric placeholders
newPat += StringBuffer(pats[i]).insert(z + 1, ".0").toString()    //pats[i].substring(0, z) + ".0" + pats[i].substring(z+1);
else
newPat += pats[i]    // keep original
}
else
// if decrease and no decimal found, leave alone
newPat += pats[i]    // leave alone
}

//io.starter.toolkit.Logger.log("Old Style" + pat + ".  New Style: " + newPat + ". Increase?" +  (increase?"yes":"no"));	// KSC: TESETING: TAKE OUT WHEN DONE
this.formatPattern = newPat
}
catch (e:Exception) {
Logger.logErr("Error setting style")    // KSC: TESETING: TAKE OUT WHEN DONE
}

}

/**
 * sets the Right to Left Text Direction or reading order of this style
 *
 * @param rtl possible values:
 * <br></br>0=Context Dependent
 * <br></br>1=Left-to-Right
 * <br></br>2=Right-to-Let
 * @param rtl possible values: <br></br>
 * 0=Context Dependent <br></br>
 * 1=Left-to-Right <br></br>
 * 2=Right-to-Let
*/
fun setRightToLeftReadingOrder(rtl:Int) {
val xf = cloneXf(this.xf!!)
xf.rightToLeftReadingOrder = rtl
updateXf(xf)
}

/**
 * DEPRECATED and non functional. Not neccesary as this occurs automatically
 *
 *
 * Consolidates this Format with other identical formats in workbook
 *
 *
 *
 *
 * There is a limit to the number of distinct formats in an Excel workbook.
 *
 *
 * This method allows you to share Formats between identically formatted
 * cells.
 *
 * @return
*/
@Deprecated("")
fun pack():FormatHandle {
return this // wkbook.cache.get(this);
}

/**
 * Get a JSON representation of the format
 *
 * @param cr
 * @return
*/
fun getJSON(XFNum:Int):String {
return getJSONObject(XFNum).toString()
}

/**
 * Get a JSON representation of the format
 *
 *
 * font height is represented as HTML pt size
 *
 * @param cr
 * @return
*/
fun getJSONObject(XFNum:Int):JSONObject {
val myf = font
val theStyle = JSONObject()
try
{
theStyle.put("style", XFNum)
// handle the font
val theFont = JSONObject()
theFont.put("name", myf!!.getFontName())

// round out the font size...
val sz = Math.round(myf!!.fontHeight / 22.0)

theFont.put("size", sz) // adjust smaller
theFont.put("color", colorToHexString(myf!!.colorAsColor))
theFont.put("weight", myf!!.fontWeight)
if (isBold)
{
theFont.put("bold", "1")
}
if (this.getUnderlined())
{
theFont.put("underline", "1")
}
else
{
theFont.put("underline", "0")
}

if (italic)
{
theFont.put("italic", "1")
}

theStyle.put("font", theFont)

// <Borders>
val border = JSONObject()
if (rightBorderLineStyle != 0)
{
val rBorder = JSONObject()
rBorder.put("style",
FormatConstants.BORDER_STYLES_JSON[rightBorderLineStyle])
rBorder.put("color", colorToHexString(borderRightColor!!))
border.put("right", rBorder)
}
if (bottomBorderLineStyle != 0)
{
val bBorder = JSONObject()
bBorder.put("style",
FormatConstants.BORDER_STYLES_JSON[bottomBorderLineStyle])
bBorder.put("color", colorToHexString(borderBottomColor!!))
border.put("bottom", bBorder)
}
if (leftBorderLineStyle != 0)
{
val lBorder = JSONObject()
lBorder.put("style",
FormatConstants.BORDER_STYLES_JSON[leftBorderLineStyle])
lBorder.put("color", colorToHexString(borderLeftColor!!))
border.put("left", lBorder)
}
if (topBorderLineStyle != 0)
{
val tBorder = JSONObject()
tBorder.put("style",
FormatConstants.BORDER_STYLES_JSON[topBorderLineStyle])
tBorder.put("color", colorToHexString(borderTopColor!!))
border.put("top", tBorder)
}
theStyle.put("borders", border)

// <Alignment>
val alignment = JSONObject()
alignment.put("horizontal",
FormatConstants.HORIZONTAL_ALIGNMENTS[horizontalAlignment])
alignment.put("vertical",
FormatConstants.VERTICAL_ALIGNMENTS[verticalAlignment])
if (wrapText)
{
alignment.put("wrap", "1")
}
theStyle.put("alignment", alignment)

if (indent != 0)
{
theStyle.put("indent", indent)
}
// <Interior> colors + background patterns
if (fillPattern >= 0)
{ // KSC: added >= 0 as some conditional formats have patternFillStyle==0 even though they have a pattern block and
val interior = JSONObject()
// weird black/white case
if (xf!!.foregroundColor.toInt() == 65)
{
interior.put("color", "#FFFFFF")
}
else
{
// KSC: use color string if it exists; create if doesn't
interior.put("color", xf!!.foregroundColorHEX)
}
interior.put("pattern", xf!!.fillPattern)
interior.put("fg", xf!!.foregroundColor.toInt()) // Excel-2003 Color Table index
interior.put("patterncolor", xf!!.backgroundColorHEX)
interior.put("bg", xf!!.backgroundColor.toInt()) // Excel-2003 Color Table index
theStyle.put("interior", interior)
}

if (xf!!.ifmt.toInt() != 0)
{ // only input user defined formats ...
val nFormat = JSONObject()
val fmtpat = formatPattern
try
{
if (fmtpat != "General")
{
nFormat.put("format", fmtpat) // convertXMLChars?
nFormat.put("formatid", xf!!.ifmt.toInt())
if (isDate)
nFormat.put("isdate", "1")
if (isCurrency)
nFormat.put("iscurrency", "1")
if (isRedWhenNegative)
nFormat.put("isrednegative", "1")
}

}
catch (e:Exception) { // it's possible that getFormatPattern
// returns null
}

theStyle.put("numberformat", nFormat)
}
// <Protection>
val protection = JSONObject()
try
{
if (this.xf!!.isLocked)
protection.put("Protected", true)
if (this.xf!!.isFormulaHidden)
protection.put("HideFormula", true)
}
catch (e:Exception) {}

theStyle.put("protection", protection)

}
catch (e:JSONException) {
Logger.logErr("Error getting cellRange JSON: " + e)
}

return theStyle
}

/**
 * Returns an XML fragment representing the FormatHandle
 *
 * @param convertToUnicodeFont if true, font family will be changed to ArialUnicodeMS
 * (standard unicode) for non-ascii fonts
*/
@JvmOverloads  fun getXML(XFNum:Int, convertToUnicodeFont:Boolean = false):String {
val myf = font
// <Style =main element
val sb = StringBuffer("<Style")
sb.append(" ID=\"s" + XFNum + "\"")
sb.append(">")
// <Font>
sb.append("<Font " + myf!!.getXML(convertToUnicodeFont) + "/>")

// <Borders>
sb.append("<Borders>")
if (rightBorderLineStyle != 0)
{
sb.append("		<Border")
sb.append(" Position=\"right\"")
sb.append((" LineStyle=\"" + FormatConstants.BORDER_NAMES[rightBorderLineStyle]
+ "\""))
sb.append((" Color=\"" + colorToHexString(borderRightColor!!)
+ "\""))
sb.append((" Weight=\""
+ FormatConstants.BORDER_SIZES_HTML[rightBorderLineStyle] + "\""))
sb.append("/>")
}
if (bottomBorderLineStyle != 0)
{
sb.append("<Border")
sb.append(" Position=\"bottom\"")
sb.append(" LineStyle=\"" + FormatConstants.BORDER_NAMES[bottomBorderLineStyle] + "\"")
sb.append((" Color=\"" + colorToHexString(borderBottomColor!!)
+ "\""))
sb.append((" Weight=\""
+ FormatConstants.BORDER_SIZES_HTML[bottomBorderLineStyle] + "\""))
sb.append("/>")
}
if (leftBorderLineStyle != 0)
{
sb.append("<Border")
sb.append(" Position=\"left\"")
sb.append((" LineStyle=\"" + FormatConstants.BORDER_NAMES[leftBorderLineStyle]
+ "\""))
sb.append((" Color=\"" + colorToHexString(borderLeftColor!!)
+ "\""))
sb.append((" Weight=\""
+ FormatConstants.BORDER_SIZES_HTML[leftBorderLineStyle] + "\""))
sb.append("/>")
}
if (topBorderLineStyle != 0)
{
sb.append("<Border")
sb.append(" Position=\"top\"")
sb.append((" LineStyle=\"" + FormatConstants.BORDER_NAMES[topBorderLineStyle]
+ "\""))
sb.append((" Color=\"" + colorToHexString(borderTopColor!!)
+ "\""))
sb.append((" Weight=\"" + FormatConstants.BORDER_SIZES_HTML[topBorderLineStyle]
+ "\""))
sb.append("/>")
}
sb.append("</Borders>")

// <Alignment>
sb.append("<Alignment")
sb.append((" Horizontal=\""
+ FormatConstants.HORIZONTAL_ALIGNMENTS[horizontalAlignment] + "\""))
sb.append((" Vertical=\"" + FormatConstants.VERTICAL_ALIGNMENTS[verticalAlignment]
+ "\""))
if (wrapText)
{
sb.append(" Wrap=\"1\"")
}
if (xf!!.indent > 0)
sb.append(" Indent=\"" + (xf!!.indent * 10) + "px\"") // indent // # * 3 // spaces // - do // an approximate conversion
sb.append(" />")

// <Interior> colors + background patterns
/*
 * NOTE: Foreground color = the CELL BACKGROUND color color for all
 * patterns and Background color= the PATTERN color for all patterns !=
 * Solid For PATTERN_SOLID, Foreground color=CELL BACKGROUND color and
 * Background Color=64 (white).
 *
 * NOTE: Interior Color (html color string)===fg (color #) and
 * PatternColor (html color string)===bg (color #) because it's too hard
 * to interpret Excel's "automatic" concept, sometimes it need white as
 * 1 or 9 or 64 ... very difficult to figure out how to convert back to
 * correct color int given just a color string ...
*/
sb.append("<Interior")
if (fillPattern > 0)
{ // 20070201 KSC: Background colors only valid if there is a fill pattern
// 20080815 KSC: use myxf.getForeground/getBackgroundColor as FH
// vers will switch depending on fill pattern
// Also put Pattern element last in so re-creatng will NOT switch
// fg/bg and create a confusing morass :)
val fg = xf!!.foregroundColor.toInt()
// ****************************************
// PDf processing shouldn't output white background due to z-order and overwriting image/chart objects
// ... possibly other uses need the white bg set ...???
// ****************************************
if (!(xf!!.fillPattern == FormatConstants.PATTERN_FILLED && this.workBook!!.colorTable[fg] == Color.WHITE))
{
sb.append(" Color=\"" + colorToHexString(this.workBook!!.colorTable[fg]) + "\"" + " Fg=\"" + fg + "\"")
}
sb.append(" PatternColor=\"" + colorToHexString(this.workBook!!.colorTable[xf!!.backgroundColor]) + "\"" + " Bg=\"" + xf!!.backgroundColor + "\"")
sb.append(" Pattern=\"" + xf!!.fillPattern + "\"")
}
sb.append(" />")
// <NumberFormat>
if (xf!!.ifmt.toInt() != 0)
{ // only input user defined formats ...
val fmtpat = formatPattern
try
{
sb.append("<NumberFormat")
if (fmtpat != "General")
{
sb.append((" Format=\"" + StringTool.convertXMLChars(fmtpat)
+ "\""))
sb.append(" FormatId=\"" + xf!!.ifmt + "\"")
if (isDate)
sb.append(" IsDate=\"1\"")
if (isCurrency)
sb.append(" IsCurrency=\"1\"")
}
}
catch (e:Exception) { // it's possible that getFormatPattern
// returns null
}
finally
{
sb.append(" />")
}
}
// <Protection>
// only input user defined formats ...
val locked = this.xf!!.isLocked
var lck = 0
if (locked)
lck = 1

val formulahidden = this.xf!!.isFormulaHidden
var fmlz = 0
if (formulahidden)
fmlz = 1

try
{
sb.append("<Protection")
sb.append(" Protected=\"" + lck + "\"")
sb.append(" HideFormula=\"" + fmlz + "\"")
}
catch (e:Exception) {
Logger.logWarn(("FormatHandle.getXML problem with protection setting: " + e.toString()))
}
finally
{
sb.append(" />")
}

sb.append("</Style>")
return sb.toString()
}

/**
 * creates a new font based on an existing one, adds to workbook recs
*/
private fun createNewFont(f:Font):Font {
f.idx = -1 // flag to insert anew (see Font.setWorkBook)
workBook!!.insertFont(f)
f.workBook = workBook
return f
}

/**
 * create a new font based on existing font - does not add to workbook recs
 *
 * @param src font
 * @return cloned font
*/
private fun cloneFont(src:Font):Font {
val f = Font()
f.opcode = io.starter.formats.XLS.XLSConstants.FONT
f.setData(src.bytes) // use default font as basis of new font
f.idx = -2    // avoid adding to fonts array when call setW.b. below
f.workBook = this.workBook
f.init()
return f
}

/**
 * update the font for this FormatHandle, including updating the xf if
 * necessary
 *
 * @param f
*/
private fun updateFont(f:Font) {
var f = f
val idx = workBook!!.getFontIdx(f)
if (idx == -1)
{ // can't find it so add new
f = createNewFont(f)
}
else
f = workBook!!.getFont(idx)
if (f.idx != xf!!.ifnt.toInt())
{ // then updated the font, must create new xf to link to
val xf = cloneXf(this.xf!!)
xf.setFont(f.idx)
updateXf(xf)
}
}

/**
 * Create or Duplicate Xf rec so can alter (pattern, font, colrs ...)
 *
 * @param xf xf to base off of, or null (will create new)
 * @return new Xf
*/
private fun duplicateXf(xf:Xf?):Xf {
var xf = xf
var fidx = 0
if (xf != null)
fidx = xf!!.font!!.idx
xf = Xf.updateXf(xf, fidx, workBook) // clones/creates new based upon
// original
formatId = xf!!.idx // update the pointer
// not used anymore		canModify = true; // if duplicated, it's new and unlinked (thus far)
return xf
}

/**
 * adds all formatting represented by the sourceXf to this workbook, if not
 * already present <br></br>
 * This is used internally for transferring formats from one workbook to
 * another
 *
 * @param xf - sourceXf
 * @return ixfe of added Xf
*/
fun addXf(sourceXf:Xf):Int {
// must handle font first in order to create xf below
// check to see if the font needs to be added in current workbook
val fidx = addFontIfNecessary(sourceXf.font)

/** XF  */
val localXf = FormatHandle.cloneXf(sourceXf, workBook!!.getFont(fidx),
workBook) // clone xf so modifcations don't affect original

/** NUMBER FORMAT  */
val fmt = sourceXf.formatPattern // number format pattern
if (fmt != null)
localXf.formatPattern = fmt // adds new format pattern if not // found

// now check out to see if this particular xf pattern exists; if not, add
updateXf(localXf)
return formatId
}

/**
 * if existing format matches, reuse. otherwise, create new Xf record and
 * add to cache
 *
 * @param xf
*/
private fun updateXf(xf:Xf) {
if (this.xf!!.toString() != xf.toString())
{
if (this.xf!!.useCount <= 1 && formatId > 15)
{ // used only by one cell, OK to modify
if (writeImmediate || workBook!!.formatCache.get(xf.toString()) == null)
{
// myxf hasn't been used yet; modify bytes and re-init ***
val xfbytes = xf.bytes
this.xf!!.setData(xfbytes)
if (xf.fill != null)
this.xf!!.fill = xf.fill!!.cloneElement() as io.starter.formats.OOXML.Fill
this.xf!!.init()
this.xf!!.setFont(this.xf!!.ifnt.toInt()) // set font as well ..
workBook!!.updateFormatCache(this.xf!!) // ensure new xf signature
// is stored
}
else
{
if (this.xf!!.useCount > 0)
this.xf!!.decUseCoount()    // flag original xf that 1 less record is referencing it
this.xf = workBook!!.formatCache.get(xf.toString()) as Xf
formatId = this.xf!!.idx // update the pointer
if (formatId == -1)
// hasn't been added to wb yet - should this ever happen???
this.xf = duplicateXf(xf) // create a duplicate and leave original
else
this.xf!!.incUseCount()
}
}
else
{ // cannot modify original - either find matching or create new
if (this.xf!!.useCount > 0)
this.xf!!.decUseCoount()    // flag original xf that 1 less record is referencing it
if (workBook!!.formatCache.get(xf.toString()) == null)
{ // doesn't exist yet
this.xf = duplicateXf(xf) // create a duplicate and leave original
}
else
{
this.xf = (workBook!!.formatCache.get(xf.toString())) as Xf
formatId = this.xf!!.idx // update the pointer
if (formatId == -1)
// hasn't been added to the record store yet 	// - should ever happen???
this.xf = duplicateXf(xf) // create a duplicate and leave original
else
this.xf!!.incUseCount()
}
}

for (i in mycells.indices)
{
(mycells.get(i) as BiffRec).setXFRecord(formatId) // make sure all linked cells are updated as well
}
if (mycol != null)
mycol!!.formatId = formatId
if (myrow != null)
myrow!!.formatId = formatId
}
}

/**
 * create a duplicate xf rec based on existing ...
 *
 * @param xf
 * @return cloned xf
*/
private fun cloneXf(xf:Xf):Xf {
val clone = Xf(xf.font!!.idx, workBook)
val data = xf.getBytesAt(0, xf.length - 4)
clone.setData(data)
if (xf.fill != null)
clone.fill = xf.fill!!.cloneElement() as io.starter.formats.OOXML.Fill
clone.init()
return clone
}

/**
 * clear out object references
*/
fun close() {
mycells.clear()
mycells = CompatibleVector() // all the Cells sharing this format
xf = null
mycol = null
myrow = null
workBook = null
wbh = null
}

@Throws(Throwable::class)
protected fun finalize() {
try
{
close() // close open files
}

finally
{
super.finalize()
}
}

companion object {

val numericFormatMap:Map<String, String>
val currencyFormatMap:Map<String, String>
val dateFormatMap:Map<String, String>

init{
val formats = HashMap<String, String>()
for (formatArr in FormatConstants.NUMERIC_FORMATS)
{
if (formatArr.size == 3)
{
formats.put(formatArr[0].toLowerCase(), formatArr[2])
}
}
numericFormatMap = Collections.unmodifiableMap<String, String>(formats)

formats = HashMap<String, String>()
for (formatArr in FormatConstants.CURRENCY_FORMATS)
{
if (formatArr.size == 3)
{
formats.put(formatArr[0].toLowerCase(), formatArr[2])
}
}
currencyFormatMap = Collections.unmodifiableMap<String, String>(formats)

formats = HashMap<String, String>()
for (formatArr in FormatConstants.DATE_FORMATS)
{
if (formatArr.size == 3)
{
formats.put(formatArr[0].toLowerCase(), formatArr[2])
}
else
{
// only 2 elements in date format string array: pattern and hex id
formats.put(formatArr[0].toLowerCase(), formatArr[0])
}
}
dateFormatMap = Collections.unmodifiableMap<String, String>(formats)
}

/**
 * converts an Excel-style format string to a Java Format string.
 *
 * @param String pattern - Excel Format String
 * @return String that can be used with the Java Format classes.
 * @see getJavaFormatString
*/
fun convertFormatString(pattern:String):String? {
var ret:String? = numericFormatMap.get(pattern)
if (ret != null)
return ret
ret = currencyFormatMap.get(pattern)
if (ret != null)
return ret
ret = dateFormatMap.get(pattern)
return ret
}

/**
 * adds a font to the global font store only if exact font is not already
 * present
 *
 * @param f  Font
 * @param bk WorkBookHandle
*/
fun addFont(f:Font, bk:WorkBookHandle?):Int {
if (bk == null)
{
Logger.logErr("addFont: workbook is null")
}
var fti = bk!!.workBook!!.getFontIdx(f)
// if (fti > 3) {// don't use the built-ins.
if (fti == 3)
fti = 0 // use initial default font instead of last ... 20070827
// KSC
if (fti == -1)
{ // font doesn't exist yet, add to streamer
f.idx = -1 // flag to insert
fti = bk!!.workBook!!.insertFont(f) + 1
}
else
f.idx = fti
return fti
}

/**
 * returns the index of the Color within the Colortable
 *
 * @param the index of the color in the colortable
 * @return the color
*/
fun getColor(col:Int):Color {
if (col > -1 && col < FormatHandle.COLORTABLE.size)
return FormatHandle.COLORTABLE[col]
return FormatHandle.COLORTABLE[0]
}

/**
 * returns the index of the Color within the Colortable
 *
 * @param col the color
 * @return the index of the color in the colortable
*/
fun getColorInt(col:Color):Int {
for (i in FormatHandle.COLORTABLE.indices)
{
if (col == FormatHandle.COLORTABLE[i])
{
return i
}
}
val R = col.getRed()
val G = col.getGreen()
val B = col.getBlue()
var colorMatch = -1
var colorDiff = Integer.MAX_VALUE.toDouble()
for (i in FormatHandle.COLORTABLE.indices)
{
val curDif = (Math.pow((R - FormatHandle.COLORTABLE[i].getRed()).toDouble(), 2.0)
+ Math.pow((G - FormatHandle.COLORTABLE[i].getGreen()).toDouble(), 2.0)
+ Math.pow((B - FormatHandle.COLORTABLE[i].getBlue()).toDouble(), 2.0))
if (curDif < colorDiff)
{
colorDiff = curDif
colorMatch = i
}
}
return colorMatch
}

/**
 * Convert a java.awt.Color to a hex string.
 *
 * @return String representation of a Color
*/
fun colorToHexString(c:Color):String {
val r = c.getRed()
val g = c.getGreen()
val b = c.getBlue()
var rh = Integer.toHexString(r)
if (rh.length < 2)
{
rh = "0" + rh
}
var gh = Integer.toHexString(g)
if (gh.length < 2)
{
gh = "0" + gh
}
var bh = Integer.toHexString(b)
if (bh.length < 2)
{
bh = "0" + bh
}
return ("#" + rh + gh + bh).toUpperCase()
}


var colorFONT:Short = 0
var colorBACKGROUND:Short = 1
var colorFOREGROUND:Short = 2
var colorBORDER:Short = 3

/**
 * convert hex string RGB to Excel colortable int format if an exact match
 * is not find, does color-matching to try and obtain closest match
 *
 * @param s
 * @param colorType
 * @return
*/
fun HexStringToColorInt(s:String, colorType:Short):Int {
var s = s
if (s.length > 7)
s = "#" + s.substring(s.length - 6)
if (s.indexOf("#") == -1)
s = "#" + s
if (s.length == 7)
{
val rs = s.substring(1, 3)
val r = Integer.parseInt(rs, 16)
val gs = s.substring(3, 5)
val g = Integer.parseInt(gs, 16)
val bs = s.substring(5, 7)
val b = Integer.parseInt(bs, 16)
// Handle exceptions for black, white and color indexes 9 (see
// FormatConstants for more info)

if (r == 255 && r == g && r == b)
{
if (colorType == colorFONT)
return 9
}

val c = Color(r, g, b)
return getColorInt(c)
}
return 0
}

/**
 * convert hex string RGB to Excel colortable int format if no exact match
 * is found returns -1
 *
 * @param s         HTML color string (#XXXXXX) format or (#FFXXXXXX - OOXML-style
 * format)
 * @param colorType
 * @return index into color table or -1 if not found
*/
fun HexStringToColorIntExact(s:String, colorType:Short):Int {
var s = s
if (s.length > 7)
s = "#" + s.substring(s.length - 6)
if (s.indexOf("#") == -1)
s = "#" + s
if (s.length == 7)
{
val rs = s.substring(1, 3)
val r = Integer.parseInt(rs, 16)
val gs = s.substring(3, 5)
val g = Integer.parseInt(gs, 16)
val bs = s.substring(5, 7)
val b = Integer.parseInt(bs, 16)

// Handle exceptions for black, white and color indexes 9 (see
// FormatConstants for more info)
if (r == 255 && r == g && r == b)
{
if (colorType == colorFONT)
return 9
}

val c = Color(r, g, b)
for (i in FormatHandle.COLORTABLE.indices)
{
if (c == FormatHandle.COLORTABLE[i])
{
return i
}
}
}
return -1 // match NOT FOUND
}

/**
 * convert hex string RGB to a Color
 *
 * @param s web-style hex color string, including #, or OOXML-style color string, FFXXXXXX
 * @return Color
 * @see Color
*/
fun HexStringToColor(s:String):Color {
var s = s
if (s.length > 7)
{    // transform OOXML-style color strings to Web-style
s = "#" + s.substring(s.length - 6)
}
var c:Color? = null
if (s.length == 7)
{
val rs = s.substring(1, 3)
val r = Integer.parseInt(rs, 16)
val gs = s.substring(3, 5)
val g = Integer.parseInt(gs, 16)
val bs = s.substring(5, 7)
val b = Integer.parseInt(bs, 16)
c = Color(r, g, b)
}
else
c = Color(0, 0, 0) // default to black?
return c
}

/**
 * interpret color table special entries
 *
 * @param clr color index in range of 0x41-0x4F (charts) 0x40 ...
 * @return index into color table
*/
fun interpretSpecialColorIndex(clr:Int):Short {
when (clr) {
0x0041 // Default background color. This is the window background
,
// color in the sheet display and is the default
// background color for a cell.
0x004E // Default chart background color. This is the window
,
// background color in the chart display.
0x0050 // WHAT IS THIS ONE????????
-> return FormatConstants.COLOR_WHITE.toShort()
0x0040 // Default foreground color
, 0x004F // Chart neutral color which is black, an RGB value of
,
// (0,0,0).
0x004D // Default chart foreground color. This is the window text
,
// color in the chart display.
0x0051 // ToolTip text color. This is the automatic font color for
,
// comments.
0x7FFF // Font automatic color. This is the window text color
-> return FormatConstants.COLOR_BLACK.toShort()
else // 67(=0x43) ???
-> return FormatConstants.COLOR_WHITE.toShort()
}

/*
 * switch (icvFore) { case 0x40: // default fg color return
 * FormatConstants.COLOR_WHITE; case 0x41: // default bg color return
 * FormatConstants.COLOR_WHITE; case 0x4D: // default CHART fg color --
 * INDEX SPECIFIC! return -1; // flag to map via series (bar) color
 * defaults case 0x4E: // default CHART fg color return icvFore; case
 * 0x4F: // chart neutral color == black return
 * FormatConstants.COLOR_BLACK; }
 *
 * switch (icvBack) { case 0x40: // default fg color return
 * FormatConstants.COLOR_WHITE; case 0x41: // default bg color return
 * FormatConstants.COLOR_WHITE; case 0x4D: // default CHART fg color --
 * INDEX SPECIFIC! return -1; // flag to map via series (bar) color
 * defaults case 0x4E: // default CHART bg color //return
 * FormatConstants.COLOR_WHITE; // is this correct? return icvBack; case
 * 0x4F: // chart neutral color == black return
 * FormatConstants.COLOR_BLACK; }
*/
}

fun BorderStringToInt(s:String):Int {
for (i in FormatConstants.BORDER_NAMES.indices)
{
if (FormatConstants.BORDER_NAMES[i] == s)
return i
}
return 0
}

/**
 * static version of cloneXf
 *
 * @param xf
 * @param wkbook
 * @return
*/
fun cloneXf(xf:Xf, wkbook:io.starter.formats.XLS.WorkBook):Xf {
val clone = Xf(xf.font!!.idx, wkbook)
val data = xf.getBytesAt(0, xf.length - 4)
clone.setData(data)
clone.init()
return clone
}

/**
 * static version of cloneXf
 *
 * @param xf
 * @param wkbook
 * @return
*/
fun cloneXf(xf:Xf, f:Font,
wkbook:io.starter.formats.XLS.WorkBook):Xf {
val clone = Xf(f, wkbook)
val data = xf.getBytesAt(0, xf.length - 4)
clone.setData(data)
clone.setFont(f.idx) // font idx is overwritten by xf data; must reset
clone.init()
return clone
}
}
}/**
 * Returns an XML fragment representing the FormatHandle
 */