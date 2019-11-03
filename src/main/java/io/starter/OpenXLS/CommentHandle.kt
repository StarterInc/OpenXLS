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

import io.starter.formats.XLS.Note
import io.starter.toolkit.Logger

/**
 * <pre>
 * CommentHandle allows for manipulation of the Note or Comment feature of Excel
 *
 * In order to create CommentHandles programatically use the methods in WorkSheetHandle or CellHandle
 *
</pre> *
 */
class CommentHandle
/**
 * Creates a new CommentHandle object
 * <br></br>For internal use only
 *
 * @param n
 */
(n: Note) : Handle {

    /**
     * Returns the internal note record for this CommentHandle.
     *
     *
     * Be aware that this note record should not be modified directly, and that this
     * method is only for internal application use.
     *
     * @return Name record
     */
    var internalNoteRec: Note? = null
        private set

    /**
     * Returns the text of the Note (Comment)
     *
     * @return String value or text of Note
     */
    /**
     * Sets the text of the Note (Comment).
     *
     * The text may contain embedded formatting information as follows:
     * <br></br>< font specifics>text segment< font specifics for next segment>text segment...
     * <br></br>where font specifics can be one or more of (all are optional):
     * b - bold
     * <br></br>i - italic
     * <br></br>s - strikethru
     * <br></br>u - underlined
     * <br></br>f="" - font name surrounded by quotes e.g. "Tahoma"
     * <br></br>sz="" - font size in points surrounded by quotes e.g. "10"
     * <br></br>Each option must be delimited by ;'s
     * <br></br>For Example:
     * <br></br>"&#60;f=\"Tahoma\";b;sz=\"16\">Note: &#60;f=\"Cambria\";sz=\"12\">This is an important point"
     * To reset to the default font, input an empty format: <> e.g.:
     * <br></br>"&#60;b;i;sz=\"8\">Note:<>This is an important comment"
     *
     * @param text - String text of Note
     */
    var commentText: String?
        get() = if (internalNoteRec != null) internalNoteRec!!.text else null
        set(text) {
            if (internalNoteRec != null) {
                try {
                    internalNoteRec!!.text = text
                } catch (e: IllegalArgumentException) {
                    Logger.logErr(e.toString())
                }

            }
        }

    /**
     * returns the author of this Note (Comment) if set
     *
     * @return String author
     */
    /**
     * sets the author of this Note (Comment)
     *
     * @param author
     */
    var author: String?
        get() = if (internalNoteRec != null) internalNoteRec!!.author else null
        set(author) {
            if (internalNoteRec != null) internalNoteRec!!.author = author
        }

    /**
     * Returns true if this Note (Comment) is hidden until focus
     *
     * @return
     */
    val isHidden: Boolean
        get() = if (internalNoteRec != null) internalNoteRec!!.hidden else false

    /**
     * Returns the address this Note (Comment) is attached to
     *
     * @return String Cell Address
     */
    val address: String?
        get() = if (internalNoteRec != null) internalNoteRec!!.cellAddressWithSheet else null

    /**
     * return the Row number (0-based) this Note is attached to
     *
     * @return 0-based row number
     */
    val rowNum: Int
        get() = if (internalNoteRec != null) internalNoteRec!!.rowNumber else -1

    /**
     * return the Column this note is attached to
     *
     * @return Column number as an integer e.g. A=0, B=1 ...
     */
    val colNum: Int
        get() = if (internalNoteRec != null) internalNoteRec!!.colNumber.toInt() else -1

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
     * Sets the bounds (size and position) of the Text Box for this Note
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
     */
    var textBoxBounds: ShortArray?
        get() = if (internalNoteRec != null) internalNoteRec!!.textBoxBounds else null
        set(bounds) {
            if (internalNoteRec != null)
                internalNoteRec!!.textBoxBounds = bounds
        }

    init {
        this.internalNoteRec = n
    }

    /**
     * Removes or deletes this Note (Comment) from the worksheet
     */
    fun remove() {
        internalNoteRec!!.sheet!!.removeNote(internalNoteRec!!)
        internalNoteRec = null
    }

    /**
     * Sets this Note (Comment) to always show, even when the attached cell loses focus
     */
    fun show() {
        if (internalNoteRec != null) internalNoteRec!!.hidden = false
    }

    /**
     * Sets this Note (Comment) to be hidden until the attached cell has focus
     *
     *
     * This is the default state of note records
     */
    fun hide() {
        if (internalNoteRec != null) internalNoteRec!!.hidden = true
    }

    /**
     * Sets this Note (Comment) to be attached to a cell at [row, col]
     *
     * @param row int row number (0-based)
     * @param col int column number (0-based)
     */
    fun setRowCol(row: Int, col: Int) {
        if (internalNoteRec != null)
            internalNoteRec!!.setRowCol(row, col)
    }

    /**
     * return the String representation of this CommentHandle
     */
    override fun toString(): String {
        return if (internalNoteRec != null) internalNoteRec!!.toString() else "Not initialized"
    }

    /**
     * Sets the width and height of the bounding text box of the note
     * <br></br>Units are in pixels
     * <br></br>NOTE: the height algorithm w.r.t. varying row heights is not 100%
     *
     * @param width  short desired text box width in pixels
     * @param height short desired text box height in pixels
     */
    fun setTextBoxSize(width: Int, height: Int) {
        if (internalNoteRec != null) {
            internalNoteRec!!.setTextBoxWidth(width.toShort())
            internalNoteRec!!.setTextBoxHeight(height.toShort())
        }
    }

    /**
     * Returns the OOXML representation of this Note object
     *
     * @param authId 0-based author index for the author linked to this Note
     * @return String OOMXL representation
     */
    fun getOOXML(authId: Int): String {
        val ooxml = StringBuffer()
        // TODO: Handle FORMATS
        ooxml.append("<comment ref=\"" + ExcelTools.formatLocation(intArrayOf(internalNoteRec!!.rowNumber, internalNoteRec!!.colNumber.toInt())) + "\" authorId=\"" + authId + "\">")
        ooxml.append(internalNoteRec!!.ooxml)
        ooxml.append("</comment>")
        return ooxml.toString()
    }

}

