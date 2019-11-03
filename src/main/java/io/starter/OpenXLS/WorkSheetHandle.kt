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

import io.starter.formats.XLS.WorkBook
import io.starter.formats.XLS.*
import io.starter.formats.XLS.charts.Chart
import io.starter.formats.XLS.formulas.GenericPtg
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgRef
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool

import java.io.*
import java.sql.Timestamp
import java.util.*

/**
 * The WorkSheetHandle provides a handle to a Worksheet within an XLS file<br></br>
 * and includes convenience methods for working with the Cell values within the
 * sheet.<br></br>
 * <br></br>
 * for example: <br></br>
 * <br></br>
 * <blockquote> WorkBookHandle book = new WorkBookHandle("testxls.xls");<br></br>
 * WorkSheetHandle sheet1 = book.getWorkSheet("Sheet1");<br></br>
 * CellHandle cell = sheet1.getCell("B22");<br></br>
 *
 * <br></br>
 * to add a cell:<br></br>
 * <br></br>
 * CellHandle cell = sheet1.add("Hello World","C22");<br></br>
 *
 * <br></br>
 * to add a numeric cell:<br></br>
 * <br></br>
 * CellHandle cell = sheet1.add(Integer.valueOf(120),"C23");<br></br>
 *
 * <br></br>
 * to add a formula cell:<br></br>
 * <br></br>
 * CellHandle cell = sheet1.add("=PI()","C24");<br></br>
 *
</blockquote> *  <br></br>
 * <br></br>
 *
 * @see WorkSheet
 *
 * @see WorkBookHandle
 *
 * @see CellHandle
 */
class WorkSheetHandle
/**
 * Constructor which takes a WorkBook and sheetname as parameters.
 *
 * @param sht  The name of the WorkSheet
 * @param mybk The WorkBook
 */
(sht: Boundsheet, b: WorkBookHandle) : Handle {

    /**
     * Returns the underlying low-level Boundsheet object.
     *
     * @return Boundsheet sheet
     */
    var sheet: Boundsheet? = null
        private set
    private var mybook: WorkBook? = null
    /**
     * Returns the WorkBookHandle for this Sheet
     *
     * @return
     */
    var workBook: WorkBookHandle? = null
        internal set
    private val DEBUGLEVEL = 0
    private var dateFormats: Hashtable<String, Int>? = Hashtable()
    /**
     * @return setting on whether to use cache or not
     */
    var useCache = true
        private set // 20080917 KSC: set var for caching, default to true [BugTracker 1862]

    /**
     * Get the first row on the Worksheet
     *
     * @return the Minimum Row Number on the Worksheet
     */
    val firstRow: Int
        get() = sheet!!.minRow

    /**
     * Get the first column on the Worksheet
     *
     * @return the Minimum Column Number on the Worksheet
     */
    val firstCol: Int
        get() = sheet!!.minCol

    /**
     * Get the last row on the Worksheet
     *
     * @return the Maximum Row Number on the Worksheet
     */
    val lastRow: Int
        get() = sheet!!.maxRow

    /**
     * Get the last column on the Worksheet
     *
     * @return the Maximum Column Number on the Worksheet
     */
    val lastCol: Int
        get() = sheet!!.maxCol

    /**
     * Gets the hash of the sheet protection password. This method returns the
     * hashed password as stored in the file. It has been passed through a one-way
     * hash function. It is therefore not possible to recover the actual password.
     * You can, however, use [.setProtectionPasswordHashed] to apply the same
     * password to another worksheet.
     *
     * @return the password hash or "0000" if the sheet doesn't have a password
     */
    val hashedProtectionPassword: String
        get() = sheet!!.protectionManager.password

    /**
     * Returns whether the sheet is protected. Note that this is separate from
     * whether the sheet has a protection password. It can be protected without a
     * password or have a password but not be protected.
     *
     * @return whether protection is enabled for the sheet
     */
    /**
     * Sets whether the worksheet is protected.
     *
     * @param protect whether worksheet protection should be enabled
     */
    var protected: Boolean
        get() = sheet!!.protectionManager.protected
        set(protect) {
            sheet!!.protectionManager.protected = protect
        }

    /**
     * get whether this sheet is selected upon opening the file.
     *
     * @return boolean b selected state
     */
    /**
     * set whether this sheet is selected upon opening the file.
     *
     * @param boolean b selected value
     */
    var selected: Boolean
        get() = sheet!!.selected()
        set(b) = sheet!!.setSelected(b)

    /**
     * get whether this sheet is hidden from the user opening the file.
     *
     * @return boolean b hidden state
     */
    /**
     * set whether this sheet is hidden from the user opening the file.
     *
     *
     * if the sheet is selected, the API will set the first visible sheet to
     * selected as you cannot have your selected sheet be hidden.
     *
     *
     * to override this behavior, set your desired sheet to selected after calling
     * this method.
     *
     * @param boolean b hidden state
     */
    // set the next sheet selected...
    var hidden: Boolean
        get() = sheet!!.hidden
        set(b) {
            var h = 0
            if (b)
                h = Boundsheet.HIDDEN.toInt()
            sheet!!.setHidden(h)
            if (sheet!!.sheetNum == 0) {
                try {
                    val s2 = mybook!!.getWorkSheetByNumber(sheet!!.sheetNum + 1)
                    mybook!!.setFirstVisibleSheet(s2)
                } catch (e: WorkSheetNotFoundException) {
                }

            }
            if (sheet!!.selected()) {
                try {
                    var x = 1
                    var s2 = mybook!!.getWorkSheetByNumber(sheet!!.sheetNum + x)
                    while (s2.hidden)
                        s2 = mybook!!.getWorkSheetByNumber(sheet!!.sheetNum + x++)
                    s2.setSelected(true)
                } catch (e: WorkSheetNotFoundException) {
                }

            }
        }

    /**
     * return the 'veryhidden' state of the sheet
     *
     * @return
     */
    /**
     * set whether this sheet is VERY hidden opening the file.
     *
     *
     * VERY hidden means users will not be able to unhide the sheet without using VB
     * code.
     *
     * @param boolean b hidden state
     */
    // set the next sheet selected...
    var veryHidden: Boolean
        get() = sheet!!.veryHidden
        set(b) {
            var h = 0
            if (b)
                h = Boundsheet.VERY_HIDDEN.toInt()
            sheet!!.setHidden(h)
            val t = sheet!!.sheetNum
            try {
                val s2 = mybook!!.getWorkSheetByNumber(t + 1)
                s2.setSelected(true)
            } catch (e: WorkSheetNotFoundException) {
            }

        }

    /**
     * get the tab display order of this Worksheet
     *
     *
     * this is a zero based index with zero representing the left-most WorkSheet
     * tab.
     *
     * @return int idx the index of the sheet tab
     */
    /**
     * set the tab display order of this Worksheet
     *
     *
     * this is a zero based index with zero representing the left-most WorkSheet
     * tab.
     *
     * @param int idx the new index of the sheet tab
     */
    var tabIndex: Int
        get() = sheet!!.sheetNum
        set(idx) = sheet!!.workBook!!.changeWorkSheetOrder(sheet, idx)

    /**
     * returns all of the Columns in this WorkSheet
     *
     * @return ColHandle[] Columns
     */
    val columns: Array<ColHandle>
        get() {
            val columns = CompatibleVector()

            for (c in sheet!!.colinfos) {
                try {
                    val start = c.colFirst
                    val end = c.colLast
                    for (i in start..end) {
                        try {
                            columns.add(this.getCol(i))
                        } catch (e: ColumnNotFoundException) {
                        }

                    }
                } catch (ex: Exception) {
                }

            }

            return columns.toTypedArray() as Array<ColHandle>
        }

    /**
     * returns a List of Column names
     *
     * @return List column names
     */
    val colNames: List<*>
        get() = sheet!!.colNames

    /**
     * returns a List of Row numbers
     *
     * @return List of row numbers
     */
    val rowNums: List<*>
        get() = sheet!!.rowNums

    /**
     * get an array of all RowHandles for this WorkSheet
     *
     * @return RowHandle[] all Rows on this WorkSheet
     */
    val rows: Array<RowHandle>
        get() {
            val rs = sheet!!.rows
            val ret = arrayOfNulls<RowHandle>(rs.size)
            for (t in rs.indices) {
                ret[t] = RowHandle(rs[t], this)
            }
            return ret
        }

    /**
     * get an array of BIFFREC Rows
     *
     * @return RowHandle[] all Rows on this WorkSheet
     */
    val rowMap: Map<*, *>
        get() = sheet!!.rowMap

    /**
     * Returns the number of rows in this WorkSheet
     *
     * @return int Number of Rows on this WorkSheet
     */
    val numRows: Int
        get() = this.sheet!!.numRows

    /**
     * Returns the number of Columns in this WorkSheet
     *
     * @return int Number of Cols on this WorkSheet
     */
    val numCols: Int
        get() = this.sheet!!.numCols

    /**
     * Returns the index of the Sheet.
     *
     * @return String Sheet Name
     */
    val sheetNum: Int
        get() = sheet!!.sheetNum

    /**
     * Returns all Named Range Handles scoped to this Worksheet.
     *
     *
     * Note this will not include workbook scoped named ranges
     *
     * @return NameHandle[] all of the Named ranges that are scoped to the present
     * worksheet
     */
    val namedRangesInScope: Array<NameHandle>
        get() {
            val nand = sheet!!.sheetScopedNames
            val nands = arrayOfNulls<NameHandle>(nand.size)
            for (x in nand.indices) {
                nands[x] = NameHandle(nand[x], this.workBook)
            }
            return nands
        }

    /**
     * Returns the name of the Sheet.
     *
     * @return String Sheet Name
     */
    /**
     * Set the name of the Worksheet. This method will change the name on the
     * Worksheet's tab as displayed in the WorkBook, as well as all programmatic and
     * internal references to the name.
     *
     *
     * This change takes effect immediately, so all attempts to reference the
     * Worksheet by its previous name will fail.
     *
     * @param String the new name for the Worksheet
     */
    // keep sheethandles (name->wsh) updated
    var sheetName: String
        get() = sheet!!.sheetName
        set(name) {
            workBook!!.sheethandles!!.remove(this.sheetName)
            sheet!!.sheetName = name
            workBook!!.sheethandles!![name] = this
        }

    /**
     * return the sheetname properly qualified or quoted used when the sheetname
     * contains spaces, commas or parentheses
     *
     * @return
     */
    val qualifiedSheetName: String?
        get() = GenericPtg.qualifySheetname(sheet!!.sheetName)

    /**
     * Returns the Serialized bytes for this WorkSheet.
     *
     *
     * The output of this method can be used to insert a copy of this WorkSheet into
     * another WorkBook using the WorkBookHandle.addWorkSheet(byte[] serialsheet,
     * String NewSheetName) method.
     *
     * @return byte[] the WorkSheet's Serialized bytes
     * @see WorkBookHandle.addWorkSheet
     */
    // return mysheet.getSheetBytes();
    val serialBytes: ByteArray?
        get() {
            sheet!!.setLocalRecs()
            var obs: ObjectOutputStream? = null
            var b: ByteArray? = null
            try {
                val baos = ByteArrayOutputStream()
                val bufo = BufferedOutputStream(baos)
                obs = ObjectOutputStream(bufo)
                obs.writeObject(sheet)
                bufo.flush()
                b = baos.toByteArray()
            } catch (e: Throwable) {
                Logger.logWarn("Serializing Sheet: $this failed: $e")
            }

            return b
        }

    var addedrows = CompatibleVector()
        private set
    private val range_init = true

    /**
     * returns an array of FormatHandles for the ConditionalFormats applied to this
     * cell
     *
     * @return an array of FormatHandles, one for each of the Conditional Formatting
     * rules
     */
    val conditionalFormatHandles: Array<ConditionalFormatHandle>
        get() {
            val cfx = arrayOfNulls<ConditionalFormatHandle>(this.sheet!!.conditionalFormats!!.size)
            for (i in cfx.indices) {
                val cfmt = this.sheet!!.conditionalFormats!![i] as Condfmt
                cfx[i] = ConditionalFormatHandle(cfmt, this)
            }
            return cfx
        }

    /**
     * Get a handle to all of the images in this worksheet
     *
     * @return
     */
    val images: Array<ImageHandle>?
        get() = this.sheet!!.images

    /**
     * returns the actual amount of images contained in the sheet and is determined
     * by imageMap
     *
     * @return
     */
    val numImages: Int
        get() = this.sheet!!.imageMap!!.size

    /**
     * Get the current fast add cell mode for this worksheet
     */
    /**
     * Toggle fast cell add mode.
     *
     *
     * Set to true to turn off checking for existing cells, conditional formats and
     * merged ranges in order to accelerate adding new cells
     *
     * @param fastadds whether to disable checking for existing cells and
     */
    var fastCellAdds: Boolean
        get() = this.sheet!!.fastCellAdds
        set(fastadds) = this.sheet!!.setFastCellAdds(fastadds)

    /**
     * Returns all CellHandles defined on this WorkSheet.
     *
     * @return CellHandle[] - the array of Cells in the Sheet
     */
    // Handle MULBLANKS proper column number
    // try harder
    // handle Mulblanks: ref a range of cells; to get correct cell address,
    // traverse thru range and set cellhandle ref to correct column
    // for Mulblank use only -sets correct column reference for multiple blank
    // cells ...
    val cells: Array<CellHandle>
        get() {
            val cells = sheet!!.cells
            val retval = arrayOfNulls<CellHandle>(cells.size)
            var aMul: Mulblank? = null
            var c: Short = -1
            for (i in retval.indices) {
                try {
                    if (cells[i].opcode != XLSConstants.MULBLANK) {
                        retval[i] = getCell(cells[i].rowNumber, cells[i].colNumber.toInt())
                    } else {
                        if (cells[i] === aMul) {
                            c++
                        } else {
                            aMul = cells[i] as Mulblank
                            c = aMul.colFirst.toShort()
                        }
                        retval[i] = getCell(cells[i].rowNumber, c.toInt())
                    }
                } catch (cnfe: CellNotFoundException) {
                    retval[i] = CellHandle(cells[i], this.workBook)
                    retval[i].workSheetHandle = this
                    if (cells[i].opcode == XLSConstants.MULBLANK) {
                        if (cells[i] === aMul) {
                            c++
                        } else {
                            aMul = cells[i] as Mulblank
                            c = aMul.colFirst.toShort()
                        }
                        retval[i].setBlankRef(c.toInt())
                    }
                }

            }
            return retval
        }

    /**
     * Get the text for the Footer printed at the bottom of the Worksheet
     *
     * @return String footer text
     */
    /**
     * Set the text for the Footer printed at the bottom of the Worksheet
     *
     * @param String footer text
     */
    var footerText: String?
        get() = sheet!!.footer!!.footerText
        set(t) {
            sheet!!.footer!!.footerText = t
        }

    /**
     * Get the text for the Header printed at the top of the Worksheet
     *
     * @return String header text
     */
    /**
     * Set the text for the Header printed at the top of the Worksheet
     *
     * @param String header text
     */
    var headerText: String?
        get() = sheet!!.header!!.headerText
        set(t) {
            sheet!!.header!!.headerText = t
        }

    /**
     * Get the print area set for this WorkSheetHandle.
     *
     *
     * If no print area is set return null;
     */
    val printArea: String?
        get() = sheet!!.printArea

    /**
     * Get the Print Titles set for this WorkSheetHandle.
     *
     *
     * If no Print Titles are set, this returns null;
     */
    val printTitles: String?
        get() = sheet!!.printTitles

    /**
     * Get the printer settings handle for this WorkSheetHandle.
     */
    val printerSettings: PrinterSettingsHandle
        get() = sheet!!.printerSetupHandle

    /**
     * Set the default column width of the worksheet
     *
     *
     * <br></br>
     *
     *
     * This setting is roughly the width of the character '0' The default width of a
     * column is 8.
     */
    var defaultColWidth: Int
        get() = sheet!!.defaultColumnWidth.toInt()
        set(t) {
            sheet!!.defaultColumnWidth = t
        }

    /**
     * Set whether to show calculated formula results in the output sheet.
     *
     * @return boolean whether to show calculated formula results
     */
    var showFormulaResults: Boolean
        get() = this.sheet!!.window2!!.showFormulaResults
        set(b) {
            this.sheet!!.window2!!.showFormulaResults = b
        }

    /**
     * Get whether to show gridlines in the output sheet.
     *
     * @param boolean whether to show gridlines
     */
    /**
     * Set whether to show gridlines in the output sheet.
     *
     * @return boolean whether to show gridlines
     */
    var showGridlines: Boolean
        get() = this.sheet!!.window2!!.showGridlines
        set(b) {
            this.sheet!!.window2!!.showGridlines = b
        }

    /**
     * Get whether to show sheet headers in the output sheet.
     *
     * @return boolean whether to show sheet headers
     */
    /**
     * Set whether to show sheet headers in the output sheet.
     *
     * @param boolean whether to show sheet headers
     */
    var showSheetHeaders: Boolean
        get() = this.sheet!!.window2!!.showSheetHeaders
        set(b) {
            this.sheet!!.window2!!.showSheetHeaders = b
        }

    /**
     * Get whether to show zero values in the output sheet.
     *
     * @return boolean whether to show zero values
     */
    /**
     * Set whether to show zero values in the output sheet.
     *
     * @return boolean whether to show zero values
     */
    var showZeroValues: Boolean
        get() = this.sheet!!.window2!!.showZeroValues
        set(b) {
            this.sheet!!.window2!!.showZeroValues = b
        }

    /**
     * Get whether to show outline symbols in the output sheet.
     *
     * @return boolean whether to outline symbols
     */
    /**
     * Set whether to show outline symbols in the output sheet.
     *
     * @param boolean whether to show outline symbols
     */
    var showOutlineSymbols: Boolean
        get() = this.sheet!!.window2!!.showOutlineSymbols
        set(b) {
            this.sheet!!.window2!!.showOutlineSymbols = b
        }

    /**
     * Get whether to show normal view or page break preview view in the output
     * sheet.
     *
     * @return boolean whether to show normal view or page break preview view
     */
    /**
     * Set whether to show normal view or page break preview view in the output
     * sheet.
     *
     * @param boolean whether to show normal view or page break preview view
     */
    // opposite of expected behavior
    var showInNormalView: Boolean
        get() = this.sheet!!.window2!!.showInNormalView
        set(b) {
            this.sheet!!.window2!!.showInNormalView = !b
        }

    /**
     * if this sheet has freeze panes, return the address of the top left cell
     * otherwise, return null
     *
     * @return
     */
    val topLeftCell: String?
        get() = if (this.sheet!!.pane != null) {
            this.sheet!!.pane!!.topLeftCell
        } else null

    /**
     * gets the zoom for the sheet
     *
     * @return the zoom as a float percentage (.25 = 25%)
     */
    /**
     * sets the zoom for the sheet
     *
     * @param the zoom as a float percentage (.25 = 25%)
     */
    var zoom: Float
        get() = this.sheet!!.scl.zoom
        set(zm) {
            this.sheet!!.scl.zoom = zm
        }

    /**
     * Gets the row number (0 based) that the sheet split is located on. If the
     * sheet is not split returns -1
     *
     * @return
     */
    val splitRowLocation: Int
        get() = if (this.sheet!!.pane == null) -1 else this.sheet!!.pane!!.visibleRow

    /**
     * gets the column number (0-based)that the sheet split is locaated on; if the
     * sheet is not split, returns -1
     *
     * @return 0-based index of split column, if any
     */
    val splitColLocation: Int
        get() = if (this.sheet!!.pane == null) -1 else this.sheet!!.pane!!.visibleCol

    /**
     * Gets the twips split location returns -1
     *
     * @return
     */
    val splitLocation: Int
        get() = if (this.sheet!!.pane == null) -1 else this.sheet!!.pane!!.rowSplitLoc

    /**
     * Get whether to use manual grid color in the output sheet.
     *
     * @return boolean whether to use manual grid color
     */
    /**
     * Set whether to use manual grid color in the output sheet.
     *
     * @param boolean whether to use manual grid color
     */
    var manualGridLineColor: Boolean
        get() = this.sheet!!.window2!!.manualGridLineColor
        set(b) {
            this.sheet!!.window2!!.manualGridLineColor = b
        }

    /**
     * returns an array of all CommentHandles that exist in the sheet
     *
     * @return
     */
    val commentHandles: Array<CommentHandle>
        get() {
            val notes = getMysheet()!!.notes
            val nHandles = arrayOfNulls<CommentHandle>(notes.size)
            for (i in nHandles.indices) {
                nHandles[i] = CommentHandle(notes[i] as Note)
            }
            return nHandles
        }

    /**
     * Return all validation handles that refer to this worksheet
     *
     * @return array of all validationhandles valid for this worksheet
     */
    val allValidationHandles: Array<ValidationHandle>
        get() {
            if (sheet!!.dvRecs == null)
                return arrayOfNulls(0)
            val vh = arrayOfNulls<ValidationHandle>(sheet!!.dvRecs!!.size)
            val dvrecs = sheet!!.dvRecs
            for (i in vh.indices) {
                vh[i] = ValidationHandle(dvrecs!![i] as Dv)
            }
            return vh
        }

    /**
     * Returns a list of all AutoFilterHandles on this sheet <br></br>
     * An AutoFilterHandle allows access and manipulation of AutoFilters on the
     * sheet
     *
     * @return array of AutoFilterHandles if any exist on sheet, null otherwise
     */
    val autoFilterHandles: Array<AutoFilterHandle>?
        get() {
            if (sheet!!.autoFilters == null)
                return null
            val af = arrayOfNulls<AutoFilterHandle>(sheet!!.autoFilters.size)
            val afs = sheet!!.autoFilters
            for (i in afs.indices) {
                af[i] = AutoFilterHandle(afs[i] as AutoFilter)
            }
            return af
        }
    // public Map cellhandles = new HashMap();

    fun addChart(serialchart: ByteArray, name: String, coords: ShortArray) {
        sheet!!.addChart(serialchart, name, coords)
    }

    /**
     * Sets whether the worksheet is protected. If `protect` is
     * `true`, the worksheet will be protected and the password will be
     * set to `password`. If it's `false`, the worksheet will
     * be unprotected and the password will be removed.
     *
     * @param protect  whether the worksheet should be protected
     * @param password the password to set if protect is `true`. ignored when
     * protect is `false`.
     * @throws WorkBookException never. This used to be thrown when unprotecting if the password
     * was incorrect.
     */
    @Throws(WorkBookException::class)
    fun setProtected(protect: Boolean, password: String) {
        val protector = sheet!!.protectionManager

        // we need to check if this password can be used to unprotect...
        // otherwise it is totally insecure...
        val oldpass = protector.password

        val pss = Password()
        pss.setPassword(password)
        val passcheck = pss.passwordHashString

        if (oldpass != null) {
            if (oldpass != passcheck && oldpass !== "0000") {
                throw WorkBookException("Incorrect Password Attempt to Unprotect Worksheet.",
                        WorkBookException.SHEETPROTECT_INCORRECT_PASSWORD)
            }
        }
        protector.protected = protect
        protector.password = if (protect) password else null
    }

    /**
     * Sets the password used to unlock the sheet when it is protected.
     *
     * @param password the clear text of the password to be applied or null to remove the
     * existing password
     */
    fun setProtectionPassword(password: String) {
        sheet!!.protectionManager.password = password
    }

    /**
     * Sets the password used to unlock the sheet when it is protected. This method
     * is useful in combination with [.getHashedProtectionPassword] to copy
     * the password from one worksheet to another.
     *
     * @param hash the hash of the protection password to be applied or null to
     * remove the existing password
     */
    fun setProtectionPasswordHashed(hash: String) {
        sheet!!.protectionManager.setPasswordHashed(hash)
    }

    /**
     * Checks whether the given password matches the protection password.
     *
     * @param guess the password to be checked against the stored hash
     * @return whether the given password matches the stored hash
     */
    fun checkProtectionPassword(guess: String): Boolean {
        return sheet!!.protectionManager.checkPassword(guess)
    }

    /**
     * Sets the worksheet enhanced protection option
     *
     * @param int protectionOption
     * @see WorkBookHandle.iprot options
     */
    fun setEnhancedProtection(protectionOption: Int, set: Boolean) {
        sheet!!.protectionManager.setProtected(protectionOption, set)
    }

    /**
     * returns true if the indicated Enhanced Protection Setting is turned on
     *
     * @param protectionOption
     * @return boolean true if the indicated Enhanced Protection Setting is turned
     * on
     * @see WorkBookHandle.iprot options
     */
    fun getEnhancedProtection(protectionOption: Int): Boolean {
        return sheet!!.protectionManager.getProtected(protectionOption)
    }

    /**
     * set this WorkSheet as the first visible tab on the left
     */
    fun setFirstVisibleTab() {
        sheet!!.workBook!!.setFirstVisibleSheet(sheet!!)
    }

    /**
     * returns the ColHandle for the column at index position the column index is
     * zero based ie: column A = 0
     *
     * @return ColHandle the Column
     */
    @Throws(ColumnNotFoundException::class)
    fun getCol(clnum: Int): ColHandle {
        var ci = sheet!!.getColInfo(clnum)
        val mycol: ColHandle
        if (ci == null || !ci.isSingleCol) {
            try {
                if (ci == null) {
                    ci = sheet!!.createColinfo(clnum, clnum)
                } else {
                    ci = sheet!!.createColinfo(clnum, clnum, ci)
                }
                mycol = ColHandle(ci, this)

            } catch (e: Exception) {
                throw ColumnNotFoundException("Unable to getCol for col number $clnum $e")
            }

        } else {
            mycol = ColHandle(ci, this) // usual case
        }
        return mycol
    }

    /**
     * adds the column (col1st, colLast) and returns the new ColHandle
     *
     * @param c1st
     * @param clast
     * @return
     */
    @Deprecated("use addCol(int)")
    fun addCol(c1st: Int, clast: Int): ColHandle {
        val ci = sheet!!.createColinfo(c1st, clast)
        return ColHandle(ci, this)
    }

    /**
     * adds the column (col1st, colLast) and returns the new ColHandle
     *
     * @param colNum, zero based number of the column
     * @return ColHandle
     */
    fun addCol(colNum: Int): ColHandle {
        val ci = sheet!!.createColinfo(colNum, colNum)
        return ColHandle(ci, this)
    }

    /**
     * returns the Column at the named position
     *
     * @return ColHandle the Column
     */
    @Throws(ColumnNotFoundException::class)
    fun getCol(name: String): ColHandle {
        return this.getCol(ExcelTools.getIntVal(name))
    }

    /**
     * returns the RowHandle for the row at index position
     *
     *
     * the row index is zero based ie: Excel row 1 = 0
     *
     * @param int row number to return
     * @return RowHandle a Row on this WorkSheet
     */
    @Throws(RowNotFoundException::class)
    fun getRow(t: Int): RowHandle {
        val x = sheet!!.getRowByNumber(t) ?: throw RowNotFoundException("Row " + t + " not found in :" + this.sheetName)
        return RowHandle(x, this)
    }

    /**
     * Returns whether a Cell exists in the WorkSheet.
     *
     * @param String the address of the Cell to check for
     * @return boolean whether the Cell exists
     */
    internal fun hasCell(addr: String): Boolean {
        try {
            this.getCell(addr)
            return true
        } catch (e: CellNotFoundException) {
            return false
        }

    }

    /**
     * Remove a Cell from this WorkSheet.
     *
     * @param CellHandle to remove
     */
    fun removeCell(celldel: CellHandle) {
        sheet!!.removeCell(celldel.cell!!)
    }

    /**
     * removes an Image from the Spreadsheet
     *
     *
     *
     *
     * Jan 22, 2010
     *
     * @param img
     */
    fun removeImage(img: ImageHandle) {
        sheet!!.removeImage(img)
        img.remove()
    }

    /**
     * Remove a Cell from this WorkSheet.
     *
     * @param String celladdr - the Address of the Cell to remove
     */
    fun removeCell(celladdr: String) {
        sheet!!.removeCell(celladdr.toUpperCase())
    }

    /**
     * Remove a Row and all associated Cells from this WorkSheet.
     *
     * @param int
     * rownum - the number of the row to remove not used public void
     * removeRow(int rownum) throws RowNotFoundException{
     * mysheet.removeRow(rownum); }
     */

    /**
     * Remove a Row and all associated Cells from this WorkSheet. Optionally shift
     * all rows below target row up one.
     *
     * @param int     rownum - the number of the row to remove
     * @param boolean shiftrows - true will shift all lower rows up one.
     */
    @Throws(RowNotFoundException::class)
    fun removeRow(rownum: Int, shiftrows: Boolean) {
        if (shiftrows)
            removeRow(rownum)
        else
            removeRow(rownum, WorkSheetHandle.ROW_DELETE_NO_REFERENCE_UPDATE)
    }

    /**
     * Remove all cells and formatting from a row within this WorkSheet. Has no
     * other affect upon the workbook
     *
     * @param int rownum - the number of the row contents to remove
     */
    @Throws(RowNotFoundException::class)
    fun removeRowContents(rownum: Int) {
        sheet!!.removeRowContents(rownum)
    }

    /**
     * Remove a Row and all associated Cells from this WorkSheet.
     *
     * @param int rownum - the number of the row to remove
     * @param int flag - controls whether row deletions updates references as well
     * ...
     */
    @Throws(RowNotFoundException::class)
    @JvmOverloads
    fun removeRow(rownum: Int, flag: Int = WorkSheetHandle.ROW_DELETE) {

        /* TODO: deal with merges! */
        sheet!!.removeRows(rownum, 1, true)

        // Delete chart series IF SERIES ARE ROW-BASED -- do before updateReferences
        val charts = this.sheet!!.charts
        for (i in charts.indices) {
            val sht = GenericPtg.qualifySheetname(this.sheetName)
            val c = charts[i] as Chart
            val seriesmap = c.seriesPtgs
            val ii = seriesmap.keys.iterator()
            while (ii.hasNext()) {
                val s = ii.next() as io.starter.formats.XLS.charts.Series
                val ptgs = seriesmap[s] as Array<Ptg>
                var pr: PtgRef
                var cursheet: String?
                var rc: IntArray?
                for (j in ptgs.indices) {
                    try {
                        pr = ptgs[j] as PtgRef
                        cursheet = pr.sheetName
                        rc = pr.intLocation
                        if (rc!![1] != rc[3] && sht!!.equals(cursheet!!, ignoreCase = true)) { // series are in rows, if existing
                            // series fall within deleted row
                            if (rc[0] == rownum - 1) {
                                c.removeSeries(j)
                                break // got it
                            }
                        } else
                            break // isn't row-based so split
                    } catch (e: Exception) {
                        continue // shouldn't happen!
                    }

                }
            }
            // also shift chart up if necessary [BugTracker 2858]
            val row = c.row0
            // only move images whose top is >= rnum
            val rnum = rownum + 1
            if (row > rnum) {
                val h = c.height
                // move down 1 row
                c.setRow(row - 1)
                c.height = h
            }
        }

        if (flag != WorkSheetHandle.ROW_DELETE_NO_REFERENCE_UPDATE)
            ReferenceTracker.updateReferences(rownum, -1, this.sheet!!, true)

        // Adjust image row so that height remains constant
        val rnum = rownum + 1
        val images = this.images
        for (i in images!!.indices) {
            val ih = images[i]
            val row = ih.row
            // only move images whose top is >= rnum
            if (row > rnum) {
                val h = ih.height
                // move down 1 row
                ih.row = row - 1
                ih.setHeight(h.toInt())
            }
        }
    }

    /**
     * Removes columns and all their associated cells from the sheet. This method
     * does not shift the subsequent columns left, for that use [.removeCols].
     *
     * @param first the zero-based index of the first column to be removed
     * @param count the number of columns to remove.
     */
    fun clearCols(first: Int, count: Int) {
        this.removeCols(first, count, false)
    }

    /**
     * Removes columns from the sheet and shifts the following columns left.
     *
     * @param first the zero-based index of the first column to be removed
     * @param count the number of columns to remove.
     */
    fun removeCols(first: Int, count: Int) {
        this.removeCols(first, count, true)
    }

    private fun removeCols(first: Int, count: Int, shift: Boolean) {
        if (first < 0)
            throw IllegalArgumentException("column index must be zero or greater")
        if (count < 1)
            throw IllegalArgumentException("count must be at least one")
        sheet!!.removeCols(first, count, shift)
    }

    /**
     * Removes a column and all associated cells from this sheet. This does not
     * shift subsequent columns.
     *
     * @param colstr the name of the column to remove
     */
    @Deprecated("Use {@link #clearCols} instead.")
    @Throws(ColumnNotFoundException::class)
    fun removeCol(colstr: String) {
        this.removeCols(ExcelTools.getIntVal(colstr), 1, false)
    }

    /**
     * Remove a column and all associated cells from this sheet. Optionally shift
     * all subsequent columns left to fill the gap.
     *
     * @param colstr    the name of the column to remove
     * @param shiftcols whether to shift subsequent columns
     */
    @Deprecated("Use {@link #removeCols} or {@link #clearCols} instead.")
    @Throws(ColumnNotFoundException::class)
    fun removeCol(colstr: String, shiftcols: Boolean) {
        this.removeCols(ExcelTools.getIntVal(colstr), 1, shiftcols)
    }

    /**
     * Returns a Named Range Handle if it exists in the specified scope.
     *
     *
     * This can be used to distinguish between multiple named ranges with the same
     * name but differing scopes
     *
     * @return NameHandle a Named range in the Worksheet that exists in the scope
     */
    @Throws(CellNotFoundException::class)
    fun getNamedRangeInScope(rangename: String): NameHandle {
        val nand = sheet!!.getScopedName(rangename) ?: throw CellNotFoundException(rangename)
        return NameHandle(nand, this.workBook)
    }

    /**
     * write this sheet as tabbed text output: <br></br>
     * All rows and all characters in each cell are saved. Columns of data are
     * separated by tab characters, and each row of data ends in a carriage return.
     * If a cell contains a comma, the cell contents are enclosed in double
     * quotation marks. All formatting, graphics, objects, and other worksheet
     * contents are lost. The euro symbol will be converted to a question mark. If
     * cells display formulas instead of formula values, the formulas are saved as
     * text.
     */
    @Throws(IOException::class)
    fun writeAsTabbedText(dest: OutputStream) {
        this.sheet!!.writeAsTabbedText(dest)
    }

    init {
        this.workBook = b
        this.sheet = sht
        this.mybook = sht.workBook
        // 20080624 KSC: add flag for shift formula rules upon row insertion/deletion
        val shiftRule = System.getProperties()["io.starter.OpenXLS.WorkSheetHandle.shiftInclusive"] as String
        if (shiftRule != null && shiftRule.equals("true", ignoreCase = true)) {
            sheet!!.setShiftRule(shiftRule.equals("true", ignoreCase = true))
        }
        // 20080917 KSC: set cache setting via system property [BugTracker 1862]
        if (System.getProperty("io.starter.OpenXLS.cacheCellHandles") != null)
            useCache = java.lang.Boolean.valueOf(System.getProperty("io.starter.OpenXLS.cacheCellHandles")).booleanValue()
    }

    /**
     * Set the Object value of the Cell at the given address.
     *
     * @param String Cell Address (ie: "D14")
     * @param Object new Cell Object value
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, `val`: Any) {
        val c = this.getCell(address)
        c.`val` = `val`
    }

    /**
     * Set the double value of the Cell at the given address
     *
     * @param String  Cell Address (ie: "D14")
     * @param double  new Cell double value
     * @param address
     * @param d
     * @throws io.starter.formats.XLS.CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, d: Double) {
        val c = this.getCell(address)
        c.setVal(d)
    }

    /**
     * Set the String value of the Cell at the given address
     *
     * @param String Cell Address (ie: "D14")
     * @param String new Cell String value
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, s: String) {
        val c = this.getCell(address)
        c.`val` = s
    }

    /**
     * Set the int value of the Cell at the given address
     *
     * @param String Cell Address (ie: "D14")
     * @param int    new Cell int value
     * @throws io.starter.formats.XLS.CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class, CellTypeMismatchException::class)
    fun setVal(address: String, i: Int) {
        val c = this.getCell(address)
        c.setVal(i)
    }

    /**
     * Get the Object value of a Cell.
     *
     *
     * Numeric Cell values will return as type Long, Integer, or Double. String Cell
     * values will return as type String.
     *
     * @return the value of the Cell as an Object.
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class)
    fun getVal(address: String): Any? {
        val c = this.getCell(address)
        return c.`val`
    }

    /**
     * Insert a row of Objects into the worksheet. Automatically shifts all rows
     * below the cell down one.
     *
     *
     * Method takes an array of Objects to insert into the rows.
     *
     *
     * Object array must match columns in number starting with column A.
     *
     *
     * For emptly cells, put a null Object reference in your array.
     *
     *
     * example: Object[] newCellHandles = { null, // col A "Hello", // col B
     * Integer.valueOf(120), // col C "=sum(A1+B1)", // col D null, // col E null,
     * // col F "World" // col G };
     *
     *
     * CellHandle ret = sheet.insertRow(newCellHandles, 1, true); if(ret !=null)
     * Logger.log("It worked");
     *
     * @param an      array of Objects to insert into the new row
     * @param rownum  the rownumber to insert
     * @param whether to shift down existing Cells
     */
    fun insertRow(row1: Int, data: Array<Any>): Array<CellHandle> {
        return insertRow(row1, data, true)
    }

    /**
     * Insert a blank row into the worksheet. Shift all rows below the cell down
     * one.
     *
     *
     * Adding new cells to non-existent rows will automatically create new rows in
     * the file, This method is only necessary to "move" existing cells by inserting
     * empty rows.
     *
     * @param rownum the rownumber to insert
     */
    fun insertRow(rownum: Int): Boolean {
        return insertRow(rownum, null as Row?, ROW_INSERT, true)
    }

    /**
     * Insert a blank row into the worksheet. Shift all rows below the cell down
     * one.
     *
     *
     * Adding new cells to non-existent rows will automatically create new rows in
     * the file, This method is only necessary to "move" existing cells by inserting
     * empty rows.
     *
     *
     * Same as insertRow(rownum) except with addition of flag
     *
     * @param rownum the rownumber to insert
     * @param flag   row insertion rule
     */
    fun insertRow(rownum: Int, flag: Int) {
        insertRow(rownum, null as Row?, flag, true)
    }

    /**
     * Insert a blank row into the worksheet. Shift all rows below the cell down
     * one.
     *
     *
     * This method differs from insertRow in that it can be used to repeatedly
     * insert rows at the same row index.
     *
     *
     * Adding new cells to non-existent rows will automatically create new rows in
     * the file,
     *
     *
     * After calling this method, setVal() can be used on the newly created cells to
     * update with new values.
     *
     * @param rownum  the rownumber to insert
     * @param whether to shift down existing Cells
     * @return whether the insert was successful
     */
    fun insertRowAt(rownum: Int, shiftrows: Boolean): Boolean {
        addedrows.remove(Integer.valueOf(rownum))
        return insertRow(rownum, sheet!!.getRowByNumber(rownum), rownum, shiftrows)
    }

    /**
     * Insert a blank row into the worksheet. Shift all rows below the cell down
     * one.
     *
     *
     * Adding new cells to non-existent rows will automatically create new rows in
     * the file,
     *
     *
     * After calling this method, setVal() can be used on the newly created cells to
     * update with new values.
     *
     * @param rownum  the rownumber to insert (NOTE: rownum is 0-based)
     * @param whether to shift down existing Cells
     * @return whether the insert was successful
     */
    fun insertRow(rownum: Int, shiftrows: Boolean): Boolean {
        var myr: Row? = null
        try {
            myr = sheet!!.getRowByNumber(rownum)
        } catch (e: Exception) {
        }

        if (myr != null || shiftrows) {
            return insertRow(rownum, myr, ROW_INSERT_MULTI, shiftrows)
        } else {
            // essentially a high performance row insert for the bottom of the workbook,
            // used frequently in streaming workbook insertion
            val newRow = sheet!!.insertRow(rownum, 0, ROW_INSERT_MULTI, shiftrows)
            return true
        }

    }

    /**
     * Insert a blank row into the worksheet. Shift all rows below the cell down
     * one.
     *
     *
     * Adding new cells to non-existent rows will automatically create new rows in
     * the file,
     *
     *
     * After calling this method, setVal() can be used on the newly created cells to
     * update with new values.
     *
     * @param rownum  the rownumber to insert
     * @param whether to shift down existing Cells
     */
    fun insertRow(rownum: Int, copyRow: RowHandle, flag: Int, shiftrows: Boolean): Boolean {
        return this.insertRow(rownum, copyRow.myRow, flag, shiftrows)
    }

    /**
     * Insert a row of Objects into the worksheet. Shift all rows below the cell
     * down one.
     *
     *
     * Method takes an array of Objects to insert into the rows.
     *
     *
     * Object array must match columns in number starting with column A.
     *
     *
     * For emptly cells, put a null Object reference in your array.
     *
     *
     * example: Object[] newCellHandles = { null, // col A "Hello", // col B
     * Integer.valueOf(120), // col C "=sum(A1+B1)", // col D null, // col E null,
     * // col F "World" // col G };
     *
     *
     * boolean okay = sheet.insertRow(newCellHandles, 1, true);
     * if(okay)Logger.log("It worked");
     *
     * @param an      array of Objects to insert into the new row
     * @param rownum  the rownumber to insert
     * @param whether to shift down existing Cells
     */
    fun insertRow(rownum: Int, data: Array<Any>, shiftrows: Boolean): Array<CellHandle> {
        val retc = arrayOfNulls<CellHandle>(data.size)
        try {
            insertRow(rownum, shiftrows)
            for (t in data.indices) {
                if (data[t] != null)
                    retc[t] = add(data[t], rownum, t)
            }
        } catch (ex: Exception) {
            throw WorkBookException(ex.toString(), WorkBookException.RUNTIME_ERROR)
        }

        return retc
    }

    /**
     * Insert a blank row into the worksheet. Shift all rows below the cell down
     * one.
     *
     *
     * Adding new cells to non-existent rows will automatically create new rows in
     * the file,
     *
     *
     * After calling this method, setVal() can be used on the newly created cells to
     * update with new values.
     *
     * @param rownum  the rownumber to insert (0-based)
     * @param copyrow the row to copy formats and formulas from
     * @param flag    determines handling tracking of inserted rows and only allow
     * insertion once
     * @param whether to shift down existing Cells
     */
    private fun insertRow(rownum: Int, copyRow: Row?, flag: Int, shiftrows: Boolean): Boolean {
        val offset = 1
        return shiftRow(rownum, copyRow, flag, shiftrows, offset)
    }

    /**
     * replacement method for delete row that handles references better
     *
     * @param rownum
     * @param flag
     * @param shiftrows
     * @return
     */
    internal fun deleteRow(rownum: Int, flag: Int, shiftrows: Boolean): Boolean {
        val offset = -1
        return shiftRow(rownum, null, flag, shiftrows, offset)
    }

    /**
     * insert/delete agnostic row copy/insert/delete and formula shifter
     *
     *
     * TODO: Better comments
     *
     * @param rownum
     * @param copyRow
     * @param flag
     * @param shiftrows
     * @param offset
     * @return
     */
    private fun shiftRow(rownum: Int, copyRow: Row?, flag: Int, shiftrows: Boolean, offset: Int): Boolean {
        var copyRow = copyRow
        var offset = offset

        // If the copyrow is null, such as an insert row on an empty row, create that
        // row, otherwise
        // we end up using different logic for row insertion, which makes no sense.
        if (copyRow == null) {
            // insert a blank
            this.add(null, "A" + rownum + 1)
            copyRow = sheet!!.getRowByNumber(rownum)

        }
        // handle tracking of inserted rows -- if flag is false rows can be inserted
        // multiple times at the same index
        if (flag == WorkSheetHandle.ROW_INSERT_ONCE) { // not inserted
            if (this.addedrows.contains(Integer.valueOf(rownum)))
                return false // can't add an existing row!
        }

        // sheetster ui means insert
        // 'on top of' row, shift down
        if (offset == 0)
            offset = 1

        // shiftrefs BEFORE inserting new row
        if (shiftrows && flag != WorkSheetHandle.ROW_INSERT_NO_REFERENCE_UPDATE) {
            var refUpdateStart = rownum
            // OpenXLS default behavior is to update one row too high. If we are using
            // ROW_INSERT, update per excel standard,
            // see TestInsertRows.testUpdateFormulaSettings() for testing
            if (flag == WorkSheetHandle.ROW_INSERT)
                refUpdateStart++
            ReferenceTracker.updateReferences(refUpdateStart, offset, this.sheet!!, true) // shift or expand/contract
            // ALL affected references
            // including named ranges
        }

        val firstcol = copyRow!!.colDimensions[0]
        val newRow = sheet!!.insertRow(rownum, firstcol, flag, shiftrows) // shifts rows down and inserts a new row,
        // also shifts shared formula refs (see note
        // below)

        // *************************************************************************************************************************************************************/
        // Named Range, Formula and AI references:
        // ALL references to the inserted row# (rownum) and rows beyond are shifted in
        // ReferenceTracker.updateReferences.
        // This method uses the ReferenceTracker collection for the specific sheet in
        // question to iterate through the stored references,
        // shifting them as the shifting rules allow.
        // In addition, all formulas in the copyrow will be duplicated in the newly
        // inserted row;
        // These formula references are shifted in ReferenceTracker.adjustFormulaRefs
        // (see below)
        // The only references that are NOT shifted in the schema described above are
        // SharedFormula references (specifically, PtgRefN & PtgAreaA),
        // which are NOT contained within the ReferenceTracker collection.
        // These references are shifted "by hand" in
        // ReferenceTracker.moveSharedFormulas, called upon
        // insertRow->->Boundsheet.shiftCellRow
        // *************************************************************************************************************************************************************/

        if (shiftrows)
            addedrows.add(Integer.valueOf(rownum))

        // TODO: Why so much logic in here, move this to Boundsheet?

        if (shiftrows && copyRow != null) {
            // Handle shifting reference rules for the newrow only and it's copyrow (NOTE:
            // copyrow may have been shifted via insertRow above although it's references
            // have not yet been shifted)
            val refMovementDiff = copyRow.rowNumber - rownum // number of rows to shift
            var refMovementRow = rownum // start row for shifting operation
            if (refMovementDiff < 0)
                refMovementRow += refMovementDiff // since copyrow < rownum, shifting should be done BEFORE insert row
            newRow!!.rowHeight = copyRow.rowHeight

            // Now iterate through all cells in the original row and copy formats and
            // formulas
            val copyRowCells = copyRow.cellArray
            val newmerges = Hashtable<String, CellRange>()
            var newmerge: CellRange? = null
            var newCellHandle: CellHandle? = null
            var aMul: Mulblank? = null // KSC: Mulblank handling
            var c: Short = -1 // ""
            val sheetname = GenericPtg.qualifySheetname(this.toString())
            for (i in copyRowCells.indices) {
                val copyRowCell = copyRowCells[i] as BiffRec
                if (copyRowCell.opcode == XLSConstants.MULBLANK) {
                    if (copyRowCell === aMul)
                        c++ // ref next blank in range - nec. for ixfe (FormatId) see below
                    else {
                        aMul = copyRowCell as Mulblank
                        c = aMul.colFirst.toShort()
                    }
                    aMul.setCurrentCell(c)
                }
                val copyCellHandle = CellHandle(copyRowCell, this.workBook)
                copyCellHandle.workSheetHandle = this
                // this.cellhandles.put(copyCellHandle.getCellAddress(), copyCellHandle);
                val colnum = copyCellHandle.colNum

                // insert an empty copy of the cell OR, if it's a formula cell, copy formula and
                // adjust it's cell references appropriate for new cell position
                if (copyRowCell.opcode == XLSRecord.FORMULA
                        && flag != WorkSheetHandle.ROW_INSERT_NO_REFERENCE_UPDATE
                        && flag != WorkSheetHandle.ROW_INSERT) {
                    try {
                        // copy copyrow's formula and then shift it's references relative to copycell's
                        // original row and references
                        newCellHandle = add(copyCellHandle.formulaHandle.formulaString, rownum, colnum)
                        // streaming parser uses fast cell adds which returns null, populate here
                        if (newCellHandle == null) {
                            try {
                                newCellHandle = this.getCell(rownum, colnum)
                            } catch (e: CellNotFoundException) {
                                // should be impossible }
                            }

                        }
                        // because the original formula has been shifted by one, unshift this sucker...
                        ReferenceTracker.adjustFormulaRefs(newCellHandle!!, refMovementRow, refMovementDiff * -1, true)
                        newCellHandle.formulaHandle.formulaRec!!.clearCachedValue()
                    } catch (e: Exception) {
                        if (DEBUGLEVEL > 0)
                            Logger.logWarn("WorkSheetHandle.shiftRow() could not adjust formula references in formula: "
                                    + copyCellHandle + " while inserting new row." + e.toString())
                    }

                } else {
                    newCellHandle = add(null, rownum, colnum)
                    // streaming parser uses fast cell adds which returns null, populate here
                    if (newCellHandle == null) {
                        try {
                            newCellHandle = this.getCell(rownum, colnum)
                        } catch (e: CellNotFoundException) {
                            // should be impossible }
                        }

                    }
                }
                newCellHandle!!.formatId = copyCellHandle.formatId

                // handle merged cells -- assemble the newmerges collection for below
                val oby = copyCellHandle.mergedCellRange
                if (oby != null) { // we have a merge
                    val fr = intArrayOf(rownum, oby.firstcellcol)
                    val lr = intArrayOf(rownum, oby.lastcellcol)
                    var newrng: String? = (sheetname + "!" + ExcelTools.formatLocation(fr) + ":"
                            + ExcelTools.formatLocation(lr))
                    newrng = GenericPtg.qualifySheetname(newrng)
                    if (DEBUGLEVEL > 10)
                        Logger.logInfo("WorksheetHandle.insertRow() created new Merge Range: " + newrng!!)
                    // check if we've already created...
                    if (newmerges[newrng] == null) {
                        newmerge = CellRange(newrng, this.workBook, true)
                        if (DEBUGLEVEL > 10)
                            Logger.logInfo("WorksheetHandle.insertRow() created new Merge Range: " + newrng!!)
                        newmerges[newmerge.toString()] = newmerge
                    }
                }

            }
            // now update the new merge ranges...
            val xl = newmerges.values
            if (xl != null) {
                val itx = xl.iterator()
                while (itx.hasNext()) {
                    itx.next().mergeCells(true)
                }
            }
        }
        // Handle Image Movement
        val images = sheet!!.images
        if (images != null) {
            for (i in images.indices) {
                val ih = images[i]
                val row = ih.row
                // only move images whose top is >= copyRow
                if (row >= rownum) {
                    // move down 1 row
                    val h = ih.height
                    ih.row = row + 1
                    ih.setHeight(h.toInt())
                }
            }
        }
        // Insert chart series IF SERIES ARE ROW-BASED
        val charts = this.sheet!!.charts
        for (i in charts.indices) {
            val c = charts[i] as Chart
            ReferenceTracker.insertChartSeries(c, GenericPtg.qualifySheetname(this.sheetName), rownum)
            // also shift charts down [BugTracker 2858]
            val row = c.row0
            // only move charts whose top is >= copyRow
            if (row >= rownum) {
                // move down 1 row
                val h = c.height
                c.setRow(row + offset)
                c.height = h
            }
        }
        return true
    }

    /**
     * Get a handle to all of the images in this worksheet
     *
     * @return
     */
    @Throws(ImageNotFoundException::class)
    fun getImage(name: String): ImageHandle {
        val idz = sheet!!.imageVect.indexOf(name)
        if (idz > 0)
            return sheet!!.imageVect[idz] as ImageHandle
        throw ImageNotFoundException("Could not find " + name + " in " + this.sheetName)
    }

    /**
     * write out all of the images in the Sheet to a directory
     *
     * @param imageoutput directory
     */
    fun extractImagesToDirectory(outdir: String) {
        val extracted = images

        // extract and output images
        for (tx in extracted!!.indices) {
            var n = extracted[tx].name
            if (n == "")
                n = "image" + extracted[tx].msodrawing!!.imageIndex
            val imgname = n + "." + extracted[tx].type
            if (DEBUGLEVEL > 0)
                Logger.logInfo("Successfully extracted: $outdir$imgname")
            try {
                val outimg = FileOutputStream(outdir + imgname)
                extracted[tx].write(outimg)
                outimg.flush()
                outimg.close()
            } catch (ex: Exception) {
                Logger.logErr("Could not extract images from: $this")
            }

        }
    }

    /**
     * retrieves all charts for this sheet and writes them (in SVG form) to outpdir
     * <br></br>
     * Filename is in form of: <sheetname>_Chart<#>.svg
     *
     * @param outdir String output folder
    </sheetname> */
    fun extractChartToDirectory(outdir: String) {
        val charts = this.sheet!!.charts as ArrayList<*>
        val sheetname = this.sheetName
        for (i in charts.indices) {
            val ch = ChartHandle(charts[i] as Chart, this.workBook)
            val fname = sheetname + "_Chart" + ch.id + ".svg"
            try {
                val chartout = FileOutputStream(outdir + fname)
                chartout.write(ch.getSVG(1.0).toByteArray()) // scaled as necessary in XSL
            } catch (ex: Exception) {
                Logger.logErr("extractChartToDirectory: Could not extract charts from: $this:$ex")
            }

        }
    }

    /**
     * insert an image into this worksheet
     *
     * @param im -- the ImageHandle to insert
     * @see ImageHandle
     */
    fun insertImage(im: ImageHandle) {
        this.sheet!!.insertImage(im)
    }

    /**
     * Inserts empty columns and shifts the following columns to the right. This
     * method is used to shift existing columns right to make room for a new column.
     * To create a column in the file so its size or style can be set just add a
     * blank cell to its first row with [.add].
     *
     * @param first the zero-based index of the first column to insert
     * @param count the number of columns to insert
     */
    fun insertCols(first: Int, count: Int) {
        if (first < 0)
            throw IllegalArgumentException("column index must be zero or greater")
        if (count < 1)
            throw IllegalArgumentException("count must be at least one")
        sheet!!.insertCols(first, count)
    }

    /**
     * Inserts an empty column and shifts the following columns to the right.
     *
     * @param colnum the zero-based index of the column to be inserted
     */
    @Deprecated("Use {@link #insertCols} instead.")
    fun insertCol(colnum: Int) {
        this.insertCols(colnum, 1)
    }

    /**
     * Inserts an empty column and shifts the following columns to the right.
     *
     * @param colnum the address of the column to be inserted
     */
    @Deprecated("Use {@link #insertCols} instead.")
    fun insertCol(colnum: String) {
        this.insertCols(ExcelTools.getIntVal(colnum), 1)
    }

    /**
     * When adding a new Cell to the sheet, OpenXLS can automatically copy the
     * formatting from the Cell directly above the inserted Cell.
     *
     *
     * ie: if set to true, newly added Cell D19 would take its formatting from Cell
     * D18.
     *
     *
     * Default is false
     *
     *
     *
     *
     *
     *
     * boolean copy the formats from the prior Cell
     */
    fun setCopyFormatsFromPriorWhenAdding(f: Boolean) {
        sheet!!.setCopyPriorCellFormats(f)
    }

    /**
     * Add a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on type-compatibility of the
     * value.
     *
     *
     * Further, this method allows passing in a format id
     *
     *
     * In other words, if the Object cannot be converted safely to a Numeric Object
     * type, then it is treated as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     *
     * If a validation record for the cell exists the validation is checked for a
     * correct value, if the value does not pass the validation a
     * ValidationException will be thrown
     *
     * @param obj the value of the new Cell
     * @param int row the row of the new Cell
     * @param int col the column of the new Cell
     * @param int the format id to apply to this cell
     */
    @Throws(ValidationException::class)
    @JvmOverloads
    fun addValidated(obj: Any, row: Int, col: Int, formatId: Int = this.workBook!!.workBook!!.defaultIxfe): Array<CellHandle> {
        val rc = intArrayOf(row, col)
        val vh = this.getValidationHandle(ExcelTools.formatLocation(rc))
        vh?.isValid(obj)
        return this.addValidated(obj, ExcelTools.formatLocation(rc))
    }

    /**
     * Add a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on type-compatibility of the
     * value.
     *
     *
     * Further, this method allows passing in a format id
     *
     *
     * In other words, if the Object cannot be converted safely to a Numeric Object
     * type, then it is treated as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     * @param obj the value of the new Cell
     * @param int row the row of the new Cell
     * @param int col the column of the new Cell
     * @param int the format id to apply to this cell
     */
    @JvmOverloads
    fun add(obj: Any?, row: Int, col: Int, formatId: Int = this.workBook!!.workBook!!.defaultIxfe): CellHandle? {
        val rc = intArrayOf(row, col)

        if (obj is java.util.Date) {
            val address = ExcelTools.formatLocation(rc)
            this.add(obj as java.util.Date?, address, null)
        } else {
            val reca = sheet!!.addValue(obj, rc, formatId)

            if (DEBUGLEVEL > 1)
                if (reca != null)
                    Logger.logInfo("WorkSheetHandle.add() $reca Successfully Added.")
                else
                    return null
        }

        return if (!sheet!!.fastCellAdds) {
            try {
                val c = this.getCell(row, col)
                if (this.workBook!!.formulaCalculationMode != workBook!!.CALCULATE_EXPLICIT)
                    c.clearAffectedCells() // blow out cache
                c
            } catch (e: CellNotFoundException) {
                Logger.logInfo("Adding Cell to row failed row:$row col: $col failed.")
                null
            }

        } else {
            null
        }
    }

    /**
     * Fast-adds a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on type-compatibility of the
     * value.
     *
     *
     * Further, this method allows passing in a format id
     *
     *
     * In other words, if the Object cannot be converted safely to a Numeric Object
     * type, then it is treated as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     * @param obj the value of the new Cell
     * @param int row the row of the new Cell
     * @param int col the column of the new Cell
     * @param int the format id to apply to this cell
     */
    fun fastAdd(obj: Any, row: Int, col: Int, formatId: Int) {
        var formatId = formatId
        val rc = intArrayOf(row, col)

        // use default format
        if (formatId == -1)
            formatId = 0

        if (obj is java.util.Date) {
            val address = ExcelTools.formatLocation(rc)
            this.add(obj, address, null)
        } else {
            val reca = sheet!!.addValue(obj, rc, formatId)
            if (this.workBook!!.formulaCalculationMode != workBook!!.CALCULATE_EXPLICIT) {
                val rt = this.workBook!!.workBook!!.refTracker
                rt!!.clearAffectedFormulaCells(reca)
            }
            if (DEBUGLEVEL > 1)
                if (reca != null)
                    Logger.logInfo("WorkSheetHandle.add() $reca Successfully Added.")
        }
    }

    /**
     * Add a Cell with the specified value to a WorkSheet, optionally attempting to
     * convert numeric values to appropriate number cells with appropriate number
     * formatting applied.
     *
     *
     * This would allow a value entered such as "$1.00" to be converted to a numeric
     * 1.00d value with a $0.00 format pattern applied.
     *
     *
     * Note that there is overhead to this method, as the value added needs to be
     * parsed, for higher performance do not use the autodetectNumberAndPattern =
     * true setting, and instead pass numeric values and format patterns.
     *
     *
     * If the value is non numeric, then it is simply added.
     *
     * @param obj             the value of the new Cell
     * @param address         the address of the new Cell
     * @param autoDetectValue whether to attempt to store as a number with a format pattern.
     */
    fun add(obj: Any?, address: String, autoDetectValue: Boolean): CellHandle? {
        if (obj == null)
            sheet!!.addValue(obj, address) // to add a blank cell
        else if (obj is Date)
            this.add(obj as Date?, address, null)
        else
            sheet!!.addValue(obj, address, autoDetectValue)//

        // fast-adds optimzation -- do not return a CellHandle
        if (this.sheet!!.fastCellAdds)
            return null

        try {
            return this.getCell(address)
        } catch (e: CellNotFoundException) {
            Logger.logInfo("Adding Cell: $address failed")
            return null
        }

    }

    /**
     * Add a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on type-compatibility of the
     * value.
     *
     *
     * In other words, if the Object cannot be converted safely to a Numeric Object
     * type, then it is treated as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     * @param obj     the value of the new Cell
     * @param address the address of the new Cell
     */
    fun add(obj: Any?, address: String): CellHandle? {
        if (obj == null)
            sheet!!.addValue(obj, address) // to add a blank cell
        else if (obj is Date)
            this.add(obj as Date?, address, null)
        else
            sheet!!.addValue(obj, address)//

        // fast-adds optimzation -- do not return a CellHandle
        return if (!sheet!!.fastCellAdds) {
            try {
                val c = this.getCell(address)
                if (this.workBook!!.formulaCalculationMode != workBook!!.CALCULATE_EXPLICIT)
                    c.clearAffectedCells() // blow out cache
                c
            } catch (e: CellNotFoundException) {
                Logger.logInfo("Adding Cell: $address failed")
                null
            }

        } else {
            null
        }
    }

    /**
     * Add a Cell with the specified value to a WorkSheet.
     *
     *
     * This method determines the Cell type based on type-compatibility of the
     * value.
     *
     *
     * In other words, if the Object cannot be converted safely to a Numeric Object
     * type, then it is treated as a String and a new String value is added to the
     * WorkSheet at the Cell address specified.
     *
     *
     *
     *
     * If a validation record for the cell exists the validation is checked for a
     * correct value, if the value does not pass the validation a
     * ValidationException will be thrown
     *
     * @param obj     the value of the new Cell
     * @param address the address of the new Cell
     */
    @Throws(ValidationException::class)
    fun addValidated(obj: Any, address: String): Array<CellHandle> {
        val vh = this.getValidationHandle(address)
        vh?.isValid(obj)
        val ch = this.add(obj, address)
        val cxrs = ch!!.calculateAffectedCellsOnSheet()

        // return the cellhandles
        val cxrx = arrayOfNulls<CellHandle>(cxrs.size + 1)
        cxrx[0] = ch
        for (t in 1 until cxrx.size) {
            cxrx[t] = cxrs[t - 1] as CellHandle
            cxrx[t].workSheetHandle = this
            // this.cellhandles.put(cxrx[t].getCellAddress(), cxrx[t]);
        }

        return cxrx
    }

    /**
     * Add a java.sql.Timestamp Cell to a WorkSheet.
     *
     *
     * Will create a default format of:
     *
     *
     * "m/d/yyyy h:mm:ss"
     *
     *
     * if none is specified.
     *
     * @param dt         the value of the new java.sql.Timestamp Cell
     * @param address    the address of the new java.sql.Date Cell
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    fun add(dt: Timestamp, address: String, fmt: String?): CellHandle? {
        var fmt = fmt
        if (fmt == null)
            fmt = "m/d/yyyy h:mm:ss"
        val dx = Date(dt.time)
        return this.add(dx, address, fmt)
    }

    /**
     * Add a java.sql.Timestamp Cell to a WorkSheet.
     *
     *
     * Will create a default format of:
     *
     *
     * "m/d/yyyy h:mm:ss"
     *
     *
     * if none is specified.
     *
     *
     * If a validation record for the cell exists the validation is checked for a
     * correct value, if the value does not pass the validation a
     * ValidationException will be thrown
     *
     * @param dt         the value of the new java.sql.Timestamp Cell
     * @param address    the address of the new java.sql.Date Cell
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    @Throws(ValidationException::class)
    fun addValidated(dt: Timestamp, address: String, fmt: String?): Array<CellHandle> {
        var fmt = fmt
        if (fmt == null)
            fmt = "m/d/yyyy h:mm:ss"
        val dx = Date(dt.time)
        return this.addValidated(dx, address, fmt)
    }

    /**
     * Add a java.sql.Date Cell to a WorkSheet.
     *
     *
     * You must specify a formatting pattern for the new date, or null for the
     * default ("m/d/yy h:mm".)
     *
     *
     * valid date format patterns "m/d/y" "d-mmm-yy" "d-mmm" "mmm-yy" "h:mm AM/PM"
     * "h:mm:ss AM/PM" "h:mm" "h:mm:ss" "m/d/yy h:mm" "mm:ss" "[h]:mm:ss" "mm:ss.0"
     *
     * @param dt         the value of the new java.sql.Date Cell
     * @param row        to add the date
     * @param col        to add the date
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    fun add(dt: java.util.Date, address: String, fmt: String?): CellHandle? {
        val rc = ExcelTools.getRowColFromString(address)
        return add(dt, rc[0], rc[1], fmt)
    }

    /**
     * Add a java.sql.Date Cell to a WorkSheet.
     *
     *
     * You must specify a formatting pattern for the new date, or null for the
     * default ("m/d/yy h:mm".)
     *
     *
     * valid date format patterns "m/d/y" "d-mmm-yy" "d-mmm" "mmm-yy" "h:mm AM/PM"
     * "h:mm:ss AM/PM" "h:mm" "h:mm:ss" "m/d/yy h:mm" "mm:ss" "[h]:mm:ss" "mm:ss.0"
     *
     * @param dt         the value of the new java.sql.Date Cell
     * @param row        to add the date
     * @param col        to add the date
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    @Throws(ValidationException::class)
    fun addValidated(dt: java.util.Date, address: String, fmt: String): Array<CellHandle> {
        val rc = ExcelTools.getRowColFromString(address)
        return addValidated(dt, rc[0], rc[1], fmt)
    }

    /**
     * Add a java.sql.Date Cell to a WorkSheet.
     *
     *
     * You must specify a formatting pattern for the new date, or null for the
     * default ("m/d/yy h:mm".)
     *
     *
     * valid date format patterns "m/d/y" "d-mmm-yy" "d-mmm" "mmm-yy" "h:mm AM/PM"
     * "h:mm:ss AM/PM" "h:mm" "h:mm:ss" "m/d/yy h:mm" "mm:ss" "[h]:mm:ss" "mm:ss.0"
     *
     * @param dt         the value of the new java.sql.Date Cell
     * @param address    the address of the new java.sql.Date Cell
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    fun add(dt: java.util.Date, row: Int, col: Int, fmt: String?): CellHandle? {
        var fmt = fmt
        val x = DateConverter.getXLSDateVal(dt, this.mybook!!.dateFormat)
        var thisCell = this.add(x, row, col)

        // first handle fast adds
        if (thisCell == null && this.sheet!!.fastCellAdds) {
            try {
                thisCell = getCell(row, col)
            } catch (exp: CellNotFoundException) {
                Logger.logWarn("adding date to WorkSheet failed: " + this.sheetName + ":" + row + ":" + col)
                return null
            }

        }

        // 20060419 KSC: Use format from cell, if any
        if (fmt == null) {
            fmt = thisCell!!.formatPattern
            if (fmt == null || fmt == "General")
                fmt = "m/d/yy h:mm"
        }
        if (dateFormats!![fmt] == null) {
            val fh = thisCell!!.formatHandle
            fh!!.formatPattern = fmt
            dateFormats!![fmt] = Integer.valueOf(thisCell.formatId)!!
        } else {
            val `in` = dateFormats!![fmt]
            thisCell!!.formatId = `in`!!.toInt()
        }
        // Logger.logInfo("Date added: " + thisCell.getFormattedStringVal());
        return thisCell
    }

    /**
     * Add a java.sql.Date Cell to a WorkSheet.
     *
     *
     * You must specify a formatting pattern for the new date, or null for the
     * default ("m/d/yy h:mm".)
     *
     *
     * valid date format patterns "m/d/y" "d-mmm-yy" "d-mmm" "mmm-yy" "h:mm AM/PM"
     * "h:mm:ss AM/PM" "h:mm" "h:mm:ss" "m/d/yy h:mm" "mm:ss" "[h]:mm:ss" "mm:ss.0"
     *
     *
     * If a validation record for the cell exists the validation is checked for a
     * correct value, if the value does not pass the validation a
     * ValidationException will be thrown
     *
     * @param dt         the value of the new java.sql.Date Cell
     * @param address    the address of the new java.sql.Date Cell
     * @param formatting pattern the address of the new java.sql.Date Cell
     */
    @Throws(ValidationException::class)
    fun addValidated(dt: java.util.Date, row: Int, col: Int, fmt: String): Array<CellHandle> {
        val rc = intArrayOf(row, col)
        val vh = this.getValidationHandle(ExcelTools.formatLocation(rc))
        vh?.isValid(dt)
        val ch = this.add(dt, row, col, fmt)
        val cxrs = ch!!.calculateAffectedCellsOnSheet()

        // FIXME: Use List.toArray instead of for loop
        // return the cellhandles
        val cxrx = arrayOfNulls<CellHandle>(cxrs.size + 1)
        cxrx[0] = ch
        for (t in 1 until cxrx.size) {
            cxrx[t] = cxrs[t - 1] as CellHandle
        }

        return cxrx

    }

    /**
     * Remove this WorkSheet from the WorkBook
     *
     *
     * NOTE: will throw a WorkBookException if the last sheet is removed. This
     * results in an invalid output file.
     */
    fun remove() {
        mybook!!.removeWorkSheet(this.sheet!!)
        workBook!!.sheethandles!!.remove(this.sheetName)
    }

    /**
     * Create a CellRange object from an OpenXLS string range passed in such as
     * "A1:F6"
     *
     * @param rangeName "A1:F6"
     * @return
     * @throws CellNotFoundException if the range cannot be created
     */
    @Throws(CellNotFoundException::class)
    fun getCellRange(rangeName: String): CellRange {
        return CellRange("$sheetName!$rangeName", workBook)
    }

    /**
     * Returns a FormulaHandle for working with the ranges of a formula on a
     * WorkSheet.
     *
     * @param addr the address of the Cell
     * @throws FormulaNotFoundException is thrown if there is no existing formula at the specified
     * address.
     */
    @Throws(FormulaNotFoundException::class, CellNotFoundException::class)
    fun getFormula(addr: String): FormulaHandle {
        val c = this.getCell(addr)
        return c.formulaHandle
    }

    /**
     * Returns a CellHandle for working with the value of a Cell on a WorkSheet.
     *
     * @param addr the address of the Cell
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class)
    fun getCell(addr: String): CellHandle {
        var addr = addr
        var ret: CellHandle? = null //

        val c = sheet!!.getCell(addr.toUpperCase())
        if (c == null) {
            var sn = ""
            try {
                sn = this.sheetName
                sn += "!"
            } catch (e: Exception) {
            }

            if (addr == null)
                addr = "undefined cell address"
            throw CellNotFoundException(sn + addr)
        }
        ret = CellHandle(c, this.workBook)
        ret.workSheetHandle = this
        return ret
    }

    /**
     * Returns a CellHandle for working with the value of a Cell on a WorkSheet.
     *
     * @param int     Row the integer row of the Cell
     * @param int     Col the integer col of the Cell
     * @param boolean whether to cache or return a new CellHandle each call
     * @throws CellNotFoundException is thrown if there is no existing Cell at the specified
     * address.
     */
    @Throws(CellNotFoundException::class)
    @JvmOverloads
    fun getCell(row: Int, col: Int, cache: Boolean = false): CellHandle {
        var ret: CellHandle? = null
        if (cache) {
            val rc = intArrayOf(row, col)
            val address = ExcelTools.formatLocation(rc)
            // ret = (CellHandle)cellhandles.get(address);
            if (ret != null)
            // caching!
                return ret
            else {
                ret = CellHandle(this.sheet!!.getCell(row, col), this.workBook)
                ret.workSheetHandle = this
                // cellhandles.put(address,ret);
                return ret
            }
        }
        ret = CellHandle(this.sheet!!.getCell(row, col), this.workBook)
        ret.workSheetHandle = this
        return ret
    }

    /**
     * Move a cell on this WorkSheet.
     *
     * @param CellHandle c - the cell to be moved
     * @param String     celladdr - the destination address of the cell
     */
    @Throws(CellPositionConflictException::class)
    fun moveCell(c: CellHandle, addr: String) {
        this.sheet!!.moveCell(c.cellAddress, addr)
        // c.moveTo(addr); < redundant call to above. 070104 -jm
    }

    /**
     * Sets the print area for the worksheet
     *
     *
     *
     *
     * sets the printarea as a CellRange
     *
     * @param printarea
     */
    fun setPrintArea(printarea: CellRange) {
        sheet!!.printArea = printarea.getRange()
    }

    /**
     * Returns the name of this Sheet.
     *
     * @see java.lang.Object.toString
     */
    override fun toString(): String {
        return sheet!!.toString()
    }

    /**
     * FOR internal Use Only!
     *
     * @return Returns the low-level sheet record.
     */
    fun getMysheet(): Boundsheet? {
        return sheet
    }

    /**
     * Calculates all formulas that reference the cell address passed in.
     *
     *
     * Please note that these cells have already been calculated, so in order to get
     * their values without re-calculating them Extentech suggests setting the book
     * level non-calculation flag, ie
     * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT) or
     * FormulaHandle.getCachedVal()
     *
     * @return List of of calculated cells
     */
    fun calculateAffectedCells(CellAddress: String): List<*>? {
        var c: CellHandle? = null
        try {
            c = this.getCell(CellAddress)
        } catch (e: CellNotFoundException) {
            return null
        }

        return c.calculateAffectedCells()
    }

    /**
     * Get whether there are freeze panes in the output sheet.
     *
     * @return boolean whether there are freeze panes
     */
    fun hasFrozenPanes(): Boolean {
        return this.sheet!!.window2!!.freezePanes
    }

    /**
     * Set whether there are freeze panes in the output sheet.
     *
     * @param boolean whether there are freeze panes
     */
    fun setHasFrozenPanes(b: Boolean) {
        this.sheet!!.window2!!.freezePanes = b
        if (!b && this.sheet!!.pane != null) {
            this.sheet!!.removePane() // remove pane rec if unfreezing ... can also convert to plain splits, but a bit
            // more complicated ...
        }
    }

    /**
     * freezes the rows starting at the specified row and creating a scrollable
     * sheet below this row
     *
     * @param row the row to start the freeze
     */
    fun freezeRow(row: Int) {
        if (this.sheet!!.pane == null)
            this.sheet!!.pane = null // will add new
        this.sheet!!.pane!!.setFrozenRow(row)
    }

    /**
     * freezes the cols starting at the specified column and creating a scrollable
     * sheet to the right
     *
     * @param col the col to start the freeze
     */
    fun freezeCol(col: Int) {
        if (this.sheet!!.pane == null)
            this.sheet!!.pane = null // will add new
        this.sheet!!.pane!!.setFrozenColumn(col)
    }

    /**
     * splits the worksheet at column col for nCols
     *
     *
     * Note: unfreezes panes if frozen
     *
     * @param col      col start col to split
     * @param splitpos position of the horizontal split
     */
    fun splitCol(col: Int, splitpos: Int) {
        if (this.sheet!!.pane == null)
            this.sheet!!.pane = null // will add new
        this.sheet!!.pane!!.setSplitColumn(col, splitpos)
    }

    /**
     * splits the worksheet at row for nRows
     *
     *
     * Note: unfreezes panes if frozen
     *
     * @param row      start row to split
     * @param splitpos position of the vertical split
     */
    fun splitRow(row: Int, splitpos: Int) {
        if (this.sheet!!.pane == null)
            this.sheet!!.pane = null // will add new
        this.sheet!!.pane!!.setSplitRow(row, splitpos)
    }

    /**
     * Create a Conditional Format handle for a cell/range
     *
     * @param cellAddress     without sheetname. Can also be a range, such as A1:B5
     * @param qualifier       = maps to CONDITION_* bytes in ConditionalFormatHandle
     * @param value1          = the error message
     * @param value2          = the error title
     * @param format          = the initial format string to use with the condition
     * @param firstCondition  = formula string
     * @param secondCondition = 2nd formula string (optional)
     * @return
     */
    fun createConditionalFormatHandle(cellAddress: String?, operator: String, value1: String,
                                      value2: String, format: String, firstCondition: String, secondCondition: String): ConditionalFormatHandle {
        var cellAddress = cellAddress

        if (cellAddress != null && cellAddress.indexOf("!") == -1)
            cellAddress = this.sheetName + "!" + cellAddress
        val cfm = this.sheet!!.createCondfmt(cellAddress, this.workBook)

        val cfr = this.sheet!!.createCf(cfm)
        cfr.setOperator(operator)

        // only place this is done..
        cfr.setCondition1(value1)
        cfr.setCondition2(value2)

        Cf.setStylePropsFromString(format, cfr)

// done above in createCondfmt this.mysheet.addConditionalFormat(cfm);
        return ConditionalFormatHandle(cfm, this)
    }

    /**
     * Creates a new annotation (Note or Comment) to the worksheeet, attached to a
     * specific cell
     *
     * @param address -- address to attach
     * @param txt     -- text of note
     * @param author  -- name of author
     * @return NoteHandle - handle which allows access to the Note object
     * @see CommentHandle
     */
    fun createNote(address: String, txt: String, author: String): CommentHandle {
        val n = this.getMysheet()!!.createNote(address, txt, author)
        return CommentHandle(n)
    }

    /**
     * Creates a new annotation (Note or Comment) to the worksheeet, attached to a
     * specific cell <br></br>
     * The note or comment is a Unicode string, thus it can contain formatting
     * information
     *
     * @param address -- address to attach
     * @param txt     -- Unicode string of note with Formatting
     * @param author  -- name of author
     * @return NoteHandle - handle which allows access to the Note object
     * @see CommentHandle
     */
    fun createNote(address: String, txt: Unicodestring, author: String): CommentHandle {
        val n = this.getMysheet()!!.createNote(address, txt, author)
        return CommentHandle(n)
    }

    /**
     * creates a new, blank PivotTable and adds it to the worksheet.
     *
     * @param name  pivot table name
     * @param range source range for the pivot table. If no sheet is specified, the
     * current sheet will be used.
     * @param sId   Stream or cachid Id -- links back to SxStream set of records
     * @return PivotTableHandle
     */
    fun createPivotTable(name: String, range: String, sId: Int): PivotTableHandle {
        val sx = getMysheet()!!.addPivotTable(range, this.workBook, sId, name)
        val pth = PivotTableHandle(sx, this.workBook)
        pth.setSourceDataRange(range)
        return pth
    }

    /**
     * Get a validation handle for the cell address passed in. If the validation is
     * for a range, the handle returned will modify the entire range, not just the
     * cell address passed in.
     *
     *
     * Returns null if a validation does not exist at the specified location
     *
     * @param cell address String
     */
    fun getValidationHandle(cellAddress: String): ValidationHandle? {
        if (this.sheet!!.dvalRec != null) {
            val d = this.sheet!!.dvalRec!!.getDv(cellAddress) ?: return null
            return ValidationHandle(d)
        }
        return null
    }

    /**
     * Create a validation handle for a cell/range
     *
     * @param cellAddress     without sheetname. Can also be a range, such as A1:B5
     * @param valueType       = maps to VALUE_* bytes in ValidationHandle
     * @param condition       = maps to CONDITION_* bytes in ValidationHandle
     * @param errorBoxText    = the error message
     * @param errorBoxTitle   = the error title
     * @param promptBoxText   = the prompt (hover) message
     * @param promptBoxTitle  = the prompt (hover) title
     * @param firstCondition  = formula string, seeValidationHandle.setFirstCondition
     * @param secondCondition = seeValidationHandle.setSecondCondition, this can be left null
     * for validations that do not require a second argument.
     * @return
     */
    fun createValidationHandle(cellAddress: String, valueType: Byte, condition: Byte,
                               errorBoxText: String, errorBoxTitle: String, promptBoxText: String, promptBoxTitle: String,
                               firstCondition: String?, secondCondition: String?): ValidationHandle {
        val d = this.sheet!!.createDv(cellAddress)
        d.valType = valueType
        /*
         * // KSC: APPARENTLY NOT NEEDED if
         * (valueType==ValidationHandle.VALUE_USER_DEFINED_LIST) { // ensure Mso
         * Drop-downs are defined int objId =
         * this.mysheet.insertDropDownBox(d.getColNumber()); //TODO: verify that drop
         * down lists are SHARED ****
         * this.mysheet.getDvalRec().setObjectIdentifier(objId); }
         */
        d.typeOperator = condition
        d.errorBoxText = errorBoxText
        d.errorBoxTitle = errorBoxTitle
        d.promptBoxText = promptBoxText
        d.promptBoxTitle = promptBoxTitle
        if (firstCondition != null)
            d.firstCond = firstCondition
        if (secondCondition != null)
            d.secondCond = secondCondition
        return ValidationHandle(d)
    }

    /**
     * return true if sheet contains data validations
     *
     * @return boolean
     */
    fun hasDataValidations(): Boolean {
        return sheet!!.dvRecs != null
    }

    /**
     * Adds a new AutoFilter for the specified column (0-based) in this sheet <br></br>
     * returns a handle to the new AutoFilter
     *
     * @param int column - column number to add an AutoFilter to
     * @return AutoFilterHandle
     */
    fun addAutoFilter(column: Int): AutoFilterHandle {
        val af = sheet!!.addAutoFilter(column)
        return AutoFilterHandle(af)
    }

    /**
     * Removes all AutoFilters from this sheet <br></br>
     * As a consequence, all previously hidden rows are shown or unhidden
     */
    fun removeAutoFilters() {
        sheet!!.removeAutoFilter()
    }

    /**
     * Updates the Row filter (hidden status) for each row on the sheet by
     * evaluating all AutoFilter conditions
     *
     *
     * NOTE: This method **must** be called after Autofilter updates or additions
     * in order to see the results of the AutoFilter(s)
     *
     *
     * NOTE: this evaluation is NOT done automatically due to performance
     * considerations, and is designed to be called after all additions and updating
     * is completed (as evaluation may be time-consuming)
     */
    fun evaluateAutoFilters() {
        sheet!!.evaluateAutoFilters()
    }

    /**
     * clear out object references in prep for closing workbook
     */
    fun close() {
        if (sheet != null)
            sheet!!.close()
        addedrows.clear()
        addedrows = CompatibleVector()
        sheet = null
        mybook = null
        workBook = null
        dateFormats!!.clear()
        dateFormats = null
    }

    fun getBoundsheet(): Boundsheet? {
        return this.sheet
    }

    /**
     * Imports the given CSV data into this worksheet. All rows in the input will be
     * inserted sequentially before any rows which already exist in this worksheet.
     *
     *
     * To change the value delimiter set the system property
     * "`io.starter.OpenXLS.csvdelimiter`".
     */
    @Throws(IOException::class)
    fun readCSV(input: BufferedReader) {
        var rws = 0
        val field_delimiter = System.getProperty("io.starter.OpenXLS.csvdelimiter", ",")

        var thisLine = ""
        while ((thisLine = input.readLine()) != null) { // while loop begins here
            val vals = StringTool.getTokensUsingDelim(thisLine, field_delimiter)
            val data = arrayOfNulls<Any>(vals.size)
            for (t in vals.indices) {
                vals[t] = StringTool.strip(vals[t], '"')
                try {
                    val i = Integer.parseInt(vals[t])
                    data[t] = Integer.valueOf(i)

                    val d = java.lang.Double.parseDouble(vals[t])
                    data[t] = d

                } catch (ax: NumberFormatException) {
                    // it's a string!
                    data[t] = vals[t]
                }

            }

            insertRow(rws++, data, true)
        }
    }

    companion object {

        // insert handling flags
        /**
         * Insert row multiple times allowed, also copies formulas to inserted row
         */
        val ROW_INSERT_MULTI = 0
        /**
         * Excel standard row insertion behavior
         */
        val ROW_INSERT = 3
        /**
         * Insert row one time, multiple calls ignored
         */
        val ROW_INSERT_ONCE = 1
        /**
         * Insert row but do not update any cell references affected by insert
         */
        val ROW_INSERT_NO_REFERENCE_UPDATE = 2

        // 20080619 KSC: Add flag constants for Delete Row
        val ROW_DELETE = 1
        val ROW_DELETE_NO_REFERENCE_UPDATE = 2

        /**
         * enhanced protection settings: Edit Object
         */
        val ALLOWOBJECTS = FeatHeadr.ALLOWOBJECTS
        /**
         * enhanced protection settings: Edit scenario
         */
        val ALLOWSCENARIOS = FeatHeadr.ALLOWSCENARIOS
        /**
         * enhanced protection settings: Format cells
         */
        val ALLOWFORMATCELLS = FeatHeadr.ALLOWFORMATCELLS
        /**
         * enhanced protection settings: Format columns
         */
        val ALLOWFORMATCOLUMNS = FeatHeadr.ALLOWFORMATCOLUMNS
        /**
         * enhanced protection settings: Format rows
         */
        val ALLOWFORMATROWS = FeatHeadr.ALLOWFORMATROWS
        /**
         * enhanced protection settings: Insert columns
         */
        val ALLOWINSERTCOLUMNS = FeatHeadr.ALLOWINSERTCOLUMNS
        /**
         * enhanced protection settings: Insert rows
         */
        val ALLOWINSERTROWS = FeatHeadr.ALLOWINSERTROWS
        /**
         * enhanced protection settings: Insert hyperlinks
         */
        val ALLOWINSERTHYPERLINKS = FeatHeadr.ALLOWINSERTHYPERLINKS
        /**
         * enhanced protection settings: Delete columns
         */
        val ALLOWDELETECOLUMNS = FeatHeadr.ALLOWDELETECOLUMNS
        /**
         * enhanced protection settings: Delete rows
         */
        val ALLOWDELETEROWS = FeatHeadr.ALLOWDELETEROWS
        /**
         * enhanced protection settings: Select locked cells
         */
        val ALLOWSELLOCKEDCELLS = FeatHeadr.ALLOWSELLOCKEDCELLS
        /**
         * enhanced protection settings: Sort
         */
        val ALLOWSORT = FeatHeadr.ALLOWSORT
        /**
         * enhanced protection settings: Use Autofilter
         */
        val ALLOWAUTOFILTER = FeatHeadr.ALLOWAUTOFILTER
        /**
         * enhanced protection settings: Use PivotTable reports
         */
        val ALLOWPIVOTTABLES = FeatHeadr.ALLOWPIVOTTABLES
        /**
         * enhanced protection settings: Select unlocked cells
         */
        val ALLOWSELUNLOCKEDCELLS = FeatHeadr.ALLOWSELUNLOCKEDCELLS
    }

}
/**
 * Remove a Row and all associated Cells from this WorkSheet.
 *
 * @param int rownum - the number of the row to remove uses default row deletion
 * rules regarding updating references
 */
/**
 * Add a Cell with the specified value to a WorkSheet.
 *
 *
 * This method determines the Cell type based on type-compatibility of the
 * value.
 *
 *
 * In other words, if the Object cannot be converted safely to a Numeric Object
 * type, then it is treated as a String and a new String value is added to the
 * WorkSheet at the Cell address specified.
 *
 * @param obj the value of the new Cell
 * @param int row the row of the new Cell
 * @param int col the column of the new Cell
 */
/**
 * Add a Cell with the specified value to a WorkSheet.
 *
 *
 * This method determines the Cell type based on type-compatibility of the
 * value.
 *
 *
 * In other words, if the Object cannot be converted safely to a Numeric Object
 * type, then it is treated as a String and a new String value is added to the
 * WorkSheet at the Cell address specified.
 *
 *
 * If a validation record for the cell exists the validation is checked for a
 * correct value, if the value does not pass the validation a
 * ValidationException will be thrown
 *
 * @param obj the value of the new Cell
 * @param int row the row of the new Cell
 * @param int col the column of the new Cell
 */
/**
 * Returns a CellHandle for working with the value of a Cell on a WorkSheet.
 *
 *
 * returns a new CellHandle with each call
 *
 *
 * use caching method getCell(int row, int col, boolean cache) to control
 * caching of CellHandles.
 *
 * @param int Row the integer row of the Cell
 * @param int Col the integer col of the Cell
 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified
 * address.
 */