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
package io.starter.OpenXLS

import io.starter.formats.LEO.BlockByteReader
import io.starter.formats.LEO.InvalidFileException
import io.starter.formats.LEO.LEOFile
import io.starter.formats.XLS.*
import io.starter.formats.XLS.charts.Chart
import io.starter.formats.XLS.charts.OOXMLChart
import io.starter.toolkit.*

import java.io.*
import java.net.URL
import java.nio.ByteBuffer
import java.nio.ByteOrder
import java.util.*

/**
 * The WorkBookHandle provides a handle to the XLS file and includes convenience
 * methods for working with the WorkSheets and Cell values within the XLS file.
 * For example: <br></br>
 * `
 * WorkBookHandle  book  = new WorkBookHandle("testxls.xls");<br></br>
 * WorkSheetHandle sheet = book.getWorkSheet("Sheet1");<br></br>
 * CellHandle      cell  = sheet.getCell("B22");<br></br>
` *
 *
 *
 *
 * By default, OpenXLS will lock open WorkBook files. To close the file after
 * parsing and work with a temporary file instead, use the following setting:
 * <br></br>
 * `
 * System.getProperties().put(WorkBookHandle.USETEMPFILE, "true");
` * <br></br>
 * If you enable this mode you will need to periodically clean up the generated
 * temporary files in your working directory. All OpenXLS temporary file names
 * begin with "ExtenXLS_".
 */
class WorkBookHandle : DocumentHandle, WorkBook, Handle {

    /**
     * Returns a low-level WorkBook.
     *
     *
     * NOTE: The WorkBook class is NOT a part of the published API. Any of the
     * methods and/or variables on a WorkBook object are subject to change without
     * notice in new versions of OpenXLS.
     */
    override var workBook: io.starter.formats.XLS.WorkBook? = null
        protected set
    /**
     * Gets the internal LEOFile object.
     *
     *
     * **WARNING:** This method is not part of the public API. Its use
     * is not supported and behavior is subject to change.
     *
     */
    var leoFile: LEOFile? = null
        protected set
    /**
     * Gets the internal Factory object.
     *
     *
     * **WARNING:** This method is not part of the public API. Its use
     * is not supported and behavior is subject to change.
     *
     */
    var factory: WorkBookFactory? = null
        protected set

    protected var plist: ProgressListener? = null

    /**
     * Returns all strings that are in the SharedStringTable for this workbook. The
     * SST contains all standard string records in cells, but may not include such
     * things as strings that are contained within formulas. This is useful for such
     * things as full text indexing of workbooks
     *
     * @return Strings in the workbook.
     */
    val allStrings: Array<String>
        get() = this.workBook!!.allStrings

    /**
     * Get either the default color table, or the color table from the custom
     * palatte(if exists) from the WorkBook
     *
     * @return
     */
    override val colorTable: Array<java.awt.Color>
        get() = this.workBook!!.colorTable

    /**
     * Returns whether this book uses the 1904 date format.
     *
     */
    val is1904: Boolean
        @Deprecated("Use {@link #getDateFormat()} instead.")
        get() = workBook!!.is1904

    /**
     * Gets the date format used by this book.
     */
    override val dateFormat: DateConverter.DateFormat
        get() = workBook!!.dateFormat

    /**
     * Returns the lowest version of Excel compatible with the input file.
     *
     * @return an Excel version string
     */
    val xlsVersionString: String
        get() = workBook!!.xlsVersionString

    /**
     * Return useful statistics about this workbook.
     *
     * @return a string contatining various statistics.
     */
    val stats: String
        get() = workBook!!.stats

    /**
     * Returns an Array of the CellRanges existing in this WorkBook specifically the
     * Ranges referenced in Formulas, Charts, and Named Ranges.
     *
     *
     * This is necessary to allow for automatic updating of references when
     * adding/removing/moving Cells within these ranges, as well as shifting
     * references to Cells in Formulas when Formula records are moved.
     *
     * @return all existing Cell Range references used in Formulas, Charts, and
     * Names
     */
    val cellRanges: Array<CellRange>
        get() = this.workBook!!.refTracker!!.cellRanges

    /**
     * get an array of handles to all PivotTables in the WorkBook
     *
     * @return PivotTable[] all of the WorkBooks PivotTables
     */
    override val pivotTables: Array<PivotTableHandle>
        @Throws(PivotTableNotFoundException::class)
        get() {
            val sxv = workBook!!.allPivotTableViews
            if (sxv == null || sxv.size == 0)
                throw PivotTableNotFoundException(
                        "There are no PivotTables defined in: " + this.name)
            val pth = arrayOfNulls<PivotTableHandle>(sxv.size)
            for (t in pth.indices) {
                pth[t] = PivotTableHandle(sxv[t], this)
            }
            return pth
        }

    /**
     * Get the calculation mode for the workbook.
     *
     *
     * CALCULATE_ALWAYS is the default for new workbooks. Calling Cell.getVal() will
     * calculate formulas if they exist within the cell.
     *
     *
     * CALCULATE_EXPLICIT will return present value of the cell. Formula calculation
     * will only occur when explicitly called through the Formula Handle
     *
     *
     * WorkBookHandle.CALCULATE_ALWAYS -- recalc every time the cell value is
     * requested (no cacheing) WorkBookHandle.CALCULATE_EXPLICIT -- recalc only when
     * FormulaHandle.calculate() called WorkBookHandle.CALCULATE_AUTO -- only recac
     * when changes
     *
     * @param CalcMode Calculation mode to use in workbook.
     */
    /**
     * Set the calculation mode for the workbook.
     *
     *
     * CALCULATE_AUTO is the default for new workbooks. Calling Cell.getVal() will
     * calculate formulas if they exist within the cell.
     *
     *
     * CALCULATE_EXPLICIT will return present, cached value of the cell. Formula
     * calculation will ONLY occur when explicitly called through the Formula
     * Handle.calculate() method.
     *
     *
     * CALCULATE_ALWAYS will ignore the cache and force a recalc every time a cell
     * value is requested.
     *
     *
     *
     *
     * WorkBookHandle.CALCULATE_AUTO WorkBookHandle.CALCULATE_ALWAYS
     * WorkBookHandle.CALCULATE_EXPLICIT
     *
     * @param CalcMode Calculation mode to use in workbook.
     */
    override var formulaCalculationMode: Int
        get() = workBook!!.calcMode
        set(CalcMode) {
            workBook!!.calcMode = CalcMode
        }

    /**
     * Returns all ImageHandles in the workbook
     *
     * @return
     */
    val images: Array<ImageHandle>
        get() {
            val ret = Vector<ImageHandle>()
            for (t in 0 until this.numWorkSheets) {
                try {
                    val r = this.getWorkSheet(t).images
                    for (x in r!!.indices)
                        ret.add(r!![x])
                } catch (ex: Exception) {
                }

            }
            val retx = arrayOfNulls<ImageHandle>(ret.size)
            ret.toTypedArray<ImageHandle>()
            return retx
        }

    /**
     * Returns all Chart Handles contained in the WorkBook
     *
     * @return ChartHandle[] an array of all Charts in the WorkBook
     */
    override val charts: Array<ChartHandle>
        get() {
            val cv = workBook!!.chartVect
            val cht = arrayOfNulls<ChartHandle>(cv.size)
            for (x in cv.indices) {
                cht[x] = ChartHandle(cv[x] as Chart, this)
            }
            return cht
        }

    /**
     * Returns all Named Range Handles
     *
     * @return NameHandle[] all of the Named ranges in the WorkBook
     */
    override val namedRanges: Array<NameHandle>
        get() {
            val nand = workBook!!.names
            val nands = arrayOfNulls<NameHandle>(nand.size)
            for (x in nand.indices) {
                nands[x] = NameHandle(nand[x], this)
            }
            return nands
        }

    /**
     * Returns all Named Range Handles scoped to WorkBook.
     *
     *
     * Note this will not include worksheet scoped named ranges
     *
     * @return NameHandle[] all of the Named ranges that are scoped to WorkBook
     */
    val namedRangesInScope: Array<NameHandle>
        get() {
            val nand = workBook!!.workbookScopedNames
            val nands = arrayOfNulls<NameHandle>(nand.size)
            for (x in nand.indices) {
                nands[x] = NameHandle(nand[x], this)
            }
            return nands
        }

    /**
     * Returns the name of this WorkBook
     *
     * @return String name of WorkBook
     */
    override var name: String
        get() = if (name != null) name else "New Spreadsheet"
        set

    /**
     * Returns an array containing all cells in the WorkBook
     *
     * @return CellHandle array of all book cells
     */
    override// handle Mulblanks: ref a range of cells; to get correct
    // cell address,
    // traverse thru range and set cellhandle ref to correct
    // column
    // for Mulblank use only -sets correct
    // column reference for multiple blank
    // cells
    // ...
    val cells: Array<CellHandle>
        get() {
            val allcz = this.workBook!!.cells
            val ret = arrayOfNulls<CellHandle>(allcz.size)
            var aMul: Mulblank? = null
            var c: Short = -1
            for (t in ret.indices) {
                ret[t] = CellHandle(allcz[t], this)
                if (allcz[t].opcode == XLSConstants.MULBLANK) {
                    if (allcz[t] === aMul) {
                        c++
                    } else {
                        aMul = allcz[t] as Mulblank
                        c = aMul.colFirst.toShort()
                    }
                    ret[t].setBlankRef(c.toInt())
                }
            }
            return ret
        }

    /**
     * Returns the number of Cells in this WorkBook
     *
     * @return int number of Cells
     */
    override val numCells: Int
        get() = workBook!!.numCells

    /**
     * Gets the spreadsheet as a byte array in BIFF8 (Excel '97-2003) format.
     *
     */
    override val bytes: ByteArray?
        @Deprecated("Writing the spreadsheet to a byte array uses a great deal of\n" +
                "      memory and generally provides no benefit over streaming output.\n" +
                "      Use the {@link #write} family of methods instead. If you need a\n" +
                "      byte array use {@link ByteArrayOutputStream}.")
        get() {
            try {
                val bout = ByteArrayOutputStream()
                writeBytes(bout)

                return bout.toByteArray()
            } catch (e1: Exception) {
                Logger.logErr("Getting Spreadsheet bytes failed.", e1)
                return null
            }

        }

    /**
     * Gets the constant representing this document's native format.
     */
    override val format: Int
        get() {
            val name = this.fileName.toLowerCase()

            return if (this.isExcel2007) {
                if (OOXMLAdapter.hasMacros(this))
                    if (name.endsWith(".xltm")) FORMAT_XLTM else FORMAT_XLSM
                else
                    if (name.endsWith(".xltx")) FORMAT_XLTX else FORMAT_XLSX
            } else
                FORMAT_XLS
        }

    override val fileExtension: String
        get() {
            when (this.format) {
                FORMAT_XLSX -> return ".xlsx"
                FORMAT_XLSM -> return ".xlsm"
                FORMAT_XLTX -> return ".xltx"
                FORMAT_XLTM -> return ".xltm"
                FORMAT_XLS -> return ".xls"
                else -> return ""
            }
        }

    /**
     * Returns whether the underlying spreadsheet is in Excel 2007 format by
     * default.
     *
     *
     * Even if this method returns true, it is still possible to write out the file
     * as a BIFF8 (Excel 97-2003) file, but unsupported features will be dropped,
     * and some files could experience corruption.
     *
     * @return whether the underlying spreadsheet is Excel 2007 format
     */
    /**
     * Sets whether this Workbook is in Excel 2007 format. Excel 2007 format
     * contains larger maximum column and row contraints, for example. <br></br>
     * Even if the workbook is set to Excel 2007 format, it is still possible to
     * write out the file as a BIFF8 (Excel 97-2003) file, but unsupported features
     * will be dropped, and some files could experience corruption.
     *
     * @param isExcel2007
     */
    var isExcel2007: Boolean
        get() = this.workBook!!.isExcel2007
        set(isExcel2007) {
            this.workBook!!.isExcel2007 = isExcel2007
        }

    /**
     * Returns an array of handles to all of the WorkSheets in the Workbook.
     *
     * @return WorkSheetHandle[] Array of all WorkSheets in WorkBook
     */
    override val workSheets: Array<WorkSheetHandle>?
        get() {
            try {
                if (factory != null) {
                    val numsheets = workBook!!.numWorkSheets
                    if (numsheets == 0)
                        throw WorkSheetNotFoundException(
                                "WorkBook has No Sheets.")
                    val sheets = arrayOfNulls<WorkSheetHandle>(numsheets)
                    for (i in 0 until numsheets) {
                        val bs = workBook!!.getWorkSheetByNumber(i)
                        bs.workBook = this.workBook
                        sheets[i] = WorkSheetHandle(bs, this)
                    }
                    return sheets
                }
                return null
            } catch (a: WorkSheetNotFoundException) {
                Logger.logWarn("getWorkSheets() failed: $a")
                return null
            }

        }

    internal var sheethandles: Hashtable<String, WorkSheetHandle>? = Hashtable()

    /**
     * returns the active or selected worksheet tab
     *
     * @return WorkSheetHandle
     * @throws WorkSheetNotFoundException
     */
    val activeSheet: WorkSheetHandle
        @Throws(WorkSheetNotFoundException::class)
        get() = this.getWorkSheet(this.workBook!!.selectedSheetNum)

    /**
     * Returns a WorkBookHandle containing an empty version of this WorkBook.
     *
     *
     * Use in conjunction with addSheetFromWorkBook() to create new output WorkBooks
     * containing various sheets from a master template.
     *
     *
     * ie: WorkBookHandle emptytemplate = this.getNoSheetWorkBook();
     * emptytemplate.addSheetFromWorkBook(this, "Sheet1", "TargetSheet");
     *
     * @return WorkBookHandle - the empty WorkBookHandle duplicate
     * @see addSheetFromWorkBook
     */
    override// to avoid ByteStreamer.stream records expansion
    val noSheetWorkBook: WorkBookHandle
        get() {
            val recs = this.workBook!!.streamer.biffRecords
            val gb = this.bytes
            val ret = WorkBookHandle(gb)
            this.workBook!!.streamer.setBiffRecords(Arrays.asList(*recs))
            ret.removeAllWorkSheets()
            return ret
        }

    /**
     * Returns the number of Sheets in this WorkBook
     *
     * @return int number of Sheets
     */
    val numWorkSheets: Int
        get() = workBook!!.numWorkSheets

    /**
     * Returns an array of all FormatHandles in the workbook
     */
    override// passing (this) with the format handle breaks the
    // relationship to the font.
    // if you need to pass it in we will have to handle it
    // differently
    val formats: Array<FormatHandle>
        get() {
            val l = this.workBook!!.xfrecs
            val formats = arrayOfNulls<FormatHandle>(l.size)
            val its = l.iterator()
            var i = 0
            while (its.hasNext()) {
                val x = its.next() as Xf
                try {
                    formats[i] = FormatHandle()
                    formats[i].workBook = this.workBook
                    formats[i].xf = x
                } catch (ex: Exception) {
                }

                i++
            }
            return formats
        }

    /**
     * Returns an array of all Conditional Formats in the workbook
     *
     *
     * these are formats referenced and used by the conditionally formatted ranges
     * in the workbook.
     *
     * @return
     */
    // the idea is to create a fake IXFE for use by
    // sheetster to find formats
    // int cfxe = this.getWorkBook().getNumFormats() + 50000; //
    // there would have to
    // be 50k styles on the sheet to conflict here....
    // int cfxe = this.getWorkBook().getNumXfs() + 50000; //
    // there would have to be
    // 50k styles on the sheet to conflict here....
    // cfm.initCells(this); // added!
    // cfm.setCfxe(cfxe);
    // cxfe++;
    val conditionalFormats: Array<FormatHandle>
        get() {

            val retl = Vector<FormatHandle>()
            val v = this.workBook!!.sheetVect

            val its = v.iterator()

            while (its.hasNext()) {
                val shtx = its.next() as Boundsheet
                val fmtlist = shtx.conditionalFormats
                val ixa = fmtlist!!.iterator()
                while (ixa.hasNext()) {
                    val cfm = ixa.next() as Condfmt
                    val cfxe = cfm.cfxe
                    val fz = FormatHandle(cfm, this, cfxe, null)

                    fz.formatId = cfxe
                    retl.add(fz)
                }
            }
            val formats = arrayOfNulls<FormatHandle>(retl.size)
            for (t in formats.indices)
                formats[t] = retl[t]
            return formats
        }

    // TODO Auto-generated method stub
    val workingDirectory: String?
        get() = null

    /**
     * Searches all Cells in the workbook for the string occurrence and replaces
     * with the replacement text.
     *
     * @return the number of replacements that were made
     */
    fun searchAndReplace(searchfor: String, replacewith: String): Int {
        val cx = cells
        var foundcount = 0
        for (t in cx.indices) {
            if (cx[t].cell !is Formula) {
                // find the string
                if (!cx[t].isNumber) {
                    val v = cx[t].stringVal
                    if (v!!.indexOf(searchfor) > -1) {
                        cx[t].`val` = StringTool
                                .replaceText(v, searchfor, replacewith)
                        foundcount++
                    }
                }
            }
        }
        return foundcount
    }

    /**
     * Return useful statistics about this workbook.
     *
     * @param use html line breaks
     * @return a string contatining various statistics.
     */
    fun getStats(usehtml: Boolean): String {
        return workBook!!.getStats(usehtml)
    }

    /**
     * Returns the Cell at the specified Location
     *
     * @param address
     * @return
     */
    @Throws(CellNotFoundException::class, WorkSheetNotFoundException::class)
    override fun getCell(address: String): CellHandle {
        val shtpos = address.indexOf("!")
        if (shtpos < 0)
            throw CellNotFoundException("$address not found.  You need to specify a location in the format: Sheet1!A1")
        val sheetstr = address.substring(0, shtpos)
        val sht = this.getWorkSheet(sheetstr)
        val celstr = address.substring(shtpos + 1)
        return sht!!.getCell(celstr)
    }

    /**
     * get a handle to a PivotTable in the WorkBook
     *
     * @param String name of the PivotTable
     * @return PivotTable the PivotTable
     */
    @Throws(PivotTableNotFoundException::class)
    override fun getPivotTable(ptname: String): PivotTableHandle {
        val st = workBook!!.getPivotTableView(ptname) ?: throw PivotTableNotFoundException(ptname)
        return PivotTableHandle(st, this)
    }

    /**
     * set the workbook to protected mode
     *
     *
     * Note: the password cannot be decrypted or changed in Excel -- protection can
     * only be set/removed using OpenXLS
     *
     * @param boolean whether to protect the book
     */
    override fun setProtected(protect: Boolean) {
        // TODO: Check that this behavior is correct
        // This is what the old implementation did

        val protector = workBook!!.protectionManager

        // Excel default... no kidding!
        if (protect)
            protector.password = "VelvetSweatshop"
        else
            protector.password = null

        protector.protected = protect
    }

    /**
     * set Default row height in twips (=1/20 of a point)
     *
     *
     * Note: only affects undefined Rows containing Cells
     *
     * @param int Default Row Height
     */
    // should be a double as Excel units are 1/20 of what is
    // stored in
    // defaultrowheight
    // e.g. 12.75 is Excel Units, twips = 12.75*20 = 256
    // (approx)
    // should expect users to use Excel units and target method
    // do the 20*
    // conversion
    override fun setDefaultRowHeight(t: Int) {
        workBook!!.setDefaultRowHeight(t)
    }

    /**
     * Set the default column width across all worksheets <br></br>
     * This setting is a worksheet level setting, so will be applied to all existing
     * worksheets. individual worksheets can also be set using
     * WorkSheetHandle.setDefaultColWidth
     *
     *
     * This setting is roughly the width of the character '0' The default width of a
     * column is 8.
     */
    override fun setDefaultColWidth(t: Int) {
        workBook!!.setDefaultColWidth(t)
    }

    /**
     * Returns a Formula Handle
     *
     * @return FormulaHandle a formula handle in the WorkBook
     */
    @Throws(FormulaNotFoundException::class)
    fun getFormulaHandle(celladdress: String): FormulaHandle {
        val formula = workBook!!.getFormula(celladdress)
        return FormulaHandle(formula, this)
    }

    /**
     * Returns an ImageHandle for manipulating images in the WorkBook
     *
     * @param imagename
     * @return
     */
    @Throws(ImageNotFoundException::class)
    fun getImage(imagename: String): ImageHandle {
        for (t in 0 until this.numWorkSheets) {
            try {
                val r = this.getWorkSheet(t).images
                for (x in r!!.indices)
                    if (r!![x].name == imagename)
                        return r!![x]
            } catch (ex: Exception) {
            }

        }
        throw ImageNotFoundException(
                "Image not found: $imagename in $this")
    }

    /**
     * Returns a Named Range Handle
     *
     * @return NameHandle a Named range in the WorkBook
     */
    @Throws(CellNotFoundException::class)
    override fun getNamedRange(rangename: String): NameHandle {
        val nand = workBook!!.getName(rangename.toUpperCase())
                ?: throw CellNotFoundException(rangename) // case-insensitive
        return NameHandle(nand, this)
    }

    /**
     * Returns a Named Range Handle if it exists in the specified scope.
     *
     *
     * This can be used to distinguish between multiple named ranges with the same
     * name but differing scopes
     *
     * @return NameHandle a Named range in the WorkBook that exists in the scope
     */
    @Throws(CellNotFoundException::class)
    fun getNamedRangeInScope(rangename: String): NameHandle {
        val nand = workBook!!.getScopedName(rangename) ?: throw CellNotFoundException(rangename)
        return NameHandle(nand, this)
    }

    /**
     * Create a named range in the workbook
     *
     *
     * Note that the named range designation can conform to excel specs, that is,
     * boolean values, references, or string variables can be set. Remember to
     * utilize the sheet name when setting referential names.
     *
     *
     *
     *
     * NameHandle nh = createNamedRange("cellRange", "Sheet1!A1:B3"); NameHandle nh
     * = createNamedRange("trueRange", "=true");
     *
     * @param name     The name that should be used to reference this named range
     * @param rangeDef Range of the cells for this named range, in excel syntax including
     * sheet name, ie "Sheet1!A1:D1"
     * @return NameHandle for modifying the named range
     */
    fun createNamedRange(name: String, rangeDef: String): NameHandle {
        return NameHandle(name, rangeDef, this)
    }

    /**
     * Returns a Chart Handle
     *
     * @return ChartHandle a Chart in the WorkBook
     */
    // KSC: NOTE: this methodology needs work as a book may
    // contain charts in
    // different sheets containing the same name
    // TODO: rethink
    @Throws(ChartNotFoundException::class)
    override fun getChart(chartname: String): ChartHandle {
        return ChartHandle(workBook!!.getChart(chartname), this)
    }

    /**
     * retrieve a ChartHandle via id
     *
     * @param id
     * @return
     * @throws ChartNotFoundException
     */
    @Throws(ChartNotFoundException::class)
    fun getChartById(id: Int): ChartHandle {
        val cv = workBook!!.chartVect
        var cht: Chart? = null
        for (x in cv.indices) {
            cht = cv[x] as Chart
            if (cht.id == id)
                return ChartHandle(cht, this)
        }
        throw ChartNotFoundException("Id $id")
    }

    /**
     * Returns whether the sheet selection tabs should be shown.
     */
    fun showSheetTabs(): Boolean {
        return workBook!!.showSheetTabs()
    }

    /**
     * Sets whether the sheet selection tabs should be shown.
     */
    fun setShowSheetTabs(show: Boolean) {
        workBook!!.setShowSheetTabs(show)
    }

    /**
     * Writes the document to the given path. If the filename ends with ".xlsx" or
     * ".xlsm", the workbook will be written as OOXML (XLSX). Otherwise it will be
     * written as BIFF8 (XLS). For OOXML, if the file has a VBA project the file
     * extension must be ".xlsm". It will be changed if necessary.
     *
     * @param path the path to which the document should be written
     */
    @Deprecated("The filename-based format choosing is counter-intuitive and\n" +
            "      failure-prone. Use {@link #write(OutputStream, int)} instead.")
    fun write(path: String) {
        val ext = path.toLowerCase()
        write(path, ext.endsWith(".xlsx") || ext.endsWith(".xlsm"))
    }

    /**
     * Writes the document to the given file in either XLS or XLSX. For OOXML, if
     * the file has a VBA project the file extension must be ".xlsm". It will be
     * changed if necessary.
     *
     * @param path  the path to which the document should be written
     * @param ooxml If `true`, write as OOXML (XLSX). Otherwise, write as
     * BIFF8 (XLS).
     */
    @Deprecated("The boolean format parameter is not flexible enough to represent\n" +
            "      all supported formats. Use {@link #write(File, int)} instead.")
    fun write(path: String, ooxml: Boolean) {
        var path = path
        val format: Int
        if (ooxml) {
            if (this.isExcel2007)
                format = this.getFormat(path)
            else
                format = FORMAT_XLSX

            if (!OOXMLAdapter.hasMacros(this))
                path = StringTool.replaceExtension(path, ".xlsx")
            else
            // it's a macro-enabled workbook
                path = StringTool.replaceExtension(path, ".xlsm")
        } else
            format = FORMAT_XLS

        try {
            this.write(File(path), format)
        } catch (e: Exception) {
            throw WorkBookException("error writing workbook",
                    WorkBookException.WRITING_ERROR, e)
        }

    }

    /**
     * Writes the document to the given stream in either XLS or XLSX format.
     *
     * @param dest  the stream to which the document should be written
     * @param ooxml If `true`, write as OOXML (XLSX). Otherwise, write as
     * BIFF8 (XLS).
     */
    @Deprecated("The boolean format parameter is not flexible enough to represent\n" +
            "      all supported formats. Use {@link #write(OutputStream, int)}\n" +
            "      instead.")
    fun write(dest: OutputStream, ooxml: Boolean) {
        val format: Int
        if (ooxml) {
            if (this.isExcel2007)
                format = this.format
            else
                format = FORMAT_XLSX
        } else
            format = FORMAT_XLS

        try {
            if (format > WorkBookHandle.FORMAT_XLS && this.file != null) {
                OOXMLAdapter.refreshPassThroughFiles(this)
            }
            this.write(dest, format)
        } catch (e: Exception) {
            throw WorkBookException("error writing workbook",
                    WorkBookException.WRITING_ERROR, e)
        }

    }

    /**
     * Gets the constant representing this document's desired format
     */
    fun getFormat(path: String?): Int {
        if (path == null)
            return format
        return if (this.isExcel2007) {
            if (OOXMLAdapter.hasMacros(this))
                if (path.endsWith(".xltm")) FORMAT_XLTM else FORMAT_XLSM
            else
                if (path.endsWith(".xltx")) FORMAT_XLTX else FORMAT_XLSX
        } else
            FORMAT_XLS
    }

    /**
     * Writes the document to the given stream in the requested format.
     *
     *
     * format choices:
     *
     *
     * WorkBookHandle.FORMAT_XLS for 2003 and previous versions <br></br>
     * WorkBookHandle.FORMAT_XLSX for non-macro-enabled 2007 version <br></br>
     * WorkBookHandle.FORMAT_XLSM for macro-enabled 2007 version <br></br>
     * WorkBookHandle.FORMAT_XLTM for macro-enabled 2007 templates. <br></br>
     * WorkBookHandle.FORMAT_XLTX for 2007 templates,
     *
     *
     * **IMPORTANT NOTE:** if the resulting filename contains the .XLSM extension
     * <br></br>
     * the WorkBook **MUST** be written in FORMAT_XLSM; otherwise open errors
     * will occur
     *
     *
     * **NOTE:** If the format is FORMAT_XLSX and the filename contains macros
     * <br></br>
     * the file will be written as Macro-Enabled i.e. in FORMAT_XLSM. In these
     * cases, <br></br>
     * the filename must contain the .XLSM extension
     *
     * @param dest   the stream to which the document should be written
     * @param format the constant representing the desired output format
     * @throws IllegalArgumentException if the given type code is invalid
     * @throws IOException              if an error occurs while writing to the stream
     */
    @Throws(IOException::class)
    override fun write(dest: OutputStream, format: Int) {
        var format = format
        if (format == DocumentHandle.FORMAT_NATIVE)
            format = this.format

        when (format) {
            FORMAT_XLSX, FORMAT_XLSM, FORMAT_XLTX, FORMAT_XLTM -> try {
                if (this.file != null)
                    OOXMLAdapter.refreshPassThroughFiles(this)

                val adapter = OOXMLWriter()
                adapter.format = format
                adapter.getOOXML(this, dest)
            } catch (e: IOException) {
                throw e
            } catch (e: Exception) {
                // TODO: OOXMLAdapter only throws IOException, change its
                // throws
                throw WorkBookException("error writing workbook",
                        WorkBookException.WRITING_ERROR, e)
            }

            FORMAT_XLS -> try {
                this.workBook!!.streamer.writeOut(dest)
            } catch (e: io.starter.formats.XLS.WorkBookException) {
                val cause = e.cause
                if (cause is IOException)
                    throw cause
                throw e
            }

            else -> throw IllegalArgumentException("unknown output format")
        }
    }

    /**
     * Writes the document to the given stream in the requested OOXML format.
     *
     * @param dest   the stream to which the document should be written
     * @param format the constant representing the desired output format
     * @throws IllegalArgumentException if the given type code is invalid
     * @throws IOException              if an error occurs while writing to the stream
     */
    @Deprecated("This method is like {@link #write(OutputStream, int)} except it\n" +
            "      only supports OOXML formats. Use that instead.")
    @Throws(Exception::class)
    fun writeXLSXBytes(dest: OutputStream, format: Int) {
        this.write(dest, format)
    }

    /**
     * Writes the document to the given stream in the default OOXML format.
     *
     * @param dest the stream to which the document should be written
     * @throws IOException if an error occurs while writing to the stream
     */
    @Deprecated("Use {@link #write(OutputStream, int}) instead.")
    @Throws(Exception::class)
    fun writeXLSXBytes(dest: OutputStream) {
        this.write(dest, true)
    }

    /**
     * Writes the document to the given stream in BIFF8 (XLS) format.
     *
     *
     *
     * To output a debugging `StringBuffer`, you must first set the
     * autolockdown setting:<br></br>
     * `
     * props.put("io.starter.OpenXLS.autocreatelockdown","true");
    ` *
     *
     *
     * @param dest the stream to which the document should be written
     * @return for debugging: a StringBuffer containing an output of the record
     * bytes streamed
     */
    @Deprecated("Use {@link #write(OutputStream, int)} instead.")
    override fun writeBytes(dest: OutputStream): StringBuffer? {
        return workBook!!.streamer.writeOut(dest)
    }

    /**
     * Default constructor creates a new, empty Spreadsheet with
     *
     *
     * 3 WorkSheets: "Sheet1","Sheet2",and "Sheet3".
     */
    constructor() {
        // Xf.DEFAULTIXFE= 15; // reset to default in cases of
        // having previously read
        // Excel2007 template which may have set defaultXF
        // differently
        this.initDefault()
    }

    /**
     * Constructor creates a new, empty Spreadsheet with 3 worksheets: "Sheet1",
     * "Sheet2" and "Sheet3" <br></br>
     * This version allows flagging the workbook as Excel 2007 format. <br></br>
     * Excel 2007 format contains larger maximum column and row contraints, for
     * example. <br></br>
     * Even if the workbook is set to Excel 2007 format, it is still possible to
     * write out the file as a BIFF8 (Excel 97-2003) file, but unsupported features
     * will be dropped, and some files could experience corruption.
     *
     * @param boolean Excel2007 - true if set to Excel 2007 version
     */
    constructor(Excel2007: Boolean) {
        this.initDefault()
        this.isExcel2007 = Excel2007
    }

    /**
     * another handle to the useful ability to load a book from the prorotype bytes
     */
    protected fun initDefault() {
        try {
            val b = prototypeBook ?: throw io.starter.formats.XLS.WorkBookException(
                    "Unable to load prototype workbook.",
                    WorkBookException.LICENSING_FAILED)
            val bbf = ByteBuffer.wrap(b)
            bbf.order(ByteOrder.LITTLE_ENDIAN)
            leoFile = LEOFile(bbf)
        } catch (e: Exception) {
            throw InvalidFileException(
                    "WorkBook could not be instantiated: $e")
        }

        this.initFromLeoFile(leoFile)
    }

    /**
     * constructor which takes an InputStream containing the bytes of a valid XLS
     * file.
     *
     * @param InputStream contains the valid BIFF8 bytes for reading
     */
    constructor(inx: InputStream) {
        this.initFromStream(inx)
    }

    /**
     * Initialization of this workbook handle from a leoFile;
     */
    private fun initFromLeoFile(leo: LEOFile) {
        this.leoFile = leo
        try {
            val bar = leoFile!!.xlsBlockBytes
            this.initBytes(bar)
            this.isExcel2007 = false
            leoFile!!.clearAfterInit()
        } catch (e: Exception) {
            if (e is io.starter.formats.XLS.WorkBookException)
                throw e
            throw io.starter.formats.XLS.WorkBookException(
                    "ERROR: instantiating WorkBookHandle failed: $e",
                    WorkBookException.UNSPECIFIED_INIT_ERROR, e)
        }

    }

    /**
     * Initialize this workbook from a stream, unfortunately our byte backer
     * requires a file, so create a tempfile and init from that
     */
    protected fun initFromStream(input: InputStream) {
        try {
            val target = TempFileManager.createTempFile("WBP", ".tmp")

            JFileWriter.writeToFile(input, target)
            this.initFromFile(target.absoluteFile)
            if (this.leoFile != null)
            // it would be if XLSX or XLSM ...
            // 20090323 KSC
                this.leoFile!!.closefb()
            // this.myLEOFile.close(); // close now flushes buffers +
            // storages ...
            input.close()

            val fdel = File(target.toString())
            if (!fdel.delete()) {
                if (this.DEBUGLEVEL > Document.DEBUG_LOW)
                    Logger.logWarn("Could not delete tempfile: $target")
            }
        } catch (ex: IOException) {
            Logger.logErr("Initializing WorkBookHandle failed.", ex)
        }

    }

    /**
     * Create a new WorkBookHandle from the byte array passed in. Byte array passed
     * in must contain a valid xls or xlsx workbook file
     *
     * @param byte[] byte array containing the valid XLS or XLSX file for reading
     */
    constructor(barray: ByteArray) {
        initializeFromByteArray(barray)
    }

    /**
     * Protected method that handles WorkBookHandle(byte[]) constructor
     *
     * @param barray
     */
    protected fun initializeFromByteArray(barray: ByteArray) {
        // check first bytes to see if this is a zipfile (OOXML)
        if (barray[0].toChar() == 'P' && barray[1].toChar() == 'K') {
            try {
                // added "." fixes Baxter Open Bug [BugTracker 2909]
                val ftmp = TempFileManager.createTempFile("WBP", ".tmp")
                val fous = FileOutputStream(ftmp)
                fous.write(barray)
                fous.flush()
                fous.close()
                this.initFromFile(ftmp)
                return
            } catch (e: Exception) {
                Logger.logErr("Could not parse XLSX from bytes.$e")
                return
            }

        }

        val bbf = ByteBuffer.wrap(barray)
        bbf.order(ByteOrder.LITTLE_ENDIAN)
        leoFile = LEOFile(bbf)
        if (leoFile!!.hasWorkBook()) {
            try {
                val bar = leoFile!!.xlsBlockBytes
                this.initBytes(bar)
            } catch (e: Throwable) {
                if (e is OutOfMemoryError)
                    throw e
                if (e is WorkBookException)
                    throw e
                val errstr = "Instantiating WorkBookHandle failed: $e"
                throw io.starter.formats.XLS.WorkBookException(errstr,
                        WorkBookException.UNSPECIFIED_INIT_ERROR)
            }

        } else {
            Logger.logWarn("Initializing WorkBookHandle failed: byte array does not contain a supported Excel WorkBook.")
            throw InvalidFileException(
                    "byte array does not contian a supported Excel WorkBook.")
        }
    }

    /**
     * Fetches a workbook from a URL
     *
     *
     * If you need to authenticate your URL connection first then use the
     * InputStream constructor
     *
     * @param urlx
     * @return
     * @throws Exception
     */
    constructor(url: URL) : this(DocumentHandle.getFileFromURL(url)) {}/*
         * OK, both this method and the (inputstream) constructor
         * set a temp file, is
         * this not possible to do without hitting the disk? TODO:
         * look into fix
         */

    /**
     * constructor which takes the XLS file name and has an optional debug setting
     * to assist with output. Setting this value will cause verbose logging and is
     * discouraged unless required for support.
     *
     * @param String filePath the name of the XLS file to read
     * @param Debug  level
     */
    @JvmOverloads
    constructor(filePath: String, debug: Int = 0) {
        this.debugLevel = debug
        val f = File(filePath)
        this.initFromFile(f)
        this.file = f // XXX KSC: Save for potential re-input of pass-through
        // ooxml files

    }

    /**
     * constructor which takes the XLS file
     *
     * @param File  the XLS file to read
     * @param Debug level
     */
    constructor(fx: File?) {
        this.initFromFile(fx!!)
    }

    protected fun initWorkBookFactory() {
        factory = WorkBookFactory()
    }

    /**
     * initialize from an XLSX/OOXML workbook.
     */
    private fun initXLSX(fname: String): Boolean {
        // do before parseNBind so can set myfactory & fname
        // set state vars for this workbookhandle
        this.initWorkBookFactory()

        factory!!.debugLevel = this.DEBUGLEVEL
        factory!!.fileName = this.name

        if (plist != null)
            factory!!.register(plist!!) // register progress notifier
        try {
            // iterate sheets,inputting cell values, named ranges and
            // formula strings
            val oe = OOXMLReader()
            val bk = WorkBookHandle()
            bk.removeAllWorkSheets()
            factory!!.debugLevel = this.DEBUGLEVEL
            bk.DEBUGLEVEL = this.DEBUGLEVEL
            oe.parseNBind(bk, fname)
            this.sheethandles = bk.sheethandles
            this.workBook = bk.workBook
        } catch (e: Exception) {
            throw WorkBookException(
                    "WorkBookHandle OOXML Read failed: $e",
                    WorkBookException.UNSPECIFIED_INIT_ERROR, e)
        }

        workBook!!.isExcel2007 = true
        return true
    }

    /**
     * do all initialization with a filename
     *
     * @param fname
     */
    protected fun initFromFile(fx: File) {
        val fname = fx.path
        var finch = ""

        // handle csv import
        val fincheck: FileReader
        try {
            fincheck = FileReader(fx)
            if (fx.length() > 100) {
                val cbuf = CharArray(100)
                fincheck.read(cbuf)
                finch = String(cbuf)
            }
            fincheck.close()

        } catch (e: FileNotFoundException) {
            Logger.logErr("WorkBookHandle: Cannot open file " + fname + ": "
                    + e)
        } catch (e1: Exception) {
            Logger.logErr("Invalid XLSX/OOXML File.")
        }

        this.name = fname // 20081231 KSC: set here
        if (finch.toUpperCase().startsWith("PK")) { // it's a zip file... give
            // XLSX parsing a shot
            if (this.file != null)
                OOXMLAdapter.refreshPassThroughFiles(this)
            if (initXLSX(fname))
                return
        }
        try {
            leoFile = LEOFile(fx, this.DEBUGLEVEL)
        } catch (ifx: InvalidFileException) {
            if (finch.indexOf(",") > -1 && finch.indexOf(",") > -1) {
                // init a blank workbook
                this.initDefault()

                // map CSV into workbook
                try {
                    val sheet = getWorkSheet(0)
                    sheet.readCSV(BufferedReader(FileReader(fx)))
                    return
                } catch (e: Exception) {
                    throw WorkBookException(
                            "Error encountered importing CSV: $e",
                            WorkBookException.ILLEGAL_INIT_ERROR)
                }

            } else {
                throw ifx
            }

        }

        if (leoFile!!.hasWorkBook()) {
            this.initFromLeoFile(leoFile)
        } else {
            // total failure to load
            Logger.logErr("Initializing WorkBookHandle failed: " + fname
                    + " does not contain a supported Excel WorkBook.")
            throw InvalidFileException(
                    "$fname does not contian a supported Excel WorkBook.")
        }
    }

    /**
     * Constructor which takes a ProgressListener which monitors the progress of
     * creating a new Excel file.
     *
     * @param ProgressListener object which is monitoring progress of WorkBook read
     */
    constructor(pn: ProgressListener) {
        this.plist = pn
        try {
            val b = prototypeBook
            val bbf = ByteBuffer.wrap(b)
            bbf.order(ByteOrder.LITTLE_ENDIAN)
            leoFile = LEOFile(bbf)
        } catch (e: Exception) {
            throw InvalidFileException(
                    "WorkBook could not be instantiated: $e")
        }

        this.initFromLeoFile(leoFile)
    }

    /**
     * Constructor which takes the XLS file name
     *
     *
     * and a ProgressListener which monitors the progress of reading the Excel file.
     *
     * @param String           fname the name of the XLS file to read
     * @param ProgressListener object which is monitoring progress of WorkBook read
     */
    constructor(fname: String, pn: ProgressListener) {
        this.plist = pn
        this.initFromFile(File(fname))
    }

    /**
     * For internal creation of a workbook handle from
     *
     * @param leo
     */
    constructor(leo: LEOFile) {
        this.initFromLeoFile(leo)
    }

    /**
     * init the new WorkBookHandle
     */
    @Synchronized
    protected fun initBytes(blockByteReader: BlockByteReader) {
        this.initWorkBookFactory()

        if (plist != null)
            factory!!.register(plist!!) // register progress notifier
        factory!!.debugLevel = this.DEBUGLEVEL

        workBook = factory!!
                .getWorkBook(blockByteReader, leoFile) as io.starter.formats.XLS.WorkBook

        if (dump_input != null) {
            try {
                dump_input!!.flush()
                dump_input!!.close()
                dump_input = null
            } catch (e: Exception) {
            }

        }
        this.postLoad()
    }

    /**
     * Handles tasks that need to occur after workbook has been loaded
     */
    internal fun postLoad() {
        initHlinks()
        initMerges()
        workBook!!.initializeNames() // must initialize name expressions AFTER
        // loading sheet records
        workBook!!.mergeMSODrawingRecords()
        workBook!!.initializeIndirectFormulas()
        initPivotCache() // if any
    }

    internal fun initMerges() {
        val mergelookup = workBook!!.mergecelllookup
        for (t in mergelookup.indices) {
            val mc = mergelookup[t] as Mergedcells
            mc.initCells(this)
        }
    }

    internal fun initHlinks() {
        val hlinklookup = workBook!!.hlinklookup
        for (t in hlinklookup.indices) {
            val hl = hlinklookup[t] as Hlink
            hl.initCells(this)
        }
    }

    /**
     * reads in the pivot cache storage and parses the pivot cache records <br></br>
     * pivot cache(s) are used by pivot tables as data source storage
     */
    internal fun initPivotCache() {
        if (leoFile!!.hasPivotCache()) {
            val pc = PivotCache() // grab any pivot caches
            try {
                pc.init(leoFile!!.directoryArray!!, this)
                workBook!!.pivotCache = pc
            } catch (e: Exception) {

            }

        }
    }

    /**
     * Closes the WorkBook and releases resources.
     */
    override fun close() {
        try {
            if (leoFile != null)
                leoFile!!.shutdown()
            leoFile = null
        } catch (e: Exception) {
            if (DEBUGLEVEL > 3)
                Logger.logWarn("Closing Document: " + toString() + " failed: "
                        + e.toString())
        }

        if (workBook != null)
            workBook!!.close() // clear out object refs to release memory
        workBook = null
        factory = null
        name = null
        sheethandles = null
        // Runtime.getRuntime().gc();
    }

    @Throws(Throwable::class)
    protected fun finalize() {
        close()
    }

    override fun reset() {
        initFromFile(File(leoFile!!.fileName))
    }

    /**
     * returns the handle to a WorkSheet by number.
     *
     *
     * Sheet 0 is the first Sheet.
     *
     * @param index of worksheet (ie: 0)
     * @return WorkSheetHandle the WorkSheet
     * @throws WorkSheetNotFoundException if the specified WorkSheet is not found in the WorkBook.
     */
    @Throws(WorkSheetNotFoundException::class)
    override fun getWorkSheet(sheetnum: Int): WorkSheetHandle {
        val st = workBook!!.getWorkSheetByNumber(sheetnum)
        if (sheethandles!![st.sheetName] != null)
            return sheethandles!![st.sheetName]
        else {
            val shth = WorkSheetHandle(st, this)
            sheethandles!![st.sheetName] = shth
            return shth
        }
    }

    /**
     * returns the handle to a WorkSheet by name.
     *
     * @param String name of worksheet (ie: "Sheet1")
     * @return WorkSheetHandle the WorkSheet
     * @throws WorkSheetNotFoundException if the specified WorkSheet is not found in the WorkBook.
     */
    @Throws(WorkSheetNotFoundException::class)
    override fun getWorkSheet(handstr: String): WorkSheetHandle? {
        if (sheethandles!![handstr] != null) {
            return if (workBook!!.getWorkSheetByName(handstr) != null)
                sheethandles!![handstr]
            else
                throw WorkSheetNotFoundException("$handstr not found")
        }
        if (factory != null) {
            val bs = workBook!!.getWorkSheetByName(handstr)
            if (bs != null) {
                bs.workBook = this.workBook
                val ret = WorkSheetHandle(bs, this)
                sheethandles!![handstr] = ret
                return ret
            } else {
                throw WorkSheetNotFoundException(handstr)
            }
        }
        throw WorkSheetNotFoundException(
                "Cannot find WorkSheet $handstr")
    }

    /**
     * Set Encoding mode of new Strings added to file.
     *
     *
     * OpenXLS has 3 modes for handling the internal encoding of String data that is
     * added to the file.
     *
     *
     * OpenXLS can save space in the file if it knows that all characters in your
     * String data can be represented with a single byte (Compressed.)
     *
     *
     * If your String contains characters which need 2 bytes to represent (such as
     * Eastern-language characters) then it needs to be stored in an uncompressed
     * Unicode format.
     *
     *
     * OpenXLS can either automatically detect the mode for each String, or you can
     * set it explicitly. The auto mode is the most flexible but requires processing
     * overhead.
     *
     *
     * Default mode is WorkBookHandle.STRING_ENCODING_AUTO.
     *
     *
     * Valid Modes Are:
     *
     *
     * WorkBookHandle.STRING_ENCODING_AUTO Use if you are adding mixed Unicode and
     * non-unicode Strings and can accept the performance hit -slowest String adds
     * -optimal file size for mixed Strings WorkBookHandle.STRING_ENCODING_UNICODE
     * Use if all of your new Strings are Unicode - faster than AUTO -faster than
     * AUTO -largest file size WorkBookHandle.STRING_ENCODING_COMPRESSED Use if all
     * of your new Strings are non-Unicode and can have high-bytes compressed
     * -faster than AUTO -smallest file size
     *
     * @param int String Encoding Mode
     */
    override fun setStringEncodingMode(mode: Int) {
        workBook!!.setStringEncodingMode(mode)
    }

    /**
     * Set Duplicate String Handling Mode.
     *
     * <pre>
     * The Duplicate String Mode determines the behavior of
     * the String table when inserting new Strings.
     *
     * The String table shares a single entry for multiple
     * Cells containing the same string.  When multiple Cells
     * have the same value, they share the same underlying string.
     *
     * Changing the value of any one of the Cells will change
     * the value for any Cells sharing that reference.
     *
     * For this reason, you need to determine
     * the handling of new strings added to the sheet that
     * are duplicates of strings already in the table.
     *
     * If you will be changing the values of these
     * new Cells, you will need to set the Duplicate
     * String Mode to ALLOWDUPES.  If the string table
     * encounters a duplicate entry being added, it
     * will insert a duplicate that can then be subsequently
     * changed without affecting the other duplicate Cells.
     *
     * Valid Modes Are:
     *
     * WorkBookHandle.ALLOWDUPES - faster String inserts, larger file sizes,  changing Cells has no effect on dupe Cells
     *
     * WorkBookHandle.SHAREDUPES - slower inserts, dupe smaller file sizes, Cells share changes
    </pre> *
     *
     * @param int Duplicate String Handling Mode
     */
    override fun setDupeStringMode(mode: Int) {
        workBook!!.setDupeStringMode(mode)
    }

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chartname
     * @param sheetname
     */
    @Throws(ChartNotFoundException::class, WorkSheetNotFoundException::class)
    override fun copyChartToSheet(chartname: String, sheetname: String) {
        workBook!!.copyChartToSheet(chartname, sheetname)
    }

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chart
     * @param sheet
     */
    @Throws(ChartNotFoundException::class, WorkSheetNotFoundException::class)
    override fun copyChartToSheet(chart: ChartHandle, sheet: WorkSheetHandle) {
        workBook!!.copyChartToSheet(chart.title, sheet.sheetName)
    }

    /**
     * Copy (duplicate) a worksheet in the workbook and add it to the end of the
     * workbook with a new name
     *
     * @param String the Name of the source worksheet;
     * @param String the Name of the new (destination) worksheet;
     * @return the new WorkSheetHandle
     */
    @Throws(WorkSheetNotFoundException::class)
    override fun copyWorkSheet(SourceSheetName: String, NewSheetName: String): WorkSheetHandle? {
        try {
            workBook!!.copyWorkSheet(SourceSheetName, NewSheetName)
        } catch (e: Exception) {
            throw WorkBookException("Failed to copy WorkSheet: "
                    + SourceSheetName + ": " + e.toString(),
                    WorkBookException.RUNTIME_ERROR)
        }

        workBook!!.refTracker!!.clearPtgLocationCaches(NewSheetName)
        // update the merged cells (requires a WBH, that's why it's
        // here)
        val wsh = this.getWorkSheet(NewSheetName)
        if (wsh != null) {
            val mc = wsh.mysheet!!.mergedCellsRecs
            for (i in mc.indices) {
                val mrg = mc[i] as Mergedcells
                mrg?.initCells(this)
            }
            // now conditional formats
            /*
             * mc = wsh.getMysheet().getConditionalFormats(); for (int
             * i=0;i<mc.size();i++)
             * { Condfmt mrg = (Condfmt)mc.get(i); if (mrg !=
             * null)mrg.initCells(this); }
             */
        }
        return wsh
    }

    /**
     * Forces immediate recalculation of every formula in the workbook.
     *
     * @throws FunctionNotSupportedException if an unsupported function is used by any formula in the workbook
     * @see .forceRecalc
     * @see .recalc
     */
    override fun calculateFormulas() {
        markFormulasDirty()
        recalc()
    }

    /**
     * Marks every formula in the workbook as needing a recalc. This method does not
     * actually calculate formulas, for that use [.recalc].
     */
    fun markFormulasDirty() {
        val formulas = workBook!!.formulas
        for (idx in formulas.indices)
            formulas[idx].clearCachedValue()
    }

    /**
     * Recalculates all dirty formulas in the workbook immediately.
     *
     *
     * You generally need not call this method. Dirty formulas will automatically be
     * recalculated when their values are queried. This method is only useful for
     * forcing calculation to occur at a certain time. In the case of functions such
     * as NOW() whose value is volatile the formula will still be recalculated every
     * time it is queried.
     *
     * @throws FunctionNotSupportedException if an unsupported function is used by any formula in the workbook
     * @see .markFormulasDirty
     */
    fun recalc() {
        val calcmode = workBook!!.calcMode
        workBook!!.calcMode = WorkBook.CALCULATE_AUTO // ensure referenced functions are
        // calcualted as necesary!
        val formulas = workBook!!.formulas
        for (idx in formulas.indices) {
            try {
                formulas[idx].clearCachedValue()
                formulas[idx].calculate()
            } catch (fe: FunctionNotSupportedException) {
                Logger.logErr("WorkBookHandle.recalc:  Error calculating Formula $fe")
            }

        }
        // KSC: Clear out lookup caches!
        this.workBook!!.refTracker!!.clearLookupCaches()
        workBook!!.calcMode = calcmode // reset
    }

    /**
     * Removes all of the WorkSheets from this WorkBook.
     *
     *
     * Bytes streamed from this WorkBook will create invalid Spreadsheet files
     * unless a WorkSheet(s) are added to it.
     *
     *
     * NOTE: A WorkBook with no sheets is *invalid* and will not open in Excel. You
     * must add sheets to this WorkBook for it to be valid.
     */
    override fun removeAllWorkSheets() {

        try {
            val ob = this.workBook!!.tabID!!.tabIDs[0]
            this.workBook!!.tabID!!.tabIDs.removeAllElements()
            this.workBook!!.tabID!!.tabIDs.add(ob)
            this.workBook!!.tabID!!.updateRecord()
        } catch (ex: Exception) {
        }

        val ws = this.workSheets
        try {
            for (x in ws!!.indices) {
                try {
                    ws[x].remove()
                } catch (e: WorkBookException) {
                }
                // ignore the invalid WorkBook problem
            }
        } catch (e: Exception) {
            // in case sheets already gone...
        }

        this.sheethandles!!.clear()
        this.workBook!!.closeSheets() // replaced below with this

        /*
         * WHY ARE WE DOING THIS??? // init new book // save records
         * then reset to avoid
         * ByteStreamer.stream records expansion Object[] recs=
         * this.getWorkBook().getStreamer().getBiffRecords();
         *
         * // keep Excel 2007 status boolean isExcel2007=
         * this.getIsExcel2007();
         * WorkBookHandle ret = new WorkBookHandle(this.getBytes());
         * ret.setIsExcel2007(isExcel2007);
         *
         * this.getWorkBook().getStreamer().setBiffRecords(Arrays.
         * asList(recs));
         * this.mybook = ret.getWorkBook(); /
         **/

    }

    /**
     * Inserts a worksheet from a Source WorkBook.
     *
     * @param sourceBook      - the WorkBook containing the sheet to copy
     * @param sourceSheetName - the name of the sheet to copy
     * @param destSheetName   - the name of the new sheet in this workbook
     * @throws WorkSheetNotFoundException
     */
    @Deprecated("- use addWorkSheet(WorkSheetHandle sht, String NewSheetName){")
    @Throws(WorkSheetNotFoundException::class)
    override fun addSheetFromWorkBook(sourceBook: WorkBookHandle, sourceSheetName: String, destSheetName: String): Boolean {
        return this.addWorkSheet(sourceBook
                .getWorkSheet(sourceSheetName)!!, destSheetName) != null
    }

    /**
     * Inserts a worksheet from a Source WorkBook. Brings all string data and
     * formatting information from the source workbook.
     *
     *
     * Be aware this is programmatically creating a large amount of new formatting
     * information in the destination workbook. A higher performance option will
     * usually be using getNoSheetWorkbook and addSheetFromWorkBook.
     *
     * @param sourceBook      - the WorkBook containing the sheet to copy
     * @param sourceSheetName - the name of the sheet to copy
     * @param destSheetName   - the name of the new sheet in this workbook
     * @throws WorkSheetNotFoundException
     */
    @Deprecated("- use addWorkSheet(WorkSheetHandle sht, String NewSheetName){")
    @Throws(WorkSheetNotFoundException::class)
    fun addSheetFromWorkBookWithFormatting(sourceBook: WorkBookHandle, sourceSheetName: String, destSheetName: String): Boolean {
        return this.addWorkSheet(sourceBook
                .getWorkSheet(sourceSheetName)!!, destSheetName) != null
    }

    /**
     * Inserts a WorkSheetHandle from a separate WorkBookhandle into the current
     * WorkBookHandle.
     *
     *
     * copies charts, images, formats from source workbook
     *
     *
     * Worksheet will be the same name as in the source workbook. To add a custom
     * named worksheet use the addWorkSheet(WorkSheetHandle, String sheetname)
     * method
     *
     * @param WorkSheetHandle the source WorkSheetHandle;
     */
    fun addWorkSheet(sourceSheet: WorkSheetHandle): WorkSheetHandle? {
        return this.addWorkSheet(sourceSheet, sourceSheet.sheetName)
    }

    /**
     * Inserts a WorkSheetHandle from a separate WorkBookhandle into the current
     * WorkBookHandle.
     *
     *
     * copies charts, images, formats from source workbook
     *
     * @param WorkSheetHandle the source WorkSheetHandle;
     * @param String          the Name of the new (destination) worksheet;
     */
    override fun addWorkSheet(sourceSheet: WorkSheetHandle, NewSheetName: String): WorkSheetHandle? {
        sourceSheet.sheet!!.populateForTransfer() // copy all formatting +
        // images for this sheet
        val chts = sourceSheet.sheet!!.charts
        for (i in chts.indices) {
            val cxi = chts[i] as Chart
            cxi.populateForTransfer()
        }
        val bao = sourceSheet.serialBytes
        try {
            workBook!!.addBoundsheet(bao, sourceSheet
                    .sheetName, NewSheetName, StringTool
                    .stripPath(sourceSheet.workBook!!
                            .name), true)
            val wsh = this.getWorkSheet(NewSheetName)
            if (wsh != null) {
                val mc = wsh.mysheet!!.mergedCellsRecs
                for (i in mc.indices) {
                    val mrg = mc[i] as Mergedcells
                    mrg?.initCells(this)
                }
            }

            return wsh
        } catch (e: Exception) {
            throw WorkBookException(
                    "Failed to copy WorkSheet: $e",
                    WorkBookException.RUNTIME_ERROR)
        }

    }

    /**
     * Utility method to copy a format handle from a separate WorkBookHandle to this
     * WorkBookHandle
     *
     * @param externalFormat - FormatHandle from an external WorkBookHandle
     * @return
     */
    fun transferExternalFormatHandle(externalFormat: FormatHandle): FormatHandle {
        val xf = externalFormat.xf
        val newHandle = FormatHandle(this)
        newHandle.addXf(xf!!)
        return newHandle
    }

    /**
     * Creates a new Chart and places it at the end of the workbook
     *
     * @param String the Name of the newly created Chart
     * @return the new ChartHandle
     */
    fun createChart(name: String, wsh: WorkSheetHandle?): ChartHandle? {
        if (wsh == null) {
            // this is a sheetless chart - TODO:
        }
        /*
         * a chart needs a supbook, externsheet, & MSO object in the
         * book stream. I
         * think this is due to the fact that the referenced series
         * are usually stored
         * in the fashon 'Sheet1!A4:B6' The sheet1 reference
         * requires a supbook, though
         * the reference is internal.
         */

        try {
            val ois = ObjectInputStream(
                    ByteArrayInputStream(prototypeChart!!))
            var newchart = ois.readObject() as Chart
            newchart.workBook = this.workBook
            if (this.isExcel2007)
                newchart = OOXMLChart(newchart, this)
            workBook!!.addPreChart()
            workBook!!.addChart(newchart, name, wsh!!.sheet)
            /*
             * add font recs if nec: for the default chart: default
             * chart text fonts are # 5
             * & 6 title # 7 axis # 8
             */
            val bs = ChartHandle(newchart, this)
            var nfonts = workBook!!.numFonts
            while (nfonts < 8) { // ensure
                val f = Font("Arial", Font.PLAIN, 200)
                workBook!!.insertFont(f)
                nfonts++
            }
            var f = workBook!!.getFont(8) // axis title font
            if (f.toString() == "Arial,400,200 java.awt.Color[r=0,g=0,b=0] font style:[falsefalsefalsefalse00]") {
                // it's default text font -- change to default axis title
                // font
                f = Font("Arial", Font.BOLD, 240)
                bs.axisFont = f
            }
            f = workBook!!.getFont(7) // chart title font
            if (f.toString() == "Arial,400,200 java.awt.Color[r=0,g=0,b=0] font style:[falsefalsefalsefalse00]") {
                // it's default text font -- change to default title font
                f = Font("Arial", Font.BOLD, 360)
                bs.titleFont = f
            }
            bs.removeSeries(0) // remove the "dummied" series
            bs.setAxisTitle(ChartHandle.XAXIS, null) // remove default axis
            // titles, if any
            bs.setAxisTitle(ChartHandle.YAXIS, null) // ""
            return bs
        } catch (e: Exception) {
            Logger.logErr("Creating New Chart: $name failed: $e")
            return null
        }

    }

    /**
     * delete an existing chart of the workbook
     *
     * @param chartname
     */
    @Throws(ChartNotFoundException::class)
    fun deleteChart(chartname: String, wsh: WorkSheetHandle) {
        try {
            workBook!!.deleteChart(chartname, wsh.sheet)
        } catch (e: ChartNotFoundException) {
            throw ChartNotFoundException(
                    "Removing Chart: $chartname failed: $e")
        } catch (e: Exception) {
            Logger.logErr("Removing Chart: $chartname failed: $e")
        }

    }

    /**
     * Creates a new worksheet and places it at the specified position. The new
     * sheet will be inserted before the sheet currently at the given index. If the
     * given index is higher than the last index currently in use, the sheet will be
     * added to the end of the workbook and will receive an index one higher than
     * that of the current final sheet. If the given index is negative it will be
     * interpreted as 0.
     *
     * @param name     the name of the newly created worksheet
     * @param sheetpos the index at which the sheet should be inserted
     * @return the new WorkSheetHandle
     */
    fun createWorkSheet(name: String, sheetpos: Int): WorkSheetHandle {
        var sheetpos = sheetpos
        if (sheetpos > this.numWorkSheets)
            sheetpos = this.numWorkSheets
        if (sheetpos < 0)
            sheetpos = 0

        val s = this.createWorkSheet(name)
        s!!.tabIndex = sheetpos
        return s
    }

    /**
     * Creates a new worksheet and places it at the end of the workbook.
     *
     * @param name the name of the newly created worksheet
     * @return the new WorkSheetHandle
     */
    override fun createWorkSheet(name: String): WorkSheetHandle? {
        try {
            this.getWorkSheet(name)
            throw WorkBookException(
                    "Attempting to add worksheet with duplicate name. " + name
                            + " already exists in " + this.toString(),
                    WorkBookException.RUNTIME_ERROR)
        } catch (ex: WorkSheetNotFoundException) {
            // good!
        }

        var bo: Boundsheet? = null
        try {
            val ois = ObjectInputStream(
                    ByteArrayInputStream(prototypeSheet!!))
            bo = ois.readObject() as Boundsheet
            workBook!!.addBoundsheet(bo, null, name, null, false)
            try {
                val bs = this.getWorkSheet(name)
                if (this.workBook!!.numWorkSheets == 1) {
                    bs!!.selected = true // it's the only sheet so select!
                } else {
                    bs!!.selected = false
                }
                return bs
            } catch (e: WorkSheetNotFoundException) {
                Logger.logWarn("Creating New Sheet: $name failed: $e")
                return null
            }

        } catch (e: Exception) {
            Logger.logWarn("Error loading prototype sheet: $e")
            return null
        }

    }

    fun initSharedFormulas() {
        // TODO Auto-generated method stub

    }

    companion object {
        /**
         * A Writer to which a record dump should be written on input. This is used by
         * the dumping code in WorkBookFactory.
         */
        var dump_input: Writer? = null

        /**
         * Format constant for BIFF8 (Excel '97-2007).
         */
        val FORMAT_XLS = 100

        /**
         * Format constant for normal OOXML (Excel 2007).
         */
        val FORMAT_XLSX = 101

        /**
         * Format constant for macro-enabled OOXML (Excel 2007).
         */
        val FORMAT_XLSM = 102

        /**
         * Format constant for OOXML template (Excel 2007).
         */
        val FORMAT_XLTX = 103

        /**
         * Format constant for macro-enabled OOXML template (Excel 2007).
         */
        val FORMAT_XLTM = 104

        private var protobook: ByteArray? = null
        private var protochart: ByteArray? = null
        private var protosheet: ByteArray? = null
        var simpledateformat = java.text.SimpleDateFormat()    // static
        // to
        // reuse

        /**
         * How many recursion levels to allow formulas to be calculated before throwing
         * a circular reference error
         */
        var RECURSION_LEVELS_ALLOWED = 107

        /**
         * This is
         */
        var CONVERTMULBLANKS = "deprecated"

        protected val prototypeBook: ByteArray?
            @Throws(IOException::class)
            get() {
                if (protobook == null)

                    protobook = ResourceLoader
                            .getBytesFromJar("/io/starter/OpenXLS/templates/prototysspe.ser")

                return protobook
            }

        protected val prototypeSheet: ByteArray?
            get() {
                if (protosheet == null)
                    try {
                        val bookhandle = WorkBookHandle()
                        val book = bookhandle.workBook
                        val sheet = book!!.getWorkSheetByNumber(0)
                        protosheet = sheet.sheetBytes
                    } catch (e: Exception) {
                        throw RuntimeException(e)
                    }

                return protosheet
            }

        protected val prototypeChart: ByteArray?
            get() {
                if (protochart == null) {
                    try {
                        val bookbytes = ResourceLoader
                                .getBytesFromJar("/io/starter/OpenXLS/templates/prototypechart.ser")
                        val chartBook = WorkBookHandle(bookbytes)
                        val ch = chartBook.charts[0]
                        protochart = ch.serialBytes
                        return protochart
                    } catch (e: IOException) {
                        Logger.logErr("Unable to get default chart bytes")
                    }

                }
                return protochart
            }

        /**
         * Set the recursion levels allowed for formulas calculated in this workbook
         * before a circular reference is thrown.
         *
         *
         * Default setting is 250 levels of recursion
         *
         * @param recursion_allowed
         */
        fun setFormulaRecursionLevels(recursion_allowed: Int) {
            RECURSION_LEVELS_ALLOWED = recursion_allowed
        }
    }
}
/**
 * Constructor which takes the XLS file name(
 *
 * @param String filePath the name of the XLS file to read
 */