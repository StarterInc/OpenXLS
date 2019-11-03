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

import io.starter.formats.XLS.*
import io.starter.formats.XLS.formulas.Ptg
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONArray
import org.json.JSONObject

import io.starter.OpenXLS.JSONConstants.JSON_CELL
import io.starter.OpenXLS.JSONConstants.JSON_CELLS

/**
 * The NameHandle provides access to a Named Range and its Cells.<br></br>
 * <br></br>
 * Use the NameHandle to work with individual Named Ranges in an XLS file.<br></br>
 * <br></br>  <br></br>
 * With a NameHandle you can:
 * <br></br><br></br>
 * <blockquote>
 * get a handle to the Cells in a Name<br></br>
 * set the default formatting for a Name
</blockquote> *
 * <br></br>
 * <br></br>
 *
 * @see WorkBookHandle
 *
 * @see WorkSheetHandle
 *
 * @see FormulaHandle
 */
class NameHandle {

    private var myName: Name? = null
    private var mybook: WorkBook? = null
    private val DEBUGLEVEL = -1
    private var createblanks = false
    private var initialRange: CellRange? = null

    /**
     * Returns a handle to the object (either workbook or sheet) that is scoped
     * to the name record
     *
     *
     * Default scope is a WorkBookHandle, else the WorkSheetHandle is returned.
     *
     * @return the scope of the name
     */
    /**
     * Set the scope of this name to that of the handle passed in.
     *
     *
     * This can either be a WorkbookHandle or a WorksheetHandle
     *
     *
     * note: this will only be functional for the workbook that the name is contained in,
     * you cannot change the scope to a different Document with this method.
     *
     * @param scope Workbookhandle or WorksheetHandle
     */
    // this really shouldnt happen unless you are passing a scope in from a different workbook
    var scope: Handle?
        @Throws(WorkSheetNotFoundException::class)
        get() {
            val itab = myName!!.itab.toInt()
            return if (itab == 0) mybook else mybook!!.getWorkSheet(itab - 1)
        }
        set(scope) {
            var newitab = 0
            if (scope is WorkSheetHandle) {
                newitab = scope.sheetNum + 1
            }
            try {
                myName!!.setNewScope(newitab)
            } catch (e: WorkSheetNotFoundException) {
                Logger.logErr("ERROR: setting new scope on name: $e")
            }

        }


    /**
     * Return an XML representation of this name record
     */
    // add the sheetname
    // TODO: why no sheet defined for name???
    // it's possible that name expression can have quotes and other non-compliant characters
    // deal with snafu character names
    // if not workbook-scoped, add sheet scope for later retrieval
    val xml: String
        get() {
            val retXML = StringBuffer()
            var nmv = expressionString
            if (nmv.indexOf("=!") > -1) {
                val bs = myName!!.sheet
                if (bs != null) {
                    nmv = bs.sheetName + "!" + nmv
                } else {
                    nmv = StringTool.replaceChars("!", nmv, "")
                }
            }
            nmv = io.starter.toolkit.StringTool.convertXMLChars(nmv)

            var nmx = name
            nmx = io.starter.toolkit.StringTool.convertXMLChars(nmx)

            if (nmx.length == 1) {
                if (Character.isLetterOrDigit(nmx[0])) {
                    Logger.logInfo("NameHandle getting XML for name: $nmx")
                } else {
                    nmx = "#NAME"
                }
            }

            if (nmv.startsWith("="))
                nmv = nmv.substring(1)
            retXML.append("		<NamedRange Name=\"")
            retXML.append(nmx)
            retXML.append("\" RefersTo=\"")
            retXML.append(nmv)
            if (myName!!.itab.toInt() != 0) {
                retXML.append("\" Scope=\"")
                retXML.append(myName!!.itab.toInt())
            }
            retXML.append("\"/>")
            return retXML.toString()
        }


    /**
     * @return String XML rep of all the cells referenced by this range
     */
    // add the sheetname
    // TODO: why no sheet defined for name???
    // deal with snafu character names
    val expandedXML: String
        get() {
            var nmv = expressionString
            if (nmv.indexOf("=!") > -1) {
                val bs = myName!!.sheet
                if (bs != null) {
                    nmv = bs.sheetName + "!" + nmv
                } else {
                    nmv = StringTool.replaceChars("!", nmv, "")
                }
            }
            var nmx = name
            nmx = io.starter.toolkit.StringTool.convertXMLChars(nmx)

            if (nmx.length == 1) {
                if (Character.isLetterOrDigit(nmx[0])) {
                    Logger.logInfo("NameHandle getting XML for name: $nmx")
                } else {
                    nmx = "#NAME"
                }
            }

            if (nmv.startsWith("="))
                nmv = nmv.substring(1)
            val retXML = StringBuffer()
            retXML.append("\t<NamedRange Name=\"")
            retXML.append(nmx)
            retXML.append("\" RefersTo=\"")
            retXML.append(nmv)
            retXML.append("\">\n")
            try {
                val cells = cells

                for (i in cells!!.indices) {
                    retXML.append("\t\t" + cells[i].xml + "\n")
                }
            } catch (ex: CellNotFoundException) {
                Logger.logErr("NameHandle.getExpandedXML failed: ", ex)
            }

            retXML.append("</NamedRange>\n")

            return retXML.toString()
        }

    /**
     * get the referenced named cell range as string in standard excel syntax including sheetname,
     *
     *
     * for instance "Sheet1!A1:Z255"
     */
    /**
     * set the referenced cells for the named range
     *
     *
     * this reference should be in the standard excel syntax including sheetname,
     * for instance "Sheet1!A1:Z255"
     */
    // if there is no sheetname, and no sheet, then this is an unusable name and cannot be added
    var location: String?
        get() {
            try {
                return myName!!.location
            } catch (e: Exception) {
                Logger.logErr("Error getting named range location$e")
                return null
            }

        }
        set(strloc) {
            var strloc = strloc
            if (strloc.indexOf("!") == -1 && strloc.indexOf("=") == -1) {
                if (this .2 DSheetName == null)
                {
                    this.remove()
                    throw IllegalArgumentException("Named Range References must include a Sheet name.")
                }
                else
                strloc = this.2DSheetName+"!"+strloc
            }
            try {
                myName!!.location = strloc
            } catch (e: FunctionNotSupportedException) {
                Logger.logWarn("NameHandle.setLocation :$strloc failed: $e")
            }

        }

    /**
     * returns the name String for the range definition
     *
     * @return String definition name
     */
    /**
     * set the name String for the range definition
     *
     * @param String definition name
     */
    var name: String
        get() = myName!!.name
        set(newname) {
            myName!!.name = newname
        }

    /**
     * returns the name's formula String for the range definition
     *
     * @return String the expression string
     */
    val expressionString: String
        get() {
            try {
                return myName!!.expressionString
            } catch (e: Exception) {
                if (DEBUGLEVEL > -1) Logger.logWarn("Could not parse expression string for name: " + this.name)
            }

            return "#ERR"
        }

    /**
     * gets the array of Cells in this Name
     *
     * @return Cell[] all Cells defined in this Name
     */
    // ignores merged ranges here
    // ignores merged ranges here
    // ignores merged ranges here
    //	sbx.append("</Row>");
    val cellRangeXML: String
        get() {
            val sbx = StringBuffer()
            try {
                val celx = this.cells
                var rowhold: RowHandle? = null
                for (x in celx!!.indices) {
                    val rx = celx[x].row
                    if (x == 0) {
                        rowhold = rx
                        sbx.append("<Row Number='" + rowhold!!.rowNumber + "'>")
                        sbx.append(celx[x].xml)
                    } else if (rowhold!!.rowNumber == rx.rowNumber) {
                        sbx.append(celx[x].xml)
                    } else {
                        sbx.append("</Row>")
                        rowhold = rx
                        sbx.append("<Row Number='" + rowhold!!.rowNumber + "'>")
                        sbx.append(celx[x].xml)
                    }
                }
                sbx.append("</Row>")
            } catch (ex: CellNotFoundException) {
                Logger.logErr("NameHandle.getCellRangeXML failed: ", ex)
            }

            return sbx.toString()
        }

    /**
     * gets the array of Cells in this Name
     *
     * @return Cell[] all Cells defined in this Name
     */
    // first get the ptgArea3d from the parsed expression
    // for(int b = (cells.length-1);b>=0;b--)cellhandles.add(cells[b]);
    val cells: Array<CellHandle>?
        @Throws(CellNotFoundException::class)
        get() {
            if (false)
                return null
            try {
                val cellhandles = CompatibleVector()
                val rngz = this.cellRanges
                if (rngz != null) {
                    for (t in rngz.indices) {
                        try {
                            val cells = rngz[t].getCells()
                            for (b in cells!!.indices)
                                if (cells[b] != null)
                                    cellhandles.add(cells[b])
                        } catch (ex: Exception) {
                            Logger.logWarn("Could not get cells for range: " + rngz[t])
                        }

                    }
                } else {
                    return null
                }
                val ret = arrayOfNulls<CellHandle>(cellhandles.size)
                return cellhandles.toTypedArray() as Array<CellHandle>
            } catch (e: Exception) {
                if (e is CellNotFoundException)
                    throw e
                throw CellNotFoundException(e.toString())
            }

        }


    /**
     * Get an Array of CellRanges, one per referenced WorkSheet.
     *
     *
     * If this method throws CellNotFoundExceptions, then you are
     * addressing a sparsely populated CellRange.
     *
     *
     * Use 'setCreateBlanks(true)' to populate these Cells and avoid
     * this error.
     *
     * @return
     * @throws Exception
     */
    // 20100217 KSC: try a better way (that can handle 3D refs and complex cell ranges)
    // may contain one or more ranges, separated by ","'s if complex
    // handle commas within quoted sheet names (TestExtenXLSEngine.TestQuotedSheetsWithCommansInNRs)
    // NOTE that sheetnames must be properly qualified for the split to work below
    // below cannot handle commas embedded within quotes
    //		String[] nranges= StringTool.splitString(loc, ",");	// TODO: can be another delimeter?
    // no sheetname
    val cellRanges: Array<CellRange>
        @Throws(Exception::class)
        get() {
            if (initialRange != null) {
                return arrayOf<CellRange?>(initialRange)
            }
            val loc = myName!!.location
            val sheets = this.referencedSheets
            val nranges = loc!!.split(",(?=([^'|\"]*'[^'|\"]*'|\")*[^'|\"]*$)".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            val ranges = arrayOfNulls<CellRange>(nranges.size)
            for (i in nranges.indices) {
                var r = nranges[i]
                if (r.indexOf("!") == -1) {
                    r = sheets[i].toString() + "!" + r
                }
                ranges[i] = CellRange(r, mybook, createblanks)
                ranges[i].setParent(this.myName)
            }

            return ranges
        }

    /**
     * Get WorkSheetHandles for all of the Boundsheets referenced in
     * this NameHandle.
     *
     * @return an array of WorkSheetHandles referenced in this Name
     */
    //first get the ptgArea3d from the parsed expression
    val referencedSheets: Array<WorkSheetHandle>
        @Throws(WorkSheetNotFoundException::class)
        get() {
            val bs = myName!!.boundSheets
                    ?: throw WorkSheetNotFoundException("Worksheet for Named Range: " + this.toString() + ":" + this.myName!!.expressionString)
            if (bs[0] == null)
                throw WorkSheetNotFoundException("Worksheet for Named Range: " + this.toString() + ":" + this.myName!!.expressionString)
            val ret = arrayOfNulls<WorkSheetHandle>(bs.size)
            for (x in ret.indices) {
                ret[x] = mybook!!.getWorkSheet(bs[x].toString())
            }
            return ret
        }

    /**
     * return the calculated value of this Name
     * if it contains a parsed Expression (Formula)
     *
     * @return
     * @throws FunctionNotSupportedException
     */
    val calculatedValue: Any
        @Throws(FunctionNotSupportedException::class)
        get() = myName!!.calculatedValue

    /**
     * Return a JSON object representing this name Handle.
     * name:'nameOfRange'
     * cellrange:'Sheet1!A1:B1'
     *
     * @return
     */
    val json: String
        get() = getJSON(false)

    /**
     * Returns a JSON object in the same format as [CellRange.getJSON].
     */
    val jsonCellRange: JSONObject
        get() {
            val theRange = JSONObject()
            val cells = JSONArray()
            try {
                val chandles = cells
                for (i in chandles.indices) {
                    val thisCell = chandles[i]
                    val result = JSONObject()

                    result.put(JSON_CELL, thisCell.jsonObject)
                    cells.put(result)
                }
                theRange.put(JSON_CELLS, cells)
            } catch (e: Exception) {
                Logger.logErr("Error getting NamedRange JSON: $e")
            }

            return theRange
        }

    /**
     * return the sheetname for a 2D named range
     *
     *
     * NOTE: Does not work for 3D ranges
     *
     * @return the Sheet name if this is a 2d range
     */
    val 2DSheetName:String?
    get()
    {
        try {
            return this.myName!!.boundSheets!![0].sheetName
        } catch (e: Exception) {
            try {
                return this.myName!!.sheet!!.sheetName
            } catch (ex: Exception) {
                return null
            }

        }

    }

    /**
     * Creates a new Named Range from a CellRange
     *
     * @param Name of the Range, the CellRange referenced by the Name
     */
    constructor(namestr: String, cr: CellRange) {
        mybook = cr.workBook
        initialRange = cr // cache
        myName = Name(mybook!!.workBook, namestr)
        this.name = namestr
        this.location = cr.toString()
    }

    /**
     * Create a NameHandle from an internal Name record
     *
     * @param c
     * @param myb
     */
    constructor(c: Name, myb: WorkBookHandle) {
        myName = c
        mybook = myb
    }

    /**
     * Create a new named range in the workbook
     * note that this does not set the actual range, just the workbook and name,
     * follow up with setLocation
     *
     * @param namestr
     * @param myb
     */
    @Deprecated("")
    constructor(namestr: String, myb: WorkBookHandle) {
        mybook = myb
        myName = Name(mybook!!.workBook, namestr)
        this.name = namestr
    }

    /**
     * Create a new named range in the workbook
     *
     * @param name     name that should be used to reference this named range
     * @param location rangeDef Range of the cells for this named range, in excel syntax including sheet name, ie "Sheet1!A1:D1"
     * @param book     WorkBookHandle to insert this named range into
     */
    constructor(name: String, location: String, book: WorkBookHandle) {
        mybook = book
        myName = Name(mybook!!.workBook, name)
        this.name = name
        this.location = location
        mybook!!.workBook.associateDereferencedNames(myName!!)
    }

    /**
     * sets the default format id for the Name's Cells
     *
     * @param int Format Id for all Cells in Name
     */
    fun setFormatId(i: Int) {
        // myName.setXFRecord(i);
        // TODO: why doesn't this work with myName?
        try {
            val cs = this.cells
            for (t in cs!!.indices) {
                cs[t].formatId = i
            }
        } catch (ex: CellNotFoundException) {
            Logger.logErr("NameHandle.setFormatId failed: ", ex)
        }

    }

    /**
     * deletes a row of cells in this named range, shifts subsequent
     * rows up.
     *
     * @param the column to use as unique index
     */
    @Throws(Exception::class)
    fun deleteRow(idxcol: Int) {
        val rngs = cellRanges
        if (rngs.size > 1) {
            throw WorkBookException("NamedRange.updateRow Object array failed: too many CellRanges.", WorkBookException.RUNTIME_ERROR)
        } else if (rngs.size == 0) {
            throw WorkBookException("NamedRange.updateRow Object array failed: zero CellRanges", WorkBookException.RUNTIME_ERROR)
        }

        //
        val shtx = rngs[0].sheet
        val x = rngs[0].firstcellrow


        val rx = rngs[0].rows
        val found = false
        // iterate, find, update
        for (t in rx.indices) {

        }
    }

    /**
     * update a row of cells in this named range
     *
     * @param an  array of Objects to update existing
     * @param the column to use as unique index
     */
    @Throws(Exception::class)
    fun updateRow(objarr: Array<Any>, idxcol: Int) {
        val rngs = cellRanges
        if (rngs.size > 1) {
            throw WorkBookException("NamedRange.updateRow Object array failed: too many CellRanges.", WorkBookException.RUNTIME_ERROR)
        } else if (rngs.size == 0) {
            throw WorkBookException("NamedRange.updateRow Object array failed: zero CellRanges", WorkBookException.RUNTIME_ERROR)
        }

        //
        val shtx = rngs[0].sheet
        val x = rngs[0].firstcellrow


        val rx = rngs[0].rows
        var found = false
        // iterate, find, update
        for (t in rx.indices) {

            val cx = rx[t].cells
            if (cx.size > idxcol) {
                if (cx[idxcol].stringVal!!.equals(objarr[idxcol].toString(), ignoreCase = true)) {
                    found = true
                    for (z in cx.indices) {
                        cx[z].`val` = objarr[z]
                    }
                }
            }
        }

        if (!found)
            this.addRow(objarr)
    }


    /**
     * add a row of cells to this named range
     *
     * @param an array of Objects to insert at last rown
     */
    @Throws(Exception::class)
    fun addRow(objarr: Array<Any>) {
        val rngs = cellRanges
        if (rngs.size > 1) {
            throw WorkBookException("NamedRange.add Object array failed: too many CellRanges.", WorkBookException.RUNTIME_ERROR)
        } else if (rngs.size == 0) {
            throw WorkBookException("NamedRange.add Object array failed: zero CellRanges", WorkBookException.RUNTIME_ERROR)
        }

        //
        val shtx = rngs[0].sheet

        val x = rngs[0].firstcellrow
        val cxx = shtx!!.insertRow(x, objarr, true)

        //for(int t=0;t<cxx.length;t++)
        //	if(cxx[t]!=null)rngs[0].addCellToRange(cxx[t]);
        rngs[0] = CellRange(this.myName!!.location, mybook, createblanks)
    }

    /**
     * add a cell to this named range
     *
     * @param cx
     */
    @Throws(Exception::class)
    fun addCell(cx: CellHandle) {
        val rngs = cellRanges
        if (rngs.size > 1) {
            throw WorkBookException("NamedRange.addCell failed -- more than one cell range defined in this NameHandle, cannot determine where to add cell.", WorkBookException.RUNTIME_ERROR)
        }
        rngs[0].addCellToRange(cx)
        location = rngs[0].getRange()
    }

    override fun toString(): String {
        return name
    }

    /**
     * removes this Named Range from the WorkBook.
     *
     * @return whether the removal was a success
     */
    fun remove(): Boolean {
        var success = false
        try {
            success = this.myName!!.workBook!!.removeName(myName)
        } catch (e: Exception) {
            return false
        }

        return success
    }

    /**
     * set whether the CellRanges referenced by the NameHandle
     * will add blank records to the WorkBook for any missing Cells
     * contained within the range.
     *
     * @param b set whether to create blank records for missing Cells
     */
    fun setCreateBlanks(b: Boolean) {
        createblanks = b
    }

    /**
     * gets the array of Cells in this Name
     *
     *
     * NOTE: this method variation also returns the Sheetname for the name record if not null.
     *
     *
     * Thus this method is limited to use with 2D ranges.
     *
     * @param fragment whether to enclose result in NameHandle tag
     * @return Cell[] all Cells defined in this Name
     */
    fun getCellRangeXML(fragment: Boolean): String {
        val sbx = StringBuffer()
        if (!fragment)
            sbx.append("<?xml version=\"1\" encoding=\"utf-8\"?>")
        sbx.append("<NameHandle Name=\"" + this.name + "\">")
        sbx.append(cellRangeXML)
        sbx.append("</NameHandle>")
        return sbx.toString()
    }

    /**
     * Sets the location lock on the Cell Reference at the
     * specified  location
     *
     *
     * Used to prevent updating of the Cell Reference when
     * Cells are moved.
     *
     * @param location of the Cell Reference to be locked/unlocked
     * @param lock     status setting
     * @return boolean whether the Cell Reference was found and modified
     */
    fun setLocationLocked(loc: String, l: Boolean): Boolean {
        var x = Ptg.PTG_LOCATION_POLICY_UNLOCKED
        if (l) x = Ptg.PTG_LOCATION_POLICY_LOCKED
        return myName!!.setLocationPolicy(loc, x)
    }


    /**
     * Return a JSON object representing this name Handle.
     * name:'nameOfRange'
     * cellrange:'Sheet1!A1:B1'
     * cells:celldata
     *
     * @param whether to return cell data
     * @return
     */
    fun getJSON(celldata: Boolean): String {
        val theNameHandle = JSONObject()
        try {
            theNameHandle.put("name", this.name)
            theNameHandle.put("cellrange", myName!!.location)

            if (celldata) {
                val ret = StringBuffer()
                val cx1 = cells
                var p = cx1!![0].rowNum
                for (x in cx1.indices) {
                    if (cx1[x] != null) {
                        if (cx1[x].rowNum != p) {
                            ret.append("\r\n")
                            p = cx1[x].rowNum
                        }
                        ret.append("'")
                        try {
                            ret.append(cx1[x].stringVal)
                        } catch (ex: Exception) { // handles empties
                            //
                        }

                        if (x != cx1.size - 1)
                            ret.append("',")
                        else
                            ret.append("'")
                    }
                }

                theNameHandle.put("celldata", ret.toString())
            }
        } catch (e: Exception) {
            Logger.logErr("Error creating JSON name handle: $e")
        }

        return theNameHandle.toString()
    }

}