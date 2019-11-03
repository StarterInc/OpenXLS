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

import io.starter.OpenXLS.*
import io.starter.formats.XLS.charts.Ai
import io.starter.formats.XLS.charts.Chart
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.Logger

import java.util.*

/**
 * This class is responsible for registering cell references (Ptgs) and managing
 * reference updates etc.
 *
 *
 *
 *
 * Here's the scenario:
 *
 *
 * as we parse workbooks, WorkBook.addFormula() extracts all PtgRefs and
 * puts them in various ReferenceTracker maps.
 *
 *
 * as cells are inserted/and or changed, referenced cells can be updated
 * after calling 'getAffectedCells'.
 *
 *
 * TODO: retrofit all methods to use sheet_map cache
 *
 *
 *
 *
 *
 *
 * All PtgRefs and Areas
 */
class ReferenceTracker {


    // the sheets allow for faster refs
    // each sheet contains a collection of rows.
    private var sheetMap: MutableMap<*, *> = HashMap()
    // store ptgNames
    private var nameRefs: MutableMap<*, *> = HashMap()

    // Database calc caches
    private var criteriaDBs: MutableMap<*, *> = HashMap()
    private var CollectionDBs: MutableMap<*, *> = HashMap()
    private var vlookups: MutableMap<*, *> = HashMap()
    private var crs: MutableCollection<*> = Vector()

    // VLOOKUPs and other lookups need to calc col ptgs
    private var lookupColsCache: MutableMap<*, *> = HashMap()

    val lookupColCache: Map<*, *>
        get() = lookupColsCache


    /**
     * @return Returns the CollectionDBs.
     */
    val listDBs: Map<*, *>
        get() = CollectionDBs


    /* CELL RANGE SECTION */

    /**
     * Returns an Array of the CellRanges existing in this WorkBook
     * specifically the Ranges referenced in Formulas, Charts, and
     * Named Ranges.
     *
     *
     * This is necessary to allow for automatic updating of references
     * when adding/removing/moving Cells within these ranges, as well
     * as shifting references to Cells in Formulas when Formula records
     * are moved.
     *
     * @return all existing Cell Range references used in Formulas, Charts, and Names
     */
    val cellRanges: Array<CellRange>
        get() {
            val ret = arrayOfNulls<CellRange>(this.crs.size)
            return crs.toTypedArray() as Array<CellRange>
        }

    /**
     * @return Returns the vlookups.
     */
    fun getVlookups(): Map<*, *> {
        return vlookups
    }


    /**
     * @return Returns the criteriaDBs.
     */
    fun getCriteriaDBs(): Map<*, *> {
        return criteriaDBs
    }


    /**
     * Blow out all the cacheing
     */
    fun clearCaches() {
        // Should we?
        // crPtgMap     =   new HashMap();
        // refPtgMap    =   new HashMap();

        // Databases
        criteriaDBs = HashMap()
        CollectionDBs = HashMap()
        vlookups = HashMap()
    }

    /**
     * clear out VLOOKUP and related function caches
     */
    fun clearLookupCaches() {
        lookupColsCache.clear()
        lookupColsCache = HashMap()
        criteriaDBs.clear()
        criteriaDBs = HashMap()
        CollectionDBs.clear()
        CollectionDBs = HashMap()
        vlookups.clear()
        vlookups = HashMap()
    }

    /**
     * Returns ALL formulas on the cellhandles sheet that reference this CellHandle.
     *
     *
     * Clears the cached value on said cells so a recalc will be forced upon any getVal method
     *
     *
     * Please note that these cells have already been calculated, so in order
     * to get their values without re-calculating them
     * Extentech suggests setting the book level non-calculation flag, ie
     *
     *
     * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT);
     *
     *
     * or
     *
     *
     * FormulaHandle.getCachedVal()
     *
     * @return Collection of of calculated cells
     */
    @Synchronized
    fun clearAffectedFormulaCellsOnSheet(cx: CellHandle, sheetname: String): Map<*, *> {
        val hm = clearAffectedFormulaCells(cx) as HashMap<*, *>
        val retmap = HashMap()
        val i = hm.keys.iterator()
        while (i.hasNext()) {
            val s = i.next() as String
            if (s.indexOf(sheetname) > -1)
                retmap.put(s, hm.get(s))
        }
        return retmap
    }


    /**
     * Returns a Collection Map of cells that are affected by formula
     * references to this CellHandle.
     *
     *
     * Clears the cached value on said cells and recalcs.
     *
     *
     * Please note that these cells are calculated before return.
     *
     *
     * To get their values without re-calculating them set the book level non-calculation flag, ie
     * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT);
     * or
     * FormulaHandle.getCachedVal()
     *
     * @return Map of calculated cells
     */
    @Synchronized
    fun clearAffectedFormulaCells(cx: CellHandle): Map<*, *> {
        return clearAffectedFormulaCells(cx.cell, HashMap())
    }

    /**
     * Returns a Collection Map of cells that are affected by formula
     * references to this CellHandle.
     *
     *
     * Clears the cached value on said cells and recalcs.
     *
     *
     * Please note that these cells are calculated before return.
     *
     *
     * To get their values without re-calculating them set the book level non-calculation flag, ie
     * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT);
     * or
     * FormulaHandle.getCachedVal()
     *
     * @return Map of calculated cells
     */
    @Synchronized
    fun clearAffectedFormulaCells(cx: BiffRec): Map<*, *> {
        return clearAffectedFormulaCells(cx, HashMap())
    }

    /**
     * Returns ALL formulas on the cellhandle's sheet that reference
     * record changedRec.
     *
     *
     * Clears the cached value on said cells so a recalc will be
     * forced upon any getVal method
     *
     *
     * Please note that these cells have not yet been calculated, so
     * in order pass a null value of sheetname in order to get all sheets
     *
     * @return Collection of of calculated cells
     */
    @Synchronized
    private fun clearAffectedFormulaCells(changedRec: BiffRec?, affectedCellHandles: MutableMap<*, *>?): MutableMap<*, *> {
        var affectedCellHandles = affectedCellHandles

        if (affectedCellHandles == null)
            affectedCellHandles = HashMap()

        val newRecSheetName = changedRec!!.sheet.sheetName
        // get ref collection for the sheet
        val ptgRefs = sheetMap.get(GenericPtg.qualifySheetname(newRecSheetName)) as TrackedPtgs
                ?: return affectedCellHandles    // now tracked ptgs are stored per sheet
        val parents = ptgRefs!!.getParents(changedRec!!)    // finds ALL parents affected by cell newRec
        while (parents.hasNext()) {
            val br = parents.next() as BiffRec
            val op = br.opcode
            if (op == XLSRecord.NAME) {
                val theName = (br as Name).nameA
                if (nameRefs.containsKey(theName)) {
                    // add all formulas that refer
                    val list = nameRefs.get(theName) as ArrayList<*> // gets the ptgname
                    for (i in list.indices) {
                        val ptgParent = (list[i] as Ptg).parentRec
                        if (ptgParent.opcode == XLSConstants.NAME)
                            continue // a Named Range referencing another named range ... will be caught later
                        val adr = ptgParent.sheet.sheetName + "!" + ptgParent.cellAddress
                        if (affectedCellHandles!!.get(adr) == null) {
                            ReferenceTracker.addRec(ptgParent, affectedCellHandles!!)
                            affectedCellHandles = clearAffectedFormulaCells(ptgParent, affectedCellHandles) // recurse parent formula and get cells it affects
                        }
                    }
                }
            } else if (op == XLSConstants.CONDFMT || op == XLSConstants.AI) {    // ignore since these records are not themselves referenced
            } else if (op == XLSConstants.SHRFMLA) {     // Shared Formula references are now reference-tracked; to find specific formula affected, use Shrfmla.getAffected
                val sh = br as Shrfmla
                val f = sh.getAffected(changedRec!!)
                if (f != null) {
                    val adr = f!!.sheet!!.sheetName + "!" + f!!.cellAddress
                    if (!affectedCellHandles!!.containsKey(adr)) {
                        ReferenceTracker.addRec(f!!, affectedCellHandles!!)
                        affectedCellHandles = clearAffectedFormulaCells(f, affectedCellHandles)    // recurse parent formula and get cells it affects
                    }
                }
            } else {  // regular Formula
                if (br.sheet != null) {
                    val adr = br.sheet.sheetName + "!" + br.cellAddress
                    if (!affectedCellHandles!!.containsKey(adr)) {
                        ReferenceTracker.addRec(br, affectedCellHandles!!)
                        affectedCellHandles = clearAffectedFormulaCells(br, affectedCellHandles)    // recurse parent formula and get cells it affects
                    }
                } // ignore no sheet
            }
        }
        return affectedCellHandles
    }

    /**
     * retrieve all chart-related (==Ai) references to the particular cell
     *
     * @param newRec cell to lookup references
     * @return list of Ai records that reference cell
     */
    fun getChartReferences(newRec: BiffRec): List<Ai> {
        val newRecSheetName = newRec.sheet.sheetName
        val ret = ArrayList()
        // get ref collection for the sheet
        val ptgRefs = sheetMap.get(GenericPtg.qualifySheetname(newRecSheetName)) as TrackedPtgs
                ?: return ret    // now tracked ptgs are stored per sheet
        val parents = ptgRefs!!.getParents(newRec)    // finds ALL parents affected by cell newRec
        while (parents.hasNext()) {
            val br = parents.next() as BiffRec
            val op = br.opcode
            if (op == XLSConstants.AI)
                ret.add(br as Ai)
        }
        return ret
    }

    /**
     * Add to the collection of PtgNames for referenceTracker
     */
    fun addPtgNameReference(p: PtgName) {
        val name = p.textString.toUpperCase() // case-insensitive
        var refs: Any? = nameRefs[name]
        if (refs == null) {
            refs = ArrayList()
            (refs as ArrayList<*>).add(p)
            nameRefs[name] = refs
        } else {
            val ptgNames = refs as ArrayList<*>?
            if (!ptgNames!!.contains(p)) ptgNames!!.add(p)
        }
    }


    /**
     * adds a cellrange Ptg (Area, Area3d etc.) to be tracked
     *
     *
     * These records are stored in a row based lookup,
     * this row map is looked up off actual row number, ie row 1 is get(1).
     *
     * @param cr
     * @return
     */
    fun addCellRange(ptgRef: Ptg): Ptg {
        // system setting to disable ref tracking...
        val trackprop = System.getProperty(WorkBookHandle.REFTRACK_PROP)
        if (trackprop != null)
            if (trackprop == "false")
                return ptgRef

        if (ptgRef !is PtgRef)
            return ptgRef
        if (ptgRef is PtgAreaErr3d || ptgRef is PtgRefErr3d || ptgRef is PtgRefErr)
            return ptgRef

        var sheetname: String? = ""
        try {
            try {
                sheetname = (ptgRef as PtgRef).sheetName
                sheetname = GenericPtg.qualifySheetname(sheetname)
            } catch (ex: Exception) {
                sheetname = "WorkBookRanges"
            }

            // fast fail erroneous Sheet refs
            if (sheetname == "#REF!") {
                return ptgRef
            }

            var ptgs: TrackedPtgs? = sheetMap.get(sheetname) as TrackedPtgs // now tracked ptgs are stored per sheet not per row
            if (ptgs == null) {
                ptgs = TrackedPtgs(LocationComparer())
                sheetMap[sheetname] = ptgs
            }
            if (!ptgs!!.contains(ptgRef))
            // **no duplicates allowed** (matches on location+parent rec)
                ptgs!!.add(ptgRef)
        } catch (e: Exception) {
        }

        return ptgRef


    }

    /**
     * Clears out the cached location of ptgrefs in the target
     * sheet.  This is required for copied worksheets with ptgrefs
     * contained internally
     *
     * @param targetSheet
     */
    fun clearPtgLocationCaches(targetSheet: String?) {
        var targetSheet = targetSheet
        try {
            targetSheet = GenericPtg.qualifySheetname(targetSheet)
            val ptgs = (sheetMap.get(targetSheet) as TrackedPtgs).values.iterator()
            //            Iterator ptgs= ((TrackedPtgs) sheetMap.get(targetSheet)).iterator();
            while (ptgs.hasNext()) {
                try {
                    val p = ptgs.next() as PtgRef
                    p.clearLocationCache()
                } catch (ex: Exception) {
                }

            }
        } catch (e: Exception) {

        }

    }

    /**
     * removes a cellrange Ptg (Area, Area3d etc.) to be tracked
     *
     * @param cr
     */
    fun removeCellRange(cr: Ptg) {
        if (cr !is PtgRef) {
            return
        }
        try {
            var sheetname: String? = ""
            try {
                sheetname = (cr as PtgRef).sheetName
                sheetname = GenericPtg.qualifySheetname(sheetname)
            } catch (ex: Exception) {
                sheetname = "WorkBookRanges"
            }

            val ptgs = sheetMap.get(sheetname) as TrackedPtgs
            if (ptgs != null) {
                ptgs!!.remove(cr)
            }
        } catch (e: Exception) {
            // this is common and not a problem normally - then we won't report a warning
            //Logger.logWarn("ReferenceTracker.removeCellRange failed for: " + cr.toString() +":"+ e);
        }

    }

    /**
     * updates the tracked ptg by using a new parent record
     *
     * @param pr     original ptg contained in tracker
     * @param parent new parent record of ptg
     */
    fun updateInRefTracker(pr: PtgRef, parent: XLSRecord) {
        if (pr is PtgRefErr || pr is PtgRefErr3d)
            return
        try {
            var sheetname: String? = ""
            try {
                sheetname = pr.sheetName
                sheetname = GenericPtg.qualifySheetname(sheetname)
            } catch (ex: Exception) {
                sheetname = "WorkBookRanges"
            }

            val ptgs = sheetMap.get(sheetname) as TrackedPtgs
            if (ptgs != null) {
                ptgs!!.update(pr, parent)
            }
        } catch (e: Exception) {
            // this is common and not a problem normally - then we won't report a warning
            //Logger.logWarn("ReferenceTracker.removeCellRange failed for: " + cr.toString() +":"+ e);
        }

    }

    /**
     * clear out object references in prep for closing workbook
     */
    fun close() {
        this.sheetMap.clear()
        this.nameRefs.clear()
        this.criteriaDBs.clear()
        this.CollectionDBs.clear()
        this.vlookups.clear()
        this.crs.clear()
        this.lookupColsCache.clear()
        sheetMap = HashMap()
        nameRefs = HashMap()
        // Database calc caches
        criteriaDBs = HashMap()
        CollectionDBs = HashMap()
        vlookups = HashMap()
        crs = Vector()
        lookupColsCache = HashMap()
    }

    companion object {

        /**
         * Add a cell to the Collection of afected cell handles
         *
         * @param celly
         */
        private fun addRec(celly: BiffRec, affectedCellHandles: MutableMap<*, *>) {
            val address = celly.sheet.sheetName + "!" + celly.cellAddress
            affectedCellHandles[address] = celly
            // add the new val to the reftracker and clear any cached
            // formula vals pointing to it, but do not recalc them
            // yes get rid of them all...
            try {
                (celly as Formula).clearCachedValue()
            } catch (e: ClassCastException) {
            }

        }


        /**
         * updateReferences
         * Shifts, Expands or Contracts ALL affected ranges upon a row or col insert or delete.
         * This will eventually eliminate the need for subequent
         * shifting done in moveFormulaCellReferences]
         *
         *
         * this will move the entire range down if range first row > startrow
         * it will expand the range if range first row <= startrow and range second row <= startrow
         *
         * @param start     0-based start row
         * @param shift     shift amount can be + or -
         * @param thissheet
         * @param shiftRow  true if shifting rows (false for columns)
         */

        fun updateReferences(start: Int, shiftamount: Int, thissheet: Boundsheet, shiftRow: Boolean) {
            var start = start
            // shift is 0-based, so that references to row shiftrow-1 and up are shifted
            // claritas is different, 1-based + shifts shiftrow+1 hence shiftInclusive setting
            // NOTE: shared formula references are the only ones that are NOT shifted via
            // updateReferences since PtgRefN and PtgAreaN's are NOT included in the
            // referenceTracker collection
            val shiftInclusive = thissheet.isShiftInclusive    // claritas-specific setting which directs us to expand ranges rather than shift when start of range==start
            val isExcel2008 = thissheet.workBook!!.isExcel2007    // limits are different between BIFF8 and Excel 2007
            if (shiftInclusive) start++    // make 1-based

            val updated = HashSet() // tracks which Ptgs have been already updated

            val sheetname = GenericPtg.qualifySheetname(thissheet.sheetName)
            val trackedptgs = thissheet.workBook!!.refTracker!!.sheetMap[sheetname] as TrackedPtgs
            if (trackedptgs == null || trackedptgs!!.size == 0) return
            var ptgs: Array<Any>? = null

            ptgs = trackedptgs!!.toArray()

            var i = if (shiftamount > 0) ptgs!!.size - 1 else 0
            val end = if (shiftamount > 0) 0 else ptgs!!.size
            val inc = if (shiftamount > 0) -1 else +1
            var done = false
            while (!done) {
                val p = ptgs!![i] as Ptg
                // skip these
                if (p is PtgRefErr || p is PtgRefErr3d)
                // these shouldn't be in the reference tracker ...
                    continue

                val pr = p as PtgRef
                if (!updated.contains(pr)) {
                    val sht: String?
                    try {
                        sht = pr.sheetName
                    } catch (e: Exception) {    // shouldn't happen
                        Logger.logErr("ReferenceTracker.updateReferences:  Error in Formula Reference Location: $e")
                        continue
                    }

                    sht = GenericPtg.qualifySheetname(sht)
                    if (sheetname == sht) {
                        if (shiftPtg(pr, sht, start, shiftamount, isExcel2008, shiftRow)) {
                            updated.add(pr)        // record which has already been updated to avoid incorrect expansion or movement
                        }
                    }
                }
                i += inc
                if (shiftamount > 0)
                    done = i < 0
                else
                    done = i == ptgs!!.size
            }
            // also update merged ranges which fall within range
            if (thissheet.hasMergedCells()) {
                // update mergedcells first, as they may grow in record size and are not handled by continues.
                val itx = thissheet.mergedCellsRecs.iterator()
                while (itx.hasNext()) {
                    val mrg = itx.next() as Mergedcells
                    val rngs = mrg.mergedRanges
                    for (j in rngs!!.indices) {
                        try {
                            val rc = rngs!![j].rangeCoords
                            rc[0]--
                            rc[2]--    // 1-based ...?
                            val isRange = rc.size > 2
                            var bUpdated = false
                            if (shiftRow) {
                                if (rc[0] >= start) {    // shift
                                    rc[0] += shiftamount
                                    if (isRange)
                                        rc[2] += shiftamount
                                    bUpdated = true
                                } else if (isRange && rc[2] >= start) { // expand
                                    rc[2] += shiftamount
                                    bUpdated = true
                                }
                            }
                            if (bUpdated) {
                                val newrange = thissheet + "!" + ExcelTools.formatLocation(rc)
                                rngs!![j].range = newrange
                            }
                        } catch (e: CellNotFoundException) {
                        }

                    }
                }
            }
        }


        /**
         * given a PtgRef, shifts correctly given start (row or col), shiftamount (+ or - 1) and truth of "shiftRow"
         *
         * @param ptgref      = the ptgref to move
         * @param sht         = the sheet in which the ptgref resides
         * @param start       = Start row for shifting - This is a 0 based value
         * @param shiftamount = amount to shift the ptgref
         * @param isExcel2007 true if use Excel-2007 maximums
         * @param shiftRow    = ?
         * @return true if updated PtgRef location
         */
        fun shiftPtg(ptgref: PtgRef, sht: String, start: Int, shiftamount: Int, isExcel2007: Boolean, shiftRow: Boolean): Boolean {
            val rc: IntArray?
            val iParent = ptgref.parentRec!!.opcode.toInt()
            val isNamedRange = iParent == XLSConstants.NAME.toInt()
            val isAi = iParent == XLSConstants.AI.toInt()
            val isShared = iParent == XLSConstants.SHRFMLA.toInt()
            try {
                rc = ptgref.intLocation
            } catch (e: Exception) {    // shouldn't happen!
                if (ptgref !is PtgAreaErr3d)
                // if it's not already an error Ptg report error
                    Logger.logErr("ReferenceTracker.shiftPtg:  Error in Formula Reference Location: $e")
                return false
            }

            val isRange = rc!!.size > 2
            var bUpdated = false
            if (shiftRow) {
                if (!isNamedRange && !isAi
                        && !ptgref.isRowRel)
                    return false    // if absolute don't shift (except for names and ai/charting refs, which should expand or shift in all cases)
                if (rc!![0] + 1 == start && isRange) { // expand don't shift
                    rc[2] += shiftamount
                    if (isAi && rc!![1] != rc!![3])
                    // Series in ROWS get shifted, not expanded ...
                        rc[0] += shiftamount
                    bUpdated = true
                } else if (rc!![0] + 1 >= start) {    // shift
                    rc[0] += shiftamount
                    if (isRange)
                        rc[2] += shiftamount
                    bUpdated = true
                } else if (isRange && rc!![2] + 1 >= start) { // expand
                    rc[2] += shiftamount
                    if (isAi && rc!![1] != rc!![3])
                    // Series in ROWS get shifted, not expanded ...
                        rc[0] += shiftamount
                    bUpdated = true
                }
                // SHIFTING EXCEPTION: if the parent formula cell is located ON the shifting row, do not shift
                if (bUpdated && iParent == XLSConstants.FORMULA.toInt() && ptgref.parentRec!!.rowNumber == start - 1) {
                    bUpdated = false
                }
            } else { // deal with columns in same way as above
                if (!isNamedRange && !isAi
                        && !ptgref.isColRel)
                    return false    // if absolute don't shift (except for names and ai/charting refs, which should expand or shift in all cases)
                if (rc!![1] + 1 >= start) {
                    rc[1] += shiftamount
                    if (isRange)
                        rc[3] += shiftamount
                    bUpdated = true
                } else if (isRange && rc!![3] + 1 >= start) {
                    rc[3] += shiftamount
                    bUpdated = true
                }
            }
            if (bUpdated) {
                // deal with limits
                if (isExcel2007) {
                    if (rc!![0] >= XLSConstants.MAXROWS)
                        rc[0] = XLSConstants.MAXROWS - 1
                    if (isRange && rc!![2] >= XLSConstants.MAXROWS)
                        rc[2] = XLSConstants.MAXROWS - 1
                } else {
                    if (rc!![0] >= XLSConstants.MAXROWS_BIFF8)
                        rc[0] = XLSConstants.MAXROWS_BIFF8 - 1
                    if (isRange && rc!![2] >= XLSConstants.MAXROWS_BIFF8)
                        rc[2] = XLSConstants.MAXROWS_BIFF8 - 1
                }
                var newaddr = ExcelTools.formatLocation(rc!!, ptgref.isRowRel, ptgref.isColRel)
                if (isRange && newaddr.indexOf(":") == -1)
                // handle special case of SHOULD be a range but 1st and last match
                    newaddr = "$newaddr:$newaddr"
                newaddr = "$sht!$newaddr"
                // NOW UPDATE THE PTG LOCATION
                try {
                    if (!isAi && !isShared) {
                        ptgref.location = newaddr    // should update ref tracker (remove and add new) appropriately
                        if (isNamedRange) {
                            // find formula references to the named range -- contained in reftracker.nameRefs
                            val rt = ptgref.parentRec!!.workBook!!.refTracker
                            val theName = (ptgref.parentRec as Name).nameA
                            if (rt!!.nameRefs.containsKey(theName)) {
                                // add all formulas that refer
                                val list = rt!!.nameRefs[theName] as ArrayList<*> // gets the ptgname
                                for (i in list.indices) {
                                    val ptgParent = (list[i] as Ptg).parentRec
                                    if (ptgParent.opcode == XLSConstants.NAME)
                                        continue // a Named Range referencing another named range ... will be caught later
                                    rt!!.clearAffectedFormulaCells(ptgParent) // recurse parent formula and get cells it affects
                                }
                            }
                        } else if (iParent == XLSConstants.CONDFMT.toInt())
                            (ptgref.parentRec as Condfmt).setDirty()    // flag to rebuild record
                    } else if (isShared) {
                        (ptgref.parentRec as Shrfmla).updateLocation(shiftamount, ptgref)
                    } else {    // Ai (chart) reference
                        (ptgref.parentRec as Ai).changeAiLocation(ptgref, newaddr)
                    }
                } catch (e: Exception) {
                    Logger.logInfo("ReferenceTracker.shiftPtg:  Shifting Formula Reference failed: $e")
                }

            }

            return bUpdated
        }

        /**
         * insert chart series upon an insert row
         * called by WSH.shiftRow
         * if series are row-based
         */
        fun insertChartSeries(c: Chart, sht: String, rownum: Int) {
            val rc: IntArray?
            var pr: PtgRef? = null
            val cursheet: String?
            var inserted = false
            val seriesmap = c.seriesPtgs
            val ii = seriesmap.keys.iterator()
            while (ii.hasNext() && !inserted) {
                val s = ii.next() as io.starter.formats.XLS.charts.Series
                val ptgs = seriesmap[s] as Array<Ptg>
                for (i in ptgs.indices) {
                    try {
                        pr = ptgs[i] as PtgRef
                        cursheet = pr!!.sheetName
                        rc = pr!!.intLocation
                    } catch (e: Exception) {    // shouldn't happen unless it's a PtgErr-type
                        continue
                    }

                    if (sht.equals(cursheet!!, ignoreCase = true)) { // if series are in rows, if existing series fall within inserted row, add new series
                        if (rc!![0] == rownum) { // already shifted ai range matches inserted row; take shifted range, shift backwards to get desired range to insert
                            val isRange = rc!!.size > 2
                            rc!![0]--
                            if (isRange)
                                rc!![2]--
                            // Adjust Series/Values Rane
                            var newseries = ExcelTools.formatLocation(rc!!, pr!!.isRowRel, pr!!.isColRel)
                            if (isRange && newseries.indexOf(":") == -1)
                            // handle special case of SHOULD be a range but 1st and last match
                                newseries = "$newseries:$newseries"
                            newseries = "$sht!$newseries"
                            // Adjust Legend Range
                            val legend = s.legendAi
                            rc = ExcelTools.getRowColFromString(legend!!.definition)
                            if (rc!![0] == rownum)
                                rc!![0]--
                            val legendRange = ExcelTools.formatLocation(rc!!)
                            // Adjust Bubble Range if present
                            val bubble = s.bubbleValueAi
                            var bubbleRange = ""
                            if (bubble != null && bubble!!.definition != "") {
                                rc = ExcelTools.getRowColFromString(bubble!!.definition)
                                if (rc!![0] == rownum)
                                    rc!![0]--
                                if (rc!!.size > 2)
                                    rc!![2]--
                                bubbleRange = ExcelTools.formatLocation(rc!!)
                            }
                            // Get Category range but don't alter
                            val categoryRange = s.categoryValueAi!!.definition        // category shouldn't shift
                            c.addSeries(newseries, categoryRange, bubbleRange, legendRange, "", 0)    // 0= default chart
                            c.setDimensionsRecord()
                            inserted = true
                        }
                    }
                }
            }
        }

        /* FORMULA REFS */

        /**
         * Update the address in a Ptg using the policy defined for the Ptg in an RDF if any.
         * called by changeFormulaLocation, addCellToRange
         *
         * @param thisptg
         * @param newaddr
         */
        fun updateAddressPerPolicy(thisptg: Ptg, newaddr: String) {
            val pl = thisptg.locationPolicy
            // Logger.logInfo("ReferenceTracker.updateAddressPerPolicy called.");
            when (pl) {
                Ptg.PTG_LOCATION_POLICY_UNLOCKED ->
                    // if this ptg has not been locked, update it
                    if (newaddr.indexOf("#REF!") > -1) {
                        val formula = thisptg.parentRec as Formula
                        formula.replacePtg(thisptg, PtgErr(PtgErr.ERROR_REF))
                    } else {
                        thisptg.location = newaddr
                    }

                Ptg.PTG_LOCATION_POLICY_TRACK ->
                    // this ptg tracks the cell that belongs to it...
                    thisptg.updateAddressFromTrackerCell()

                Ptg.PTG_LOCATION_POLICY_LOCKED -> {
                }
            }// do nothing
            //		}
        }

        /**
         * IF the newly inserted cell is a Formula, we need to update its cell range(s)
         * called by shiftRow
         *
         * @param copycell
         * @param oldaddr
         * @param newaddr
         * @param rownum
         * @param shiftRow true if shifting rows and not columns
         * @throws Exception
         * @see WorkSheetHandle.shiftRow
         */
        @Throws(Exception::class)
        fun adjustFormulaRefs(
                newcell: CellHandle,
                newrownum: Int,
                offset: Int,
                shiftRow: Boolean) {
            try {
                val sheet = newcell.workSheetName
                val isExcel2007 = newcell.workBook!!.workBook.isExcel2007
                val locptgs = newcell.formulaHandle.formulaRec!!.cellRangePtgs
                for (t in locptgs.indices) {
                    if (locptgs[t] is PtgRef) {
                        val pr = locptgs[t] as PtgRef
                        shiftPtg(pr, sheet, newrownum, offset, isExcel2007, shiftRow)
                    }
                }
            } catch (e: FormulaNotFoundException) {
                //
            }

        }
    }
}

/**
 * TrackedPtgs is a TreeMap override specific for PtgRefs to compare them more completely
 * (via location + parent-rec ...)
 *
 *
 * Each PtgRef is identified via a hashcode created from it's row/column (or row/column pairs for a PtgArea)
 * In this way a given PtgRef can be determined whether it is contained already, can be removed, updated and
 * it's parent records gathered in an efficient way.
 *
 *
 * TrackedPtgs must be instantiated with the custom Comparitor LocationComparer as:
 * new TrackedPtgs(new LocationComparer())
 * LocationComparer will return the correct compare for a PtgRef object, based upon it's location and parent record, in other words
 * PtgRef-A == PtgRef-B when both the location and the parent records are equal.
 */
internal class TrackedPtgs
/**
 * set the custom Comparitor for tracked Ptgs
 * Tracked Ptgs are referened by a unique key that is based upon it's location and it's parent record
 *
 * @param c
 */
(c: Comparator<*>) : TreeMap<*, *>(c) {

    /**
     * single access point for unique key creation from a ptg location and the ptg's parent
     * would be so much cleaner without the double precision issues ... sigh
     *
     * @param o
     * @return
     */
    @Throws(IllegalArgumentException::class)
    private fun getKey(o: Any?): Any {
        val loc = (o as PtgRef).hashcode
        if (loc == -1)
        // may happen on a referr (should have been caught earlier) or if not initialized yet
            throw IllegalArgumentException()
        val ploc = (o as PtgRef).parentRec!!.hashCode().toLong()
        return longArrayOf(loc, ploc)
    }

    /**
     * access point for unique key creation from separate identities
     *
     * @param loc  -- location key for a ptgref
     * @param ploc -- location key for a parent rec
     * @return Object key
     */
    private fun getKey(loc: Long, ploc: Long): Any {
        return longArrayOf(loc, ploc)
    }

    /**
     * override of add to record ptg location hash + parent has for later lookups
     */
    fun add(o: Any): Boolean {
        try {
            super.put(this.getKey(o), o)
        } catch (e: IllegalArgumentException) {    // SHOULD NOT HAPPEN -- happens upon RefErrs but they shouldnt be added ...
            // 	TESTING: report error
            //	System.err.println("Illegal PtgRef Location: " + o.toString());
        }

        return true
    }

    /**
     * override of the contains method to look up ptg location + parent record via hashcode
     * to see if it is already contained within the store
     */
    operator fun contains(o: Any): Boolean {
        return super.containsKey(getKey(o))
    }


    /**
     * returns EVERY parent that references the cell "cell"
     *
     * @param cell
     * @return iterator of biffrec parents of cell
     */
    // attempt to avoid concurrentmod exception by generating arrays, but doesn't work
    //    public Object[] getParents(BiffRec cell) {
    fun getParents(cell: BiffRec): Iterator<*> {
        val parents = ArrayList()
        val rc = intArrayOf(cell.rowNumber, cell.colNumber.toInt())
        val loc = PtgRef.getHashCode(rc[0], rc[1])    // get location in hashcode notation
        // first see if have tracked ptgs at the test location -- match all regardless of parent rec ...
        val key: Any
        var m: Map<*, *>? = Collections.synchronizedMap<Any, Any>(this.subMap(getKey(loc, 0), getKey(loc + 1, 0)))        // +1 for max parent
        if (m != null && m!!.size > 0) {
            val ii = m!!.keys.iterator()
            while (ii.hasNext()) {
                key = ii.next()
                val testkey = (key as LongArray)[0]
                /*			Object[] keys= m.keySet().toArray();
			for (int i= 0; i < keys.length; i++) {
				long testkey= ((long[])keys[i])[0];*/
                if (testkey == loc) {    // longs to remove parent hashcode portion of double
                    parents.add((this.get(key) as PtgRef).parentRec)
                    //					parents.add(((PtgRef)this.get(keys[i])).getParentRec());
                    //System.out.print(": Found ptg" + this.get((Integer)locs.get(key)));
                } else
                    break    // shouldn't hit here
            }
        }
        // now see if test cell falls into any areas

        m = Collections.synchronizedMap<Any, Any>(this.tailMap(getKey(SECONDPTGFACTOR, 0))) // NOW GET ALL PTGAREAS ...
        if (m != null) {
            synchronized(m!!.keys) {
                val ii = m!!.keys.iterator()
                while (ii.hasNext()) {
                    key = ii.next()
                    val testkey = (key as LongArray)[0]
                    /*			Object[] keys= m.keySet().toArray();
			for (int i= 0; i < keys.length; i++) {
					long testkey= ((long[])keys[i])[0];*/
                    val firstkey = (testkey / SECONDPTGFACTOR).toDouble()
                    val secondkey = (testkey % SECONDPTGFACTOR).toDouble()
                    if (firstkey.toLong() <= loc && secondkey.toLong() >= loc) {
                        val col0 = firstkey.toInt() % XLSRecord.MAXCOLS
                        val col1 = secondkey.toInt() % XLSRecord.MAXCOLS
                        val rw0 = (firstkey / XLSRecord.MAXCOLS).toInt() - 1
                        val rw1 = (secondkey / XLSRecord.MAXCOLS).toInt() - 1
                        if (this.isaffected(rc, intArrayOf(rw0, col0, rw1, col1))) {
                            parents.add((this.get(key) as PtgRef).parentRec)
                            //							parents.add(((PtgRef)this.get(keys[i])).getParentRec());
                        }
                    } else if (firstkey > loc)
                    // we're done
                        break
                }
            }
        }
        //io.starter.toolkit.Logger.log("");
        //		return parents.toArray();	//iterator();
        return parents.iterator()
    }

    /**
     * returns true if cell coordinates are contained within the area coordinates
     *
     * @param cellrc
     * @param arearc
     * @return
     */
    private fun isaffected(cellrc: IntArray, arearc: IntArray): Boolean {
        if (cellrc[0] < arearc[0]) return false // row above the first ref row?
        if (cellrc[0] > arearc[2]) return false // row after the last ref row?

        return if (cellrc[1] < arearc[1]) false else cellrc[1] <= arearc[3] // col before the first ref col?
// col after the last ref col?
    }

    /**
     * remove this PtgRef object via it's key
     */
    @Synchronized
    override fun remove(o: Any?): Any {
        return super.remove(getKey(o))
    }


    /**
     * update the location key for this PtgRef based upon a new parent record
     *
     * @param o      ptgref object
     * @param parent parent of ptgref
     */
    fun update(o: Any, parent: XLSRecord) {
        try {
            this.remove(o)
            val newloc = parent.hashCode().toLong()
            this[getKey((o as PtgRef).hashcode, newloc)] = o
        } catch (e: IllegalArgumentException) {
            // TESTING: report error
            //System.err.println("Illegal PtgRef Location: " + o.toString());
        }

    }


    fun toArray(): Array<Any> {
        return this.values.toTypedArray()
    }

    companion object {
        private val serialVersionUID = 1L
        val SECONDPTGFACTOR = XLSRecord.MAXCOLS.toLong() + XLSRecord.MAXROWS.toLong() * XLSRecord.MAXCOLS
    }


    /**
     * avoid double comparisons by converting double value to long
     * @param value
     * @return
     * NOTE: doesn't work in all cases -- shelve for now
     *
     * private long hashCode(long location, double parentlocation) {
     * long bits = Double.doubleToLongBits(location+parentlocation);
     * return (long)(bits ^ (bits >>> 32));
     * }
     */
}

/**
 * custom comparitor which compares keys for TrackerPtgs
 * consisting of a long ptg location hash, long parent record hashcode
 */
internal class LocationComparer : Comparator<*> {
    override fun compare(o1: Any, o2: Any): Int {
        val key1 = o1 as LongArray
        val key2 = o2 as LongArray
        if (key1[0] < key2[0]) return -1
        if (key1[0] > key2[0]) return 1
        if (key1[0] == key2[0]) {
            if (key1[1] == key2[1]) return 0    // equals
            if (key1[1] < key2[1]) return -1
            if (key1[1] > key2[1]) return 1
        }
        return -1
    }
}




