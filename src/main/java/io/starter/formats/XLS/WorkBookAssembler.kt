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

import io.starter.formats.XLS.charts.Chart

import java.util.ArrayList
import java.util.Random

/**
 * WorkBookAssembler handles the details of how a biff8 file
 * should be constructed, what records go in what order, etc
 */
class WorkBookAssembler : XLSConstants {
    companion object {


        /**
         * assembleSheetRecs assembles the array of records, then ouputs
         * the ordered list to the bytestreamer, which should be the only
         * thing calling this.
         *
         *
         * The SheetRecs should contain all records excepting valrecs, rows, index, and dbcells at this
         * point, where they should be created.
         *
         *
         * The dbcell creation and population also occurs here.  The first pointer is the offset between
         * the beginning of the second row and the first valrec.  After that it is just incrementing valrecs
         *
         *
         * TODO:  John,  please review use of collections & performance.  Check the Dbcell.initDbCell method as well.
         * TODO:  Add handling for creation of mulblank and mulrk records on the fly.  We no longer have them!
         */
        fun assembleSheetRecs(thissheet: Boundsheet): List<*> {
            var addVec: MutableList<*> = ArrayList()
            WorkBookAssembler.preProcessSheet(thissheet)
            addVec.addAll(thissheet.sheetRecs)
            if (!thissheet.isChartOnlySheet) {
                addVec = WorkBookAssembler.assembleRows(thissheet, addVec)
            }
            if (thissheet.workBook!!.charts.size > 0) {
                addVec = WorkBookAssembler.assembleChartRecs(thissheet, addVec)
            }
            return addVec
        }

        /**
         * Handles functionality that needs to occur before the boundsheet can be
         * reliably created
         */
        private fun preProcessSheet(thissheet: Boundsheet) {
            if (thissheet.hasMergedCells()) {
                // update mergedcells first, as they may grow in record size and are not handled by continues.
                val itx = thissheet.mergedCellsRecs.iterator()
                while (itx.hasNext()) {
                    val mrg = itx.next() as Mergedcells
                    mrg.update()
                }
            }
        }

        /**
         * Add the rows of data into the worksheet level stream
         *
         * @param thissheet
         * @param addVec
         * @return
         */
        private fun assembleRows(thissheet: Boundsheet, addVec: MutableList<*>): MutableList<*> {
            val dim = thissheet.dimensions
            val gen = Random()
            val randomNumber = gen.nextInt()
            var insertRowidx = 0
            if (dim != null) insertRowidx = dim.recordIndex + 1
            var insertValidx = insertRowidx
            var rowCount = 0 // use this to break every 32 rows for a new dbcell
            var dbOffset = 0 // the offset between the first row and the dbcell
            var maxRow = 0
            var maxCol = 0
            var dbOffsets: MutableList<*> = ArrayList() // dbcell offsets
            var valrecOffset = 0
            // KSC: clear out index dbcells
            if (thissheet.sheetIDX != null)
                thissheet.sheetIDX!!.resetDBCells()
            // NOTE: if below sorting is time-consuming, should input sort inserting and deleting rows and not here ...
            val outRows = thissheet.sortedRows.keys.iterator()
            while (outRows.hasNext()) {
                if (rowCount == 32) {
                    // Add a new dbcell for every 32 rows
                    val d = Dbcell.prototype as Dbcell?
                    dbOffsets.add(0, java.lang.Short.valueOf(((rowCount - 1) * 20).toShort()))
                    d!!.initDbCell(dbOffset, dbOffsets)
                    addVec.add(insertValidx++, d)
                    if (thissheet.sheetIDX != null)
                        thissheet.sheetIDX!!.addDBCell(d)
                    insertRowidx = insertValidx
                    dbOffsets = ArrayList()
                    rowCount = 0
                    dbOffset = 0
                }
                val r = thissheet.rowMap[outRows.next()] as Row
                rowCount++
                maxRow = Math.max(r.rowNumber, maxRow)
                dbOffset += r.length
                addVec.add(insertRowidx++, r)
                insertValidx++

                // insert the valrec and child recs from the row
                var skipMull: Mulblank? = null
                val outRecs = r.getValRecs(randomNumber)
                val it = outRecs.iterator()
                while (it.hasNext()) {
                    val or = it.next() as BiffRec
                    if (skipMull != null) {
                        if (or === skipMull)
                            continue
                        skipMull = null
                    }
                    addVec.add(insertValidx++, or)
                    valrecOffset += or.length
                    dbOffset += or.length
                    val orc = or.colNumber
                    if (!it.hasNext() && orc > maxCol)
                        maxCol = orc.toInt()
                    if (or.opcode == XLSConstants.MULBLANK)
                        skipMull = or as Mulblank
                }
                dbOffsets.add(java.lang.Short.valueOf(valrecOffset.toShort()))
                valrecOffset = 0
            }
            // add the final dbcell.  Chart only sheets will not have an index, so ignore if so.
            if (dbOffsets.size > 0) {
                val d = Dbcell.prototype as Dbcell?
                dbOffsets.add(0, java.lang.Short.valueOf(((rowCount - 1) * 20).toShort()))
                d!!.initDbCell(dbOffset, dbOffsets)
                d.setOffset(dbOffset)    // KSC: Added to set dbCell offset
                addVec.add(insertValidx++, d)
                if (thissheet.sheetIDX != null)
                    thissheet.sheetIDX!!.addDBCell(d)
            }
            thissheet.updateDimensions(maxRow, maxCol/* 20100225 KSC: incrementing does not match Excel results: take out +1*/)
            return addVec
        }

        /**
         * Get the index to the last db cell in the boundsheet
         * this is where objects can start being inserted
         *
         * @return
         */
        private fun getLastDBCellLocation(addVec: List<*>): Int {
            for (i in addVec.indices.reversed()) {
                val b = addVec[i] as BiffRec
                if (b.opcode == XLSConstants.DBCELL) return i
            }
            return 0
        }

        /**
         * Add the chart records to the array.  There are two types of charts to be added,
         * those that have a pointer into the worksheet level stream (ie they existed when the file was parsed)
         * and those that have been added post parse and/or are converted from ooxml.
         *
         *
         * We hold the location of the parsed ones to deal with odd inconsistencies that exist
         * with multiple drawing objects and txos
         *
         * @param thissheet
         * @param addVec
         * @return
         */
        private fun assembleChartRecs(thissheet: Boundsheet, addVec: MutableList<*>): MutableList<*> {
            // insert charts that have bound obj records in the stream
            var insertValidx = WorkBookAssembler.getLastDBCellLocation(addVec)
            val insertedCharts = ArrayList()
            var chartInsert = insertValidx
            while (chartInsert < addVec.size) {
                val x = addVec[chartInsert] as XLSRecord
                if (x.opcode == XLSConstants.OBJ) {
                    val o = x as Obj
                    if (o.chart != null) {
                        insertedCharts.add(o.chart!!.title)
                        val l = o.chart!!.assembleChartRecords()
                        addVec.addAll(chartInsert + 1, l)
                        chartInsert += l.size
                        insertValidx += l.size
                    }
                }
                chartInsert++
            }

            // insert charts that are new, either from transfers, insertions, etc
            val chts = thissheet.workBook!!.charts
            chartInsert = insertValidx
            for (i in chts.indices) {
                if (!insertedCharts.contains(chts[i].title) && chts[i].sheet == thissheet) {
                    // if it's a chart only sheet, insertValidx will be '0' here, put it at the end of the current recordset
                    if (insertValidx == 0) insertValidx = addVec.size
                    val l = chts[i].assembleChartRecords()
                    if (chts[i].obj != null) l.add(0, chts[i].obj)
                    if (chts[i].msodrawobj != null) l.add(0, chts[i].msodrawobj)
                    var spid = 0
                    var isHeader = false
                    if (l[0] is MSODrawing) {
                        isHeader = (l[0] as MSODrawing).isHeader
                        spid = (l[0] as MSODrawing).spid
                    }
                    while (chartInsert < addVec.size) {
                        val x = addVec[chartInsert] as XLSRecord
                        if (x.opcode == XLSConstants.MSODRAWING) {
                            val mm = x as MSODrawing
                            if (mm.spid > spid && spid != 0 && !mm.isHeader || isHeader) {
                                insertValidx = chartInsert
                                chartInsert = addVec.size
                            }
                        }
                        if (x.opcode == XLSConstants.WINDOW2 || x.opcode == XLSConstants.MSODRAWINGSELECTION
                                || x.opcode == XLSConstants.NOTE) {
                            insertValidx = chartInsert
                            chartInsert = addVec.size
                        }
                        chartInsert++
                    }
                    addVec.addAll(insertValidx, l)
                    insertValidx += l.size
                }
            }
            return addVec
        }
    }
}
