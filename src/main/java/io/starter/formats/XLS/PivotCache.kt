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

import io.starter.OpenXLS.CellHandle
import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.LEO.BlockByteReader
import io.starter.formats.LEO.Storage
import io.starter.formats.LEO.StorageNotFoundException
import io.starter.formats.LEO.StorageTable
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.HashMap

/**
 * represents the required storage _SX_CUR_DB for Pivot Tables
 */

/**
 *
 */
class PivotCache : XLSConstants {
    internal var book: WorkBook? = null
    internal var caches: HashMap<Int, Storage> = HashMap()
    internal var pivotCacheRecs: HashMap<Int, ArrayList<BiffRec>> = HashMap()
    //	HashMap<String, ArrayList> caches= new HashMap();

    @Throws(StorageNotFoundException::class)
    fun init(directories: StorageTable, wbh: WorkBookHandle?) {
        var child = directories.getChild("_SX_DB_CUR")
        if (wbh != null) {
            caches = HashMap()
            book = wbh.workBook
            if (wbh.debugLevel > 25) {        // KSC: TESTING: make > 25
                Logger.logInfo("PivotCache.init")
            }
        }

        while (child != null) {
            if (wbh != null)
                caches[Integer.valueOf(child.name)] = child
            val curRecs = ArrayList()
            val bytes = child.blockReader
            val len = bytes.length
            var i = 0
            while (i <= len - 4) {
                val headerbytes = bytes.getHeaderBytes(i)
                val opcode = ByteTools.readShort(headerbytes[0].toInt(), headerbytes[1].toInt())
                val reclen = ByteTools.readShort(headerbytes[2].toInt(), headerbytes[3].toInt()).toInt()
                val rec = XLSRecordFactory.getBiffRecord(opcode)
                if (wbh != null) rec.setDebugLevel(wbh.debugLevel) // KSC: added to propogate debug level

                // init the mighty rec
                rec.workBook = book
                rec.byteReader = bytes
                rec.length = reclen.toShort()
                rec.offset = i
                rec.init()
                /*
// KSC: TESTING
		try {
io.starter.toolkit.Logger.log(rec.getClass().getName().substring(rec.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(((PivotCacheRecord)rec).getRecord()));
		} catch (ClassCastException e) {
io.starter.toolkit.Logger.log(rec.getClass().getName().substring(rec.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(ByteTools.shortToLEBytes(rec.getOpcode())) + Arrays.toString(ByteTools.shortToLEBytes((short)rec.getData().length)) + Arrays.toString(rec.getData()));
		}
*/
                if (wbh != null)
                    curRecs.add(rec)

                i += reclen + 4
            }
            if (wbh != null)
                pivotCacheRecs[Integer.valueOf(child.name)] = curRecs
            child = directories.getNext(child.name)
        }
        // KSC: TESTING
        //		Logger.logInfo("PivotCache.end init");
    }

    /**
     *
     * Creates the pivot table cache == defines pivot table data
     * <br></br>A pivot table cache requires 2 directory storages
     * <br></br>_SX_DB_CUR = parent pivot cache
     * <br></br>0001, 0002 ... = child streams that define the pivot cache records
     * @param directories
     * @param wb
     * @param ref  Cell Range which identifies pivot table data range
     */
    @Throws(InvalidRecordException::class)
    fun createPivotCache(directories: StorageTable, wbh: WorkBookHandle, ref: String, sId: Int) {
        try {
            // KSC: TESTING
            if (wbh.debugLevel > 100)
                io.starter.toolkit.Logger.log(String.format("creatpivotCache: ref: %s sid %d", ref, sId))
            /**
             * the Pivot Cache Storage specifies zero or more streams, each of which specify a PivotCache
             * The name of each stream (1) MUST be unique within the storage, and the name MUST be a four digit hexadecimal number stored as text.
             * The number of FDB rules that occur MUST be equal to the value of cfdbTot in the SXDBrecord (section 2.4.275).
             */
            /*
             */
            // KSC: unsure if it's absolutely necessary to also have CompObj storage
            try {
                directories.getDirectoryByName("\u0001CompObj")
            } catch (e: StorageNotFoundException) {
                val compObj = directories.createStorage("\u0001CompObj", 2, directories.getDirectoryStreamID("\u0005DocumentSummaryInformation") + 1)
                compObj.setBytesWithOverage(byteArrayOf(1, 0, -2, -1, 3, 10, 0, 0, -1, -1, -1, -1, 32, 8, 2, 0, 0, 0, 0, 0, -64, 0, 0, 0, 0, 0, 0, 70, 38, 0, 0, 0, 77, 105, 99, 114, 111, 115, 111, 102, 116, 32, 79, 102, 102, 105, 99, 101, 32, 69, 120, 99, 101, 108, 32, 50, 48, 48, 51, 32, 87, 111, 114, 107, 115, 104, 101, 101, 116, 0, 6, 0, 0, 0, 66, 105, 102, 102, 56, 0, 14, 0, 0, 0, 69, 120, 99, 101, 108, 46, 83, 104, 101, 101, 116, 46, 56, 0, -12, 57, -78, 113, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
                val compObjid = directories.getDirectoryStreamID("\u0001CompObj")
                val wb = directories.getDirectoryByName("Workbook")
                wb!!.setPrevStorageID(compObjid)
            }

            /* create _SX_DB_CUR + child (actual pivot cache) directory and put in proper order in directory array */
            val sx_db_cur = directories.createStorage("_SX_DB_CUR", 1, directories.getDirectoryStreamID("\u0005SummaryInformation"))    // Pivot Cache Storage:  (id= 1) insert just before SummaryInfo -- ALWAYS ??????
            val sxdbcurid = directories.getDirectoryStreamID("_SX_DB_CUR")    //
            val pcache1 = directories.createStorage("0001", 2, sxdbcurid + 1)            // TODO: handle multiple Caches ... 0002 ...
            directories.getDirectoryByName("Root Entry")!!.setChildStorageID(sxdbcurid)
            sx_db_cur.setPrevStorageID(directories.getDirectoryStreamID("Workbook"))
            sx_db_cur.setChildStorageID(directories.getDirectoryStreamID("0001"))            // child= 0001 Cache Stream (id= 2)
            sx_db_cur.setNextStorageID(directories.getDirectoryStreamID("\u0005SummaryInformation"))
            val si = directories.getDirectoryByName("\u0005SummaryInformation")
            si!!.setPrevStorageID(-1)    // Necessary?????????
            si.setNextStorageID(directories.getDirectoryStreamID("\u0005DocumentSummaryInformation"))
            directories.getDirectoryByName("Root Entry")!!.setChildStorageID(sxdbcurid)    // ??? ALWAYS ????

            // create pivot cache records which are source of actual pivot cache data
            val newbytes = createPivotCacheRecords(ref, wbh, sId)
            pcache1.setBytesWithOverage(newbytes)
            this.init(directories, wbh)
        } catch (e: StorageNotFoundException) { // shouldn't!
        }

    }


    /**
     * adds a specific instance of a cache field
     * <br></br>A cache item is contained in a cache field. A cache field can have zero cache items if the cache field is not in use in the PivotTable view.
     * TODO: handle unique cache items ...
     *
     * // TODO: need cache field index ***
     * @param cacheItem
     */
    fun addCacheItem(cacheId: Int, cacheItem: Int) {
        var insertIndex = 0
        for (br in pivotCacheRecs[cacheId + 1]) {
            if (br.opcode == XLSConstants.SXFDB) {
                (br as SxFDB).nCacheItems = (br as SxFDB).nCacheItems + 1
            } else if (br.opcode == XLSConstants.EOF) {
                insertIndex = pivotCacheRecs[cacheId + 1].indexOf(br)
                /**/
            }
        }
        // add required SXDBB for non-summary-cache items
        if (cacheItem > -1) {
            /* SXDBB records only exist when put cache fields on a pivot table axis == cache item*/
            val sxdbb = SxDBB.prototype as SxDBB?
            sxdbb!!.setCacheItemIndexes(byteArrayOf(Integer.valueOf(cacheItem)!!.toByte()))    //ByteTools.shortToLEBytes((short)cacheItem));
            pivotCacheRecs[cacheId + 1].add(insertIndex, sxdbb)
        }
        updateCacheRecords(cacheId)
    }

    /**
     * take pivotCacheRecs and update the actual cache bytes
     */
    private fun updateCacheRecords(cacheId: Int) {
        var newbytes = ByteArray(0)
        for (br in pivotCacheRecs[cacheId + 1]) {
            try {
                newbytes = ByteTools.append((br as PivotCacheRecord).record, newbytes)
                //io.starter.toolkit.Logger.log(br.getClass().getName().substring(br.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(((PivotCacheRecord)br).getRecord()));
            } catch (e: ClassCastException) {
                newbytes = ByteTools.append(ByteTools.shortToLEBytes(br.opcode), newbytes)
                newbytes = ByteTools.append(ByteTools.shortToLEBytes(br.data.size.toShort()), newbytes)
                newbytes = ByteTools.append(br.data, newbytes)
                //io.starter.toolkit.Logger.log(br.getClass().getName().substring(br.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(ByteTools.shortToLEBytes(br.getOpcode())) + Arrays.toString(ByteTools.shortToLEBytes((short)br.getData().length)) + Arrays.toString(br.getData()));
            }

        }
        val pcache1 = caches[cacheId + 1]
        pcache1.setBytesWithOverage(newbytes)
        // KSC: TESTING
        /*		try {
		this.init(book.factory.myLEO.getDirectoryArray(), null);
		} catch (StorageNotFoundException e) {
		}*/
    }

    /**
     * parse range and create required cache records, returning bytes defining said records.
     * <br></br>For normal ranges, the PivotCache has one cache field for each column of the range,
     * using the values in the first row of the range for cache field names,
     * and all other rows are used as source data values, specified by cache records.
     *
     * @param ref    Range Reference String, including sheet
     * @param wbh   workbookhandle
     * @param sId    Stream or cachid Id -- links back to SxStream set of records
     */
    internal fun createPivotCacheRecords(ref: String, wbh: WorkBookHandle, sId: Int): ByteArray {
        var newbytes = ByteArray(0)
        try {
            val cr = CellRange(ref, wbh, false, true)
            val ch = cr.getCells()    // cells are in row-order
            val rows = cr.rowInts    // first row= header, ensuing rows are cacherecords
            val cols = cr.colInts
            val types = IntArray(cols.size)
            val cachefieldindexes = Array(cols.size) { ByteArray(rows.size - 1) }
            val sxdb = SxDB.prototype as SxDB?
            sxdb!!.nCacheRecords = rows.size - 1
            sxdb.nCacheFields = cols.size
            sxdb.setStreamID(sId)
            newbytes = ByteTools.append(sxdb.record, newbytes)
            //io.starter.toolkit.Logger.log("SXDB: " + Arrays.toString(sxdb.getRecord()));
            val sxdbex = SXDBEx.prototype as SXDBEx?    //TODO: nFormulas
            newbytes = ByteTools.append(sxdbex!!.record, newbytes)
            //io.starter.toolkit.Logger.log("SXDBEX: " + Arrays.toString(sxdbex.getRecord()));
            // TODO: cells after row header cell ***should be*** the same type -- true in ALL cases??????
            if (ch!!.size > cols.size) { // have multiple rows
                for (i in cols.indices) {
                    val c = ch[i + cols.size]
                    var type = -1
                    if (c.isDate)
                        type = 6
                    else
                        type = c.cellType
                    types[i] = type
                }
            }
            //TODO: ranges/grouping and formulas !!!!
            //TODO: boolean vals?
            for (z in rows.indices) {
                for (i in cols.indices) {
                    if (z == 0) { // # SxFDB records==# COLUMNS==# Cache Fields
                        val sxfdb = SxFDB.prototype as SxFDB?
                        sxfdb!!.setCacheItemsType(types[i])
                        sxfdb.setCacheField(ch[i].stringVal)    // row header values
                        sxfdb.nCacheItems = 0        // only set ACTUAL cache items when put cache field(s)on the pivot table (on row, page, column or data axis)
                        newbytes = ByteTools.append(sxfdb.record, newbytes)
                        //io.starter.toolkit.Logger.log("SXFDB: " + Arrays.toString(sxfdb.getRecord()));
                        val sxfdbtype = SXFDBType.prototype as SXFDBType?
                        newbytes = ByteTools.append(sxfdbtype!!.record, newbytes)
                        //io.starter.toolkit.Logger.log("SXDFBTYPE: " + Arrays.toString(sxfdbtype.getRecord()));
                        continue
                    }
                    cachefieldindexes[i][z - 1] = i.toByte()
                    // data cells== CACHE ITEMS
                    val c = ch[z * cols.size + i]
                    // TODO: handle SxNil, SxErr, SxDtr
                    // TODO: handle SxFmla, SXName, SxPair, SxFormula
                    when (types[i]) {
                        XLSConstants.TYPE_STRING -> {
                            val sxstring = SXString.prototype as SXString?
                            sxstring!!.cacheItem = c.stringVal
                            newbytes = ByteTools.append(sxstring.record, newbytes)
                        }
                        XLSConstants.TYPE_FP, XLSConstants.TYPE_INT, XLSConstants.TYPE_DOUBLE -> {
                            val sxnum = SXNum.prototype as SXNum?
                            sxnum!!.setNum(c.doubleVal)
                            newbytes = ByteTools.append(sxnum.record, newbytes)
                        }
                        XLSConstants.TYPE_BOOLEAN -> {
                            val sxbool = SXBool.prototype as SXBool?
                            sxbool!!.setBool(c.booleanVal)
                            newbytes = ByteTools.append(sxbool.record, newbytes)
                        }
                    }//io.starter.toolkit.Logger.log("SXSTRING: " + Arrays.toString(sxstring.getRecord()));
                    //io.starter.toolkit.Logger.log("SXNUM: " + Arrays.toString(sxnum.getRecord()));
                    //io.starter.toolkit.Logger.log("SXBOOL: " + Arrays.toString(sxbool.getRecord()));
                    //TYPE_FORMULA = 3,		SxFmla *(SxName *SXPair)
                    // SXDtr
                }
            }
        } catch (e: Exception) {
            throw InvalidRecordException("PivotCache.createPivotCache: invalid source range: $ref")
        }

        // EOF -- header:
        val b = ByteArray(4)
        System.arraycopy(ByteTools.shortToLEBytes(XLSConstants.EOF), 0, b, 0, 2)
        newbytes = ByteTools.append(b, newbytes)
        //io.starter.toolkit.Logger.log("EOF: " + Arrays.toString((b)));
        return newbytes
    }
}
