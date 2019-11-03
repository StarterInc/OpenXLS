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
import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.SxAddl.SxcCache
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList

/**
 * SXStreamID 0xD5
 * The SXStreamID record specifies the start of the stream in the PivotCache storage.
 *
 *
 * idStm (2 bytes): An unsigned integer that specifies a stream in the PivotCache storage. The stream specified is the one that has its name equal to the hexadecimal representation of this field. The four-digit hexadecimal string representation of this field, where each hexadecimal letter digit is a capital letter, MUST be equal to the name of a stream (1) in the PivotCache storage.
 */
class SxStreamID : XLSRecord(), XLSConstants {
    /**
     * serialVersionUID
     */
    /**
     * returns the streamId -- index linked to appropriate SxView pivot table view
     *
     * @return
     */
    var streamID: Short = -1
        private set
    private val subRecs = ArrayList()


    /**
     * returns the cache data sources
     * <br></br>NOT FULLY IMPLEMENTED - only valid for sheet data range data soures
     *
     * @return
     */
    /**
     * sets the cell range for this pivot cache
     *
     * @param cr
     */
    var cellRange: CellRange?
        get() {
            for (i in subRecs.indices) {
                val br = subRecs.get(i) as BiffRec
                if (br.opcode == XLSConstants.SXVS) {
                    if ((br as SxVS).sourceType != SxVS.TYPE_SHEET) {
                        Logger.logErr("SXSTREAMID.getCellRange:  Pivot Table Data Sources other than Sheet are not supported")
                        return null
                    }
                } else if (br.opcode == XLSConstants.DCONREF) {
                    return (br as DConRef).cellRange
                } else if (br.opcode == XLSConstants.DCONNAME) {
                    Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported")
                    return null
                } else if (br.opcode == XLSConstants.DCONBIN) {
                    Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported")
                    return null
                }
            }
            return null
        }
        set(cr) {
            for (i in subRecs.indices) {
                val br = subRecs.get(i) as BiffRec
                if (br.opcode == XLSConstants.SXVS) {
                    if ((br as SxVS).sourceType != SxVS.TYPE_SHEET) {
                        Logger.logErr("SXSTREAMID.setCellRange:  Pivot Table Data Sources other than Sheet are not supported")
                        return
                    }
                } else if (br.opcode == XLSConstants.DCONREF) {
                    (br as DConRef).cellRange = cr
                    return
                } else if (br.opcode == XLSConstants.DCONNAME) {
                    Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported")
                    return
                } else if (br.opcode == XLSConstants.DCONBIN) {
                    Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported")
                    return
                }
            }
        }

    /**
     * init method
     */
    override fun init() {
        super.init()
        streamID = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        if (DEBUGLEVEL > 3) Logger.logInfo("SXSTREAMID: streamid:$streamID")
    }

    /**
     * store like cache-related records under SxStreamID
     *
     * @param r
     */
    fun addSubrecord(r: BiffRec) {
        subRecs.add(r)
    }

    /**
     * sets the streamId -- index linked to approriate SxView pivot table view
     *
     * @param sid
     */
    fun setStreamID(sid: Int) {
        streamID = sid.toShort()
        val b = ByteTools.shortToLEBytes(streamID)
        this.getData()[0] = b[0]
        this.getData()[1] = b[1]
    }

    /**
     * sets the cell range for this pivot cache
     *
     * @param cr
     */
    fun setCellRange(cr: String) {
        for (i in subRecs.indices) {
            val br = subRecs.get(i) as BiffRec
            if (br.opcode == XLSConstants.SXVS) {
                if ((br as SxVS).sourceType != SxVS.TYPE_SHEET) {
                    Logger.logErr("SXSTREAMID.setCellRange:  Pivot Table Data Sources other than Sheet are not supported")
                    return
                }
            } else if (br.opcode == XLSConstants.DCONREF) {
                (br as DConRef).setCellRange(cr)
                return
            } else if (br.opcode == XLSConstants.DCONNAME) {
                Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported")
                return
            } else if (br.opcode == XLSConstants.DCONBIN) {
                Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported")
                return
            }
        }
    }

    /**
     * creates the basic, default records necessary to define a pivot cache
     *
     * @param bk
     * @param ref       string datasource range or named range reference
     * @param sheetName string datasource sheetname where ref is located
     * @return arraylist of records
     */
    fun addInitialRecords(bk: WorkBook, ref: String, sheetName: String): ArrayList<*> {
        val initialrecs = ArrayList()
        val sid = this.streamID.toInt()
        val sxvs = SxVS.prototype as SxVS?
        addInit(initialrecs, sxvs!!, bk)
        if (bk.getName(ref) != null) {
            // DConName or DConBin
            Logger.logErr("PivotCache:  Name Data Sources are Not Supported")
        } else {    // assume it's a regular reference
            // DConRef
            val dc = DConRef.prototype as DConRef?
            val rc = ExcelTools.getRangeRowCol(ref)
            dc!!.setRange(rc, sheetName)
            addInit(initialrecs, dc, bk)
        }
        // required SxAddl records: stores additional PivotTableView, PivotCache info of a variety of types
        var b = ByteTools.cLongToLEBytes(sid)
        b = ByteTools.append(byteArrayOf(0, 0), b) // add 2 reserved bytes
        var sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdId.sxd().toInt(), b)    //4 bytes sid, 2 bytes reserved
        addInit(initialrecs, sa, bk)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdVer10Info.sxd().toInt(), null)
        addInit(initialrecs, sa, bk)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdVerSxMacro.sxd().toInt(), null)
        addInit(initialrecs, sa, bk)
        sa = SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdEnd.sxd().toInt(), null)
        addInit(initialrecs, sa, bk)
        return initialrecs
    }

    /**
     * utility function to properly add a Pivot Table View subrec
     *
     * @param initialrecs
     * @param rec
     * @param addToInitRecords
     * @param sheet
     */
    private fun addInit(initialrecs: ArrayList<*>, rec: XLSRecord, bk: WorkBook) {
        rec.workBook = bk
        initialrecs.add(rec)
        this.addSubrecord(rec)
    }

    companion object {
        private val serialVersionUID = 2639291289806138985L

        /**
         * creates new, default SxStreamID
         *
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val ss = SxStreamID()
                ss.opcode = XLSConstants.SXSTREAMID
                ss.setData(byteArrayOf(0, 0))
                ss.init()
                return ss
            }
    }
}
