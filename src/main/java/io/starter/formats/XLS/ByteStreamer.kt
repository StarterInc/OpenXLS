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

import io.starter.formats.LEO.BIGBLOCK
import io.starter.formats.LEO.Block
import io.starter.formats.LEO.LEOFile
import io.starter.formats.LEO.Storage
import io.starter.toolkit.ByteTools
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger

import java.io.*
import java.util.*


/**
 * ByteStreamer Handles the low-level byte array streaming
 * of the WorkBook.  Basically a collection of XLS and methods for getting
 * at their data.
 *
 * @see WorkBook
 *
 * @see Boundsheet
 */
class ByteStreamer(bk: WorkBook) : Serializable, XLSConstants {
    internal var records: AbstractList<*> = FastAddVector()
    // public ByteStreamer(WorkBook b){this.book = b;}

    var bytes: ByteArray? = null
    val name = ""
    internal var DEBUG = !true
    protected var workbook: WorkBook? = null
    private val out: OutputStream? = null
    internal var dlen = -1

    /**
     * Returns an array of all the BiffRecs in the streamer.  For debug purposes,
     * as there is no way to directy view the FastAddVector
     *
     * @return
     */
    val biffRecords: Array<Any>
        get() = records.toTypedArray()

    val recVecSize: Int
        get() = records.size

    val substreamTypeName: String
        get() = "ByteStreamer"

    val substreamType: Short
        get() = XLSConstants.WK_WORKSHEET

    internal var ridx = 0


    fun getRecordIndex(opcode: Int): Int {

        for (i in records.indices) {
            val rec = records[i] as XLSRecord
            if (rec.opcode.toInt() == opcode) return i
        }
        return -1

    }

    fun initTestRecVect() {
        // Useful for looking at the output vector & debugging ... THANKS, NICK! -jm
        val testVect: Vector<*>
        if (true) {
            Logger.logInfo("TestVector on in bytestreamer.stream")
            testVect = Vector()
            testVect.addAll(records)
        }
        Logger.logInfo("TESTING Recvec done.")
    }

    init {
        this.workbook = bk
    }

    // 20060601 KSC: setBiffRecords
    fun setBiffRecords(recs: Collection<*>) {
        if (Arrays.equals(records.toTypedArray(), recs.toTypedArray())) {
            // already set!
            return
        }
        records.clear()
        records.addAll(recs)
    }

    /**
     * get a record from the underlying array
     */
    fun getRecordAt(t: Int, rec: BiffRec): BiffRec {
        if (rec is Boundsheet || rec is Dbcell)
            return this.getRecordAt(t)
        if (rec.sheet != null) {
            val bs = rec.sheet
            return bs.sheetRecs[t] as BiffRec
        }
        throw InvalidRecordException("ERROR: ByteStreamer.getRecord() could not retrieve record from")
    }

    /**
     * get a record from the underlying array
     */
    fun getRecordAt(t: Int): BiffRec {
        return records[t] as BiffRec
    }

    /**
     * remove a record from the underlying array
     */
    fun removeRecord(rec: BiffRec): Boolean {
        /*        if(rec instanceof Mulled) {
            if(((Mulled)rec).getMyMul()!=null) return true;
        }*/
        if (rec.sheet != null) {
            val bs = rec.sheet
            if (bs.sheetRecs.contains(rec)) return bs.sheetRecs.remove(rec)
        }
        return records.remove(rec)
    }

    /** get the recvec index of this record
     *
     * public List getRecordSubList(int start, int end){
     * return  records.subList(start,end);
     * } */

    /**
     * get the recvec index of this record
     */
    fun getRecordIndex(rec: BiffRec): Int {
        if (rec.sheet != null) {
            val bs = rec.sheet
            //	Logger.logInfo("ByteStreamer getting index of sheetRec: " + bs.toString());
            val ret = bs.sheetRecs.indexOf(rec) // bs.getRidx(); //
            return if (ret > -1) ret else records.indexOf(rec)
        }
        return records.indexOf(rec)
    }


    /**
     * get the real (non-boundsheet based) recvec index of this record
     */
    fun getRealRecordIndex(rec: BiffRec): Int {
        return records.indexOf(rec)
    }

    /**
     * add an BiffRec to this streamer.
     */
    fun addRecord(rec: BiffRec) {
        rec.streamer = this
        val sht = rec.sheet
        sht?.sheetRecs?.add(rec) ?: records.add(rec)
    }

    /**
     * Bypass the sheet vecs...
     *
     * @param rec
     * @param idx
     */
    fun addRecordToBookStreamerAt(rec: BiffRec, idx: Int) {
        try {
            records.add(idx, rec)
        } catch (e: ArrayIndexOutOfBoundsException) {
            records.add(records.size, rec)
        }

    }


    /**
     * add an BiffRec to this streamer at the specified index.
     */
    fun addRecordToSheetStreamerAt(rec: BiffRec, idx: Int, sht: Boundsheet) {
        rec.streamer = this
        val sr = sht.sheetRecs
        sr.add(idx, rec)
    }

    /**
     * add an BiffRec to this streamer at the specified index.
     */
    fun addRecordAt(rec: BiffRec, idx: Int) {
        rec.streamer = this
        val sht = rec.sheet
        if (sht != null) {
            val sr = sht.sheetRecs
            //  if(!sr.contains(rec))
            sr.add(idx, rec)
            // Logger.logInfo("ByteStreamer adding recAT: " + rec.toString() + " to Sheet: " + sht.toString());
        } else {
            // Logger.logInfo("ByteStreamer adding non-Sheet recAT: " + rec.toString());
            try {
                records.add(idx, rec)
                //     rec.setRecordIndexHint(idx);
            } catch (e: Exception) {
                records.add(records.size, rec)
                //   rec.setRecordIndexHint(records.size());
            }

        }
    }

    /**
     * stream the bytes to an outputstream
     */
    fun streamOut(_out: OutputStream): Int {
        writeOut(_out)
        return dlen - 4
    }

    @Throws(IOException::class)
    fun writeRecord(out: OutputStream, rec: BiffRec) {
        val op = ByteTools.shortToLEBytes(rec.opcode)
        val dt = rec.data
        val ln = ByteTools.shortToLEBytes(dt.size.toShort())
        out.write(op)
        out.write(ln)
        if (dt.size > 0)
            out.write(dt)
        if (LEOFile.DEBUG) LEOFile.actualOutput += op.size + ln.size + dt.size // debugging
        rec.postStream()
    }

    /**
     * Write all of the records to the output stream, including
     * creating lbplypos records, assembling continues, etc.
     */
    fun writeOut(out: OutputStream): StringBuffer? {
        // create a byte level lockdown file in same directory as output
        val lockdown = StringBuffer()
        var lockit = false
        if (System.getProperties()["io.starter.OpenXLS.autocreatelockdown"] != null)
            lockit = System.getProperties()["io.starter.OpenXLS.autocreatelockdown"] == "true"

        // update tracker cells, packs formats ...
        this.workbook!!.prestream()
        var dt: ByteArray
        var recpos = 0
        var recctr = 0
        var dlen = 0


        // get a private list of the records
        val rex = FastAddVector(records.size)
        rex.addAll(records)


        // first pass -- prepare SST
        var rec: BiffRec? = null
        var e: Iterator<*> = rex.iterator()
        while (e.hasNext()) {
            rec = e.next() as BiffRec
            ++recctr
            if (rec != null) {
                if (rec.byteReader != null)
                    rec.byteReader.applyRelativePosition = true
                // Logger.logInfo("ByteStreamer.stream() PREStreaming: "+ rec);
                if (rec.opcode == XLSConstants.BOUNDSHEET) {
                    //  				add sheet recs to output vector
                    val lst = (rec as Boundsheet).assembleSheetRecs()
                    rex.addAll(rex.size, lst)
                } else if (rec.opcode == XLSConstants.SST) {
                    // add extra bytes necessary for adding continue recx
                    rec.preStream()
                    recpos = (rec as Sst).numContinues * 4
                    if (recpos < 0)
                    // deal with empty SST
                        recpos = 0
                    dlen += recpos
                } else
                    rec.preStream() // perform expensive processes
            } else {
                Logger.logWarn("Body Rec missing while preStreaming(): " + rec!!.toString())
            }
        }
        e = rex.iterator()
        var lastindex: Index? = null
        var ctr = 0
        while (e.hasNext()) {
            rec = e.next() as BiffRec
            if (ctr == 0) { // handle the first BOF offset
                rec.offset = 0
                ctr++
            } else {
                rec.offset = recpos
            }

            if (rec.opcode == XLSConstants.INDEX) { // need to get all it's component dbcell offsets set before we can process!
                lastindex?.updateDbcellPointers()
                lastindex = rec as Index?
            }
            //            offset dlen by number of new continue headers
            if (rec.opcode == XLSConstants.CONTINUE && (rec as Continue).maskedMso != null) {
                rec.data = rec.maskedMso!!.getData()    // ensure any mso changes are propogated up
            }
            var rln = rec.length  //length of total rec data including continue
            var numcx = rln / XLSConstants.MAXRECLEN    // num continues?
            if (rln % XLSConstants.MAXRECLEN <= 4 && numcx > 0)
            // hits boundary; since rlen==datalen+4, numcx is 1 more than actual continues
                numcx--
            if (rec.opcode == XLSConstants.CONTINUE) {
                val thiscont = rec as Continue?
                if (thiscont!!.isBigRecContinue) {// could cause bugs... if related rec is trimmed
                    rln = 0 // do not count data byte size for Continues...
                }
            }
            if (rln > XLSConstants.MAXRECLEN + 4
                    && numcx > 0
                    && rec.opcode != XLSConstants.SST) {
                dlen += numcx * 4
                recpos += numcx * 4
            }
            dlen += rln
            recpos += rln
        }
        lastindex?.updateDbcellPointers()
        e = rex.iterator()

        /**
         * Get the updated Storages from LEO... output the RootStorage
         * and return the other storages and block index for output after
         * the workbook recs.
         *
         */
        val leo = this.workbook!!.factory!!.leoFile
        var storages: List<*>? = null
        storages = leo.writeBytes(out, dlen)
        val hdrBlock = BIGBLOCK()
        hdrBlock.init(leo.header!!.bytes, 0, 0)

        // ********************** WRITE OUT FILE IN CORRECT ORDER **************************************/
        /** Header = 1st sector/block  */
        hdrBlock.writeBytes(out)

        // now output the workbook biff records
        if (LEOFile.DEBUG) LEOFile.actualOutput = 0    // debugging
        while (e.hasNext()) {
            rec = e.next() as BiffRec

            try { // output the rec bytes
                // deal with CONTINUE record changes before streaming
                if (ContinueHandler.createContinues(rec, out, this)) {
                    // Logger.logInfo("Created continues for: " + rec.toString());
                } else {// Not a continued rec!
                    this.writeRecord(out, rec)

                    if (lockit) {
                        val op = ByteTools.shortToLEBytes(rec.opcode)
                        dt = rec.data
                        val ln = ByteTools.shortToLEBytes(dt.size.toShort())

                        // Logger.logInfo("=== WRITING RECORD DATA ===");
                        //lockdown.append("rec:" + rec.toString());
                        //lockdown.append("\r\n");
                        lockdown.append("opc:0x" + Integer.toHexString(rec.opcode.toInt()) + " [" + ByteTools.getByteString(op, false) + "]")
                        lockdown.append("\r\n")
                        lockdown.append("len:0x" + Integer.toHexString(rec.length) + " [" + ByteTools.getByteString(ln, false) + "]")
                        lockdown.append("\r\n")
                        //lockdown.append("off:0x" + Integer.toHexString(off+0x200));
                        //lockdown.append("\r\n");
                        lockdown.append(ByteTools.getByteDump(dt, 1))
                        lockdown.append("\r\n")
                    }
                }
            } catch (a: Exception) {
                throw WorkBookException(
                        "Streaming WorkBook Bytes failed for record: "
                                + rec.toString() + ": " + a + " Output Corrupted.",
                        WorkBookException.WRITING_ERROR, a)
            }

        }

        // pad to fit FAT size
        if (LEOFile.DEBUG) {
            if (LEOFile.actualOutput != dlen)
                Logger.logInfo("Expected:" + dlen + " Actual: " + LEOFile.actualOutput + " Diff: " + (LEOFile.actualOutput - dlen))
        }
        val leftover: Int
        val nBlocks = Math.max(leo.minBlocks, Math.ceil(dlen / (BIGBLOCK.SIZE * 1.0)).toInt() + 1)
        leftover = nBlocks * BIGBLOCK.SIZE - dlen    // padding


        val filler = ByteArray(leftover)
        for (t in filler.indices) filler[t] = 0

        // finally output the rest of the storages
        try {
            out.write(filler) // workbook data filler
            val its = storages!!.iterator()
            while (its.hasNext()) {
                val ob = its.next() as Storage
                if (ob != null)
                    if (ob.name != "Root Entry" || true) {  // was newLeoProcessing, so should always run?
                        if (ob.blockType != Block.SMALL) {// already written in miniFAT Container - see buildSAT
                            ob.writeBytes(out)
                            //Logger.logInfo("Streamed Storage:" + ob.getName() + ":"+ob.getActualFileSize());
                        }
                    }
            }
        } catch (a: Exception) {
            Logger.logErr("Streaming WorkBook Storage Bytes failed.", a)
            throw WorkBookException(
                    "ByteStreamer.stream(): Body Rec missing while preStreaming(): " + rec!!.toString()
                            + " " + a.toString() + " Output Corrupted.", WorkBookException.WRITING_ERROR)

        }

        return if (lockit) { // write the lockdown file
            lockdown
        } else null
    }

    /**
     * Debug utility to write out ALL BIFF records
     *
     * @param fName          output filename
     * @param bWriteSheeRecs truth of "include sheet records in output"
     */
    @JvmOverloads
    fun WriteAllRecs(fName: String, bWriteSheetRecs: Boolean = false) {
        try {
            val f = java.io.File(fName)
            var writer: BufferedWriter? = BufferedWriter(FileWriter(f))
            val recs = ArrayList(java.util.Arrays.asList<Any>(*this.biffRecords))
            var ctr = 0
            ctr = ByteStreamer.writeRecs(recs, writer, ctr, 0)
            val sheets = this.workbook!!.sheetVect
            for (i in sheets.indices) {
                if (bWriteSheetRecs) {
                    val lst = (sheets[i] as Boundsheet).assembleSheetRecs() as ArrayList<*>
                    ctr = ByteStreamer.writeRecs(lst, writer, ctr, 0)
                } else {
                    ctr = ByteStreamer.writeRecs((sheets[i] as Boundsheet).sheetRecs as ArrayList<*>, writer, ctr, 0)
                }
            }
            writer!!.flush()
            writer.close()
            writer = null
        } catch (e: Exception) {
        }

    }

    /**
     * clear out object references
     */
    fun close() {
        for (i in records.indices) {
            val r = records[i] as XLSRecord
            r.close()
        }
        records.clear()
        workbook = null
    }

    companion object {
        private const val serialVersionUID = -8188652784510579406L

        /**
         * Debug utility to write out BIFF records in friendly fashion
         *
         * @param recArr array of records
         * @param writer BufferedWriter
         * @param level  For ChartRecs that have sub-records, recurse Level
         */
        fun writeRecs(recArr: ArrayList<*>, writer: BufferedWriter, ctr: Int, level: Int): Int {
            val tabs = "\t\t\t\t\t\t\t\t\t\t"
            for (i in recArr.indices) {
                try {
                    val b = recArr[i] as BiffRec ?: break
                    writer.write(tabs.substring(0, level) + b.javaClass.toString().substring(b.javaClass.toString().lastIndexOf('.') + 1))
                    if (b is io.starter.formats.XLS.charts.SeriesText)
                        writer.write("\t[$b]")
                    else if (b is Continue) {
                        if (b.maskedMso != null) {
                            writer.write("\t[MASKED MSO ******************")
                            writer.write("\t[" + b.maskedMso!!.toString() + "]")
                            writer.write(b.maskedMso!!.debugOutput())
                            writer.write("\t[" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                        } else {
                            writer.write("\t[" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                        }
                    } else if (b is MSODrawing) {
                        writer.write("\t[$b]")
                        //					writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
                        writer.write(b.debugOutput())
                        writer.write("\t[" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                    } else if (b is Obj) {
                        writer.write(b.debugOutput())
                    } else if (b is MSODrawingGroup) {
                        writer.write("\t[$b]")
                        //					writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
                    } else if (b is Label) {
                        writer.write("\t[" + b.stringVal + "]")
                    } else if (b is Mulblank) {
                        writer.write("\t[" + b.cellAddress + "]")
                    } else if (b is Name) {
                        try {
                            writer.write("\t[" + b.name + "-" + b.location + "]")
                        } catch (ce: Exception) {
                            writer.write("\t[" + b.name + "-ERROR IN LOCATION]")
                        }

                        //writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
                    } else if (b is Sst)
                        writer.write("\t[$b]")
                    else if (b is Pls)
                        writer.write("\t[$b]")
                    else if (b is Supbook) {
                        writer.write("\t[$b]")
                        //					writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
                    } else if (b is Crn)
                        writer.write("\t[$b]")
                    else if (b is Formula || b is Rk || b is NumberRec || b is Blank || b is Labelsst)
                        writer.write(" " + (b as XLSRecord).cellAddressWithSheet + "\t[" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                    else
                    // all else, write bytes
                        writer.write("\t[" + ByteTools.getByteDump(ByteTools.shortToLEBytes(b.opcode), 0) + "][" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                    writer.newLine()
                    if (b is io.starter.formats.XLS.charts.ChartObject) {
                        writeRecs((b as io.starter.formats.XLS.charts.ChartObject).chartRecords, writer, ctr, level + 1)
                    }
                } catch (e: Exception) {
                }

            }
            return ctr
        }
    }
}