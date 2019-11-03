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

import io.starter.toolkit.Logger

import java.io.OutputStream
import java.io.Serializable


/**
 * this class takes care of the tasks related to
 * working with Continue records.
 *
 *
 * when a Continue record is created, it needs to be
 * associated with its related XLSRecord -- the record whose
 * data it contains.
 */
class ContinueHandler(private var book: WorkBook?) : Serializable, XLSConstants {
    private var continued: BiffRec? = null
    private var handleTxo = false
    private var handleObj = false
    private var lastTxo: Txo? = null
    private var lastObj: Obj? = null
    private var lastCont: Continue? = null

    private var splitPrevRec: BiffRec? = null
    private var splitContRec: BiffRec? = null

    /**
     * add an XLSRecord to this handler
     * check if it needs a Continue, if
     * so, put in our continued_recs.
     */
    fun addRec(rec: BiffRec, datalen: Int) {
        // check if this record has Continue records
        // if so, delay initialization until all Continues are read.
        // In the body.getData() method we then need to check
        // if the record has Continues and if so, then read/write init
        // data from them.

        val opcode = rec.opcode
        var nextOpcode: Short = 0x0

        if (opcode != XLSConstants.EOF) {
            nextOpcode = book!!.factory!!.lookAhead(rec)
        }

        if (nextOpcode == XLSConstants.CONTINUE) {
            if (DEBUGLEVEL > 11) Logger.logInfo("Next OPCODE IS CONTINUE: " + Integer.toHexString(nextOpcode.toInt()))
        }
        if (nextOpcode == XLSConstants.CONTINUE && opcode != XLSConstants.CONTINUE) { // the continued rec
            if (continued != null) {
                continued!!.init()
                continued = null
            }
            this.continued = rec
            this.splitPrevRec = this.continued
            // if the rec is a Txo, we need to process special
            if (continued is Txo) {
                this.handleTxo = true
                lastTxo = continued as Txo?
                lastObj = null
                this.handleObj = false    // ""
            } else if (continued is Obj) {
                this.handleObj = true
                lastObj = continued as Obj?
                // obj records need to be init'd before setsheet
                lastObj!!.init()
                lastTxo = null
                this.handleTxo = false    // ""
            } else {
                this.handleTxo = false
                this.handleObj = false
            }
            lastCont = null
        } else if (opcode == XLSConstants.CONTINUE) { // add to the continued rec
            splitContRec = rec
            rec.init()
            if (continued == null && lastCont == null) {
                //This is a use case where a chart is in the middle of an Obj, PLS, or Txo record.  A continue
                // record appears at the end of the last chart EOF, and needs to remain inorder to not cause corruption
                if (splitPrevRec != null) {
                    (rec as Continue).predecessor = splitContRec
                }
                if (DEBUGLEVEL > 0)
                    Logger.logWarn("Warning:  Out of spec split txo continue record found, reconstructing." + splitPrevRec!!.toString())

            } else {
                if (lastCont != null)
                    (rec as Continue).predecessor = lastCont
                else
                    (rec as Continue).predecessor = continued
                if (continued!!.opcode == XLSConstants.SST || continued!!.opcode == XLSConstants.STRINGREC) {
                    if (DEBUGLEVEL > 2) Logger.logInfo("Sst Continue.  grbit:" + rec.getGrbit())
                } else {
                    rec.hasGrbit = false // if it can't have one, don't
                }
            }

            // last data rec was a Txo -- add next 2 Continues
            if (handleTxo) {
                // txo's have either 2 continues following (text, formatting runs), no continues (an "empty" txo); may also contain continues masking mso's
                if (lastTxo!!.text == null) {
                    if (!isMaskedMSODrawingRec(rec.data))
                    // in almost all cases, the next continue is a Text continue
                        lastTxo!!.text = rec as Continue
                    else { // it is possible that an empty TXO (one that does NOT contain Text) is followed by a Continue rec which is masking an MSODrawing
                        continued = createMSODrawingFromContinue(rec)        // create a new MSODrawing rec from the Continue rec's data
                        (rec as Continue).maskedMso = continued as MSODrawing?    // set maskedMso in Continue to identify
                    }
                } else if (lastTxo!!.formattingruns == null)
                    lastTxo!!.formattingruns = rec as Continue
                else {  // a third continues: will be a masked mso or possibly a "big rec" continues
                    try {
                        if (isMaskedMSODrawingRec(rec.data)) {
                            continued = createMSODrawingFromContinue(rec)        // create a new MSODrawing rec from the Continue rec's data
                            (rec as Continue).maskedMso = continued as MSODrawing?    // set maskedMso in Continue to identify
                        } else if (continued != null) {    // then it's a "big rec" continue
                            continued!!.addContinue(rec as Continue)
                            (continued as XLSRecord).mergeContinues()    // must merge continues "by hand" because data member var is set already
                            continued!!.removeContinues()
                        }
                    } catch (e: Exception) {
                        if (DEBUGLEVEL > 0)
                            Logger.logErr("ContinueHandler.txo parsing- encountered unknown Continue record")
                    }

                    lastCont = rec as Continue
                }
            } else if (handleObj) {
                try {    // When a Continue record follows an Obj record, it is either masking a Msodrawing record - or it's a "big rec" continues
                    if (isMaskedMSODrawingRec(rec.data)) {
                        continued = createMSODrawingFromContinue(rec)        // create a new MSODrawing rec from the Continue rec's data
                        (rec as Continue).maskedMso = continued as MSODrawing?    // set maskedMso in Continue to identify
                    } else if (continued != null) {    // then it's a "big rec" continue
                        continued!!.addContinue(rec as Continue)
                        (continued as XLSRecord).mergeContinues()    // must merge continues "by hand" because data member var is set already
                        continued!!.removeContinues()
                    }
                    lastCont = rec as Continue
                } catch (e: Exception) {
                    if (DEBUGLEVEL > 0)
                        Logger.logErr("ContinueHandler.Obj parsing- encountered unknown Continue record")
                }

            } else {
                // null continued?  bizarre case testfiles/assess.xls -jm 10/01/2004
                if (continued != null)
                    continued!!.addContinue(rec as Continue)

                lastCont = rec as Continue
            }
        } else {
            if (continued != null) {
                continued!!.init()
                /* If Formula Attached String was a continue, Formula cachedValue WAS NOT init-ed. Do here.*/
                if (continued!!.opcode == XLSConstants.STRINGREC)
                    this.book!!.lastFormula!!.setCachedValue((continued as StringRec).stringVal)
                continued = null
            }

            // associate boundsheet for init
            if (book!!.lastbound != null) {
                if (/*opcode == NAME ||*/
                        opcode == XLSConstants.FORMULA)
                    rec.setSheet(book!!.lastbound)
            }
            // init names without init'ing expression -- must init expression after loading sheet recs ...
            if (opcode != XLSConstants.NAME)
                rec.init()
            else
                (rec as Name).init(false)    // don't init expression here; do after loading sheet recs ...
            lastCont = null
        }
    }

    /**
     * returns true if this id is one of an MSODrawing
     * <br></br>occurs when a Continue record is masking an MSO
     * (i.e. contains the record structure of an MSO in it's data,
     * with an opcode of Continue)
     *
     * @param data - record data
     * @return true if data is in form of an MSODrawing record
     */
    private fun isMaskedMSODrawingRec(data: ByteArray): Boolean {
        if (data.size > 3) {
            val id = 0xFF and data[3] shl 8 or (0xFF and data[2])
            return id == MSODrawingConstants.MSOFBTSPCONTAINER ||
                    id == MSODrawingConstants.MSOFBTSOLVERCONTAINER ||
                    id == MSODrawingConstants.MSOFBTSPGRCONTAINER ||
                    id == MSODrawingConstants.MSOFBTCLIENTTEXTBOX
        }
        return false
    }

    /**
     * create an MSODrwawing Record from a Continue record
     * which is masking an MSODrawing (i.e. contains the record
     * structure of an MSO in it's data, with an opcode of Continue)
     *
     * @param rec
     * @return
     */
    private fun createMSODrawingFromContinue(rec: BiffRec): MSODrawing {
        // create an mso and add to drawing recs ...
        val mso = MSODrawing()
        mso.opcode = XLSConstants.MSODRAWING
        mso.workBook = rec.workBook
        mso.setData(rec.data)
        mso.length = rec.data.size
        mso.setDebugLevel(DEBUGLEVEL)
        mso.streamer = book!!.streamer
        return mso
    }

    /**
     * clear out object references in prep for closing workbook
     */
    fun close() {
        continued = null
        lastTxo = null
        lastObj = null
        lastCont = null
        book = null

    }

    companion object {
        /**
         *
         */
        private const val serialVersionUID = 164009339243774537L
        private val DEBUGLEVEL = 0
        private val processContinues = true // debug setting

        /**
         * check if the record needs to have its data
         * spanned across Continue records.
         */
        fun createContinues(rec: BiffRec, out: OutputStream, streamer: ByteStreamer): Boolean {
            val datalen = rec.length
            val opc = rec.opcode.toInt()
            // Logger.logInfo("ContinueHandler creating output continues for: " + rec.toString() + " datalen: " + datalen);
            // if greater than 8023, we need Continues
            if (opc == XLSConstants.CONTINUE.toInt()) {
                if ((rec as Continue).isBigRecContinue)
                // skip ensuing Continues as should be written by main record
                    return true
                else if (rec.maskedMso != null && rec.maskedMso!!.length - 4 > XLSConstants.MAXRECLEN) {
                    rec.maskedMso!!.opcode = XLSConstants.CONTINUE    // so can add the correct record to output
                    createBigRecContinues(rec.maskedMso!!, out, streamer)
                    rec.maskedMso!!.opcode = XLSConstants.MSODRAWING    // reset so can continue working with this record set
                    return true    // processed, return true
                }// handle masked mso's which have continues separately
            } else if (opc == XLSConstants.SST.toInt()) {
                createSstContinues(rec as Sst, out, streamer)
                return true
            } else if (opc == XLSConstants.TXO.toInt()) {
                createTxoContinues(rec as Txo, out, streamer)
                return true
            } else if (opc == XLSConstants.MSODRAWINGGROUP.toInt()) {
                createMSODGContinues(rec, out, streamer)
                return true
            } else if (datalen - 4 > XLSConstants.MAXRECLEN) {
                createBigRecContinues(rec, out, streamer)
                return true
            }
            return false
        }


        /**
         * check if the record needs to have its data
         * spanned across Continue records.
         */
        internal fun createContinues(rec: BiffRec, insertLoc: Int): Int {
            val datalen = rec.length

            //Logger.logInfo("ContinueHandler creating output continues for: " + rec.toString() + " datalen: " + datalen);
            // if greater than 8023, we need Continues
            if (rec is Obj || rec is MSODrawing || rec is MSODrawingGroup) {
                return createObjContinues(rec)
            }
            if (rec is Sst) {
                return createSstContinues(rec, insertLoc)
            } else if (datalen > XLSConstants.MAXRECLEN && rec !is Continue) {
                return createBigContinues(rec, insertLoc)
            } else if (rec is Txo) {
                return createTxoContinues(rec)
            }
            return 0
        }


        /**
         * generate Continue records for Sst records
         */
        fun createSstContinues(rec: Sst, insertLoc: Int): Int {
            var insertLoc = insertLoc
            val dta = rec.getData()
            val datalen = dta!!.size
            if (datalen < XLSConstants.MAXRECLEN) return 0
            // get the grbits and continue sizes
            val continuedef = Sst.getContinueDef(rec, false)
            val continuesizes = continuedef[0] as Array<Int>
            val sstgrbits = continuedef[1] as Array<Byte>
            // int sstoffset =  continuesizes[0].intValue() - rec.getOrigSstLen();
            val numconts = continuesizes.size - 1

            // blow out old record Continues
            removeContinues(rec)

            // account for the offset caused by the Sstgrbits
            var sizer = 0
            var dtapos = 0

            // create Continues, skip the first which is Sst recordbody
            for (i in 1..numconts) { // start after the 1st continue length which is the Sst data body
                if (continuesizes[i].toInt() == 0) break
                // numconts is one less than continuesizes, match continuesizes[1] with numconts[0]...
                val thisgr = sstgrbits[i - 1]
                dtapos += continuesizes[i - 1].toInt()

                // check for a grbit -- null grbit means Continue Breaks on a one-char UString
                var hasGrbit = false
                if (thisgr != null)
                    hasGrbit = thisgr!!.toByte() < 0x2 && thisgr!!.toByte() >= 0x0 // Sst grbit is either 0h or 1h, otherwise it's String data
                if (continuesizes[i].toInt() == XLSConstants.MAXRECLEN) hasGrbit = false // this is a non-standard Sst Continue

                sizer = continuesizes[i].toInt()
                if (i == numconts)
                    sizer = datalen - dtapos
                if (hasGrbit)
                    sizer++
                val continuedata = ByteArray(sizer)

                if (hasGrbit) {
                    if (DEBUGLEVEL > 1) {
                        Logger.logInfo("New Continue. HAS grbit.")
                        Logger.logInfo("Continue GRBIT: " + thisgr!!)
                    }
                    continuedata[0] = thisgr!!.toByte() // set a grbit on the new Continue
                    System.arraycopy(dta, dtapos, continuedata, 1, continuedata.size - 1)
                } else {
                    if (DEBUGLEVEL > 1) {
                        Logger.logInfo("New Continue. NO grbit.")
                        Logger.logInfo("Continue GRBIT: " + (dta[dtapos] and 0x1))
                    }
                    System.arraycopy(dta, dtapos, continuedata, 0, continuedata.size)
                }

                val thiscont = addContinue(rec, continuedata, insertLoc, rec.wkbook)
                if (hasGrbit) thiscont.hasGrbit = true
                insertLoc++
            }
            val sstsize = continuesizes[0].toInt()
            trimRecSize(rec, sstsize)
            return numconts
        }

        /**
         * Write a record out to the output stream.  I have zero idea why this is in here and
         * not in bytestreamer or elsewhere.  Anyway, we are passing around the streamer in here
         * for when it's needed...
         *
         * @param rec
         * @param out
         */
        private fun writeRec(rec: BiffRec, out: OutputStream, streamer: ByteStreamer) {
            if (rec.opcode == XLSConstants.CONTINUE)
                rec.preStream()

            try { // output the rec bytes
                streamer.writeRecord(out, rec)
            } catch (a: Exception) {
                Logger.logErr("Streaming WorkBook Bytes for record:$rec failed: $a Output Corrupted.")
            }

        }

        /**
         * generate Continue records for Sst records
         */
        fun createSstContinues(rec: Sst, out: OutputStream, streamer: ByteStreamer) {
            val dta = rec.getData()
            val datalen = dta!!.size
            // get the grbits and continue sizes
            val continuedef = Sst.getContinueDef(rec, false)
            val continuesizes = continuedef[0] as Array<Int>

            var sstsize = 0
            if (continuesizes.size > 0) {
                val sstz = continuesizes[0]
                if (sstz != null)
                    sstsize = sstz!!.toInt()
                trimRecSize(rec, sstsize)
            }
            // output the original rec
            writeRec(rec, out, streamer)

            val sstgrbits = continuedef[1] as Array<Byte>
            val numconts = continuesizes.size - 1

            // blow out old record Continues
            removeContinues(rec) // are they even there? Should not be!

            // account for the offset caused by the Sstgrbits
            var sizer = 0
            var dtapos = 0

            // create Continues, skip the first which is Sst recordbody
            for (i in 1..numconts) { // start after the 1st continue length which is the Sst data body
                if (continuesizes[i].toInt() == 0) break
                // numconts is one less than continuesizes, match continuesizes[1] with numconts[0]...
                val thisgr = sstgrbits[i - 1]
                dtapos += continuesizes[i - 1].toInt()

                // check for a grbit -- null grbit means Continue Breaks on a one-char UString
                var hasGrbit = false
                if (thisgr != null)
                    hasGrbit = thisgr!!.toByte() < 0x2 && thisgr!!.toByte() >= 0x0 // Sst grbit is either 0h or 1h, otherwise it's String data
                if (continuesizes[i].toInt() == XLSConstants.MAXRECLEN) hasGrbit = false // this is a non-standard Sst Continue

                sizer = continuesizes[i].toInt()
                if (i == numconts)
                    sizer = datalen - dtapos
                if (hasGrbit) {
                    sizer++
                }
                val continuedata = ByteArray(sizer)

                if (hasGrbit) {
                    if (DEBUGLEVEL > 1) {
                        Logger.logInfo("New Continue. HAS grbit.")
                        Logger.logInfo("Continue GRBIT: " + thisgr!!)
                    }
                    continuedata[0] = thisgr!!.toByte() // set a grbit on the new Continue
                    System.arraycopy(dta, dtapos, continuedata, 1, continuedata.size - 1)
                } else {
                    if (DEBUGLEVEL > 1) {
                        Logger.logInfo("New Continue. NO grbit.")
                        Logger.logInfo("Continue GRBIT: " + (dta[dtapos] and 0x1))
                    }
                    System.arraycopy(dta, dtapos, continuedata, 0, continuedata.size)
                }

                val thiscont = createContinue(continuedata, rec.wkbook!!)
                if (hasGrbit) thiscont.hasGrbit = true
                // output the original rec
                writeRec(thiscont, out, streamer)

            }
        }


        /**
         * remove Continues from a record
         *
         *
         * TODO:  Can this be removed now?  We shouldn't really have continues in memory, just on stream, NO?
         */
        fun removeContinues(rec: BiffRec) {
            // remove existing Continues (if any)
            val oldconts = rec.continueVect ?: return
            if (oldconts.size > 0) {
                val it = oldconts.iterator()
                while (it.hasNext()) {
                    var ob: Any? = it.next()
                    rec.streamer.removeRecord((ob as BiffRec?)!!)
                    ob = null // faster!!  will it work?
                }
                rec.removeContinues()
            }
        }

        /**
         * generate Continue records for records with lots of data
         */
        fun createBigRecContinues(rec: BiffRec, out: OutputStream, streamer: ByteStreamer) {
            val dta = rec.data
            val datalen = dta.size
            val numconts = datalen / XLSConstants.MAXRECLEN

            if (datalen > XLSConstants.MAXRECLEN) {
                trimRecSize(rec, XLSConstants.MAXRECLEN)
                writeRec(rec, out, streamer)
            } else {
                writeRec(rec, out, streamer)
                return
            }

            // create Continues
            val boundaries = ContinueHandler.getBoundaries(numconts)
            var sizer = XLSConstants.MAXRECLEN
            for (i in 0 until numconts) {
                // if this is the last Continue rec it is probably shorter than CONTINUESIZE
                if (datalen - boundaries[i] < XLSConstants.MAXRECLEN)
                    sizer = datalen - boundaries[i]
                val continuedata = ByteArray(sizer)
                System.arraycopy(dta, boundaries[i], continuedata, 0, continuedata.size)
                val cr = createContinue(continuedata, rec.workBook)
                writeRec(cr, out, streamer)
            }
        }

        /**
         * generate Continue records for records with lots of data
         */
        // 20070921 KSC: Is this used?
        fun createBigContinues(rec: BiffRec, insertLoc: Int): Int {
            var insertLoc = insertLoc
            val dta = rec.data
            val datalen = dta.size
            val numconts = datalen / XLSConstants.MAXRECLEN
            // create Continues
            val conts = arrayOfNulls<Continue>(numconts)
            val boundaries = ContinueHandler.getBoundaries(numconts)
            // remove existing Continues (if any)
            removeContinues(rec)
            var sizer = XLSConstants.MAXRECLEN
            for (i in 0 until numconts) {
                // if this is the last Continue rec it is probably shorter than CONTINUESIZE
                if (datalen - boundaries[i] < XLSConstants.MAXRECLEN) sizer = datalen - boundaries[i]
                val continuedata = ByteArray(sizer)
                System.arraycopy(dta, boundaries[i], continuedata, 0, continuedata.size)
                conts[i] = addContinue(rec, continuedata, insertLoc, rec.workBook)
                insertLoc++
            }
            trimRecSize(rec, XLSConstants.MAXRECLEN)
            return numconts
        }

        /**
         * Create and initialize a new Continue record
         *
         * @param rec       the XLSRecord owner of the Continue
         * @param data      the pre-sized Continue body data
         * @param streampos the position to insert the new Continue into the data stream
         */
        fun createContinue(data: ByteArray, book: Book): Continue {
            val cont = Continue()
            cont.workBook = book as WorkBook
            cont.setData(data)
            cont.streamer = book.streamer        // 20070921 KSC: Addded
            val len = data.size
            cont.opcode = XLSConstants.CONTINUE
            cont.length = len.toShort()
            return cont
        }


        /**
         * Create and initialize a new Continue record
         *
         * @param rec       the XLSRecord owner of the Continue
         * @param data      the pre-sized Continue body data
         * @param streampos the position to insert the new Continue into the data stream
         */
        fun addContinue(rec: BiffRec, data: ByteArray, streampos: Int, book: Book?): Continue {
            val cont = Continue()
            rec.addContinue(cont)
            cont.workBook = book as WorkBook?
            cont.setData(data)
            val len = data.size
            cont.opcode = XLSConstants.CONTINUE
            cont.length = len.toShort()
            return cont
        }

        /**
         * generate the mandatory Continue records for the Txo rec type
         *
         *
         * Txo must have at least 2 Continue recs
         * first one contains text data
         * second (last) one contains formatting runs
         */
        fun createTxoContinues(rec: Txo): Int {
            val txoConts = rec.continueVect ?: return 0
// iterate through existing Continues and update their data.
            //   for(int i = 0;i< txoConts.length;i++){
            // step into first Continue and find its
            // boundary -- everything after goes into
            // the 'formatting' Continue rec.  For now
            // we assume that this can't change

            // now create data Continue(s) for text

            // we can ignore the formatting Continue
            // but text may be formatted wierd -- SORRY! : )

            // insert NEW Continues into byte stream

            //  }
            if (rec.length > XLSConstants.MAXRECLEN + 4)
                ContinueHandler.trimRecSize(rec, XLSConstants.MAXRECLEN)
            return 0//txoConts.length;

        }


        /**
         * generate the mandatory Continue records for the Txo rec type
         *
         *
         * Txo must have at least 2 Continue recs
         * first one contains text data
         * second (last) one contains formatting runs
         */
        fun createTxoContinues(rec: Txo, out: OutputStream, streamer: ByteStreamer) {
            val dta = rec.bytes
            val book = rec.workBook
            if (dta!!.size > XLSConstants.MAXRECLEN + 4) {
                ContinueHandler.trimRecSize(rec, XLSConstants.MAXRECLEN)
                // stream the rec to out...
                ContinueHandler.writeRec(rec, out, streamer)
                // and now the restof the recs
                val myconts = getContinues(dta, XLSConstants.MAXRECLEN, book)
                for (x in myconts.indices) {
                    ContinueHandler.writeRec(myconts[x], out, streamer)
                }
            } else {
                // stream the rec to out...
                ContinueHandler.writeRec(rec, out, streamer)
            }

        }

        /**
         * generate the mandatory Continue records for the Obj rec type
         */
        fun createMSODGContinues(rec: BiffRec, out: OutputStream, streamer: ByteStreamer) {
            val dta = rec.data
            val datalen = dta.size
            val numconts = datalen / XLSConstants.MAXRECLEN

            if (datalen > XLSConstants.MAXRECLEN) {
                trimRecSize(rec, XLSConstants.MAXRECLEN)
                writeRec(rec, out, streamer)
            } else {
                writeRec(rec, out, streamer)
                return
            }
            // create Continues
            val boundaries = ContinueHandler.getBoundaries(numconts)
            var sizer = XLSConstants.MAXRECLEN
            for (i in 0 until numconts) {
                // if this is the last Continue rec it is probably shorter than CONTINUESIZE
                if (datalen - boundaries[i] < XLSConstants.MAXRECLEN)
                    sizer = datalen - boundaries[i]
                if (sizer == 0)
                // reclen hits boundary exactly; ignore last continue; see ByteStreamer.writeOut for boundary issues
                    break
                val continuedata = ByteArray(sizer)
                System.arraycopy(dta, boundaries[i], continuedata, 0, continuedata.size)

                var cr: BiffRec? = null
                //          now add a second MSODG -- acts like a continue, exists to confuse and dismay

                if (i == 0) {
                    cr = MSODrawingGroup.prototype
                    cr!!.data = continuedata
                } else {
                    cr = createContinue(continuedata, rec.workBook)
                }
                writeRec(cr, out, streamer)
            }
            // reset rec
            rec.data = dta
        }

        private fun getContinues(dta: ByteArray, start: Int, book: WorkBook?): Array<Continue> {

            var clen = dta.size - start
            var len = 0
            val pos = 0
            var numconts = clen / XLSConstants.MAXRECLEN
            numconts++
            val retconts = arrayOfNulls<Continue>(numconts)

            Logger.logInfo("Creating continues: $numconts")

            var dtx: ByteArray? = null
            for (x in 0 until numconts) {
                if (clen > XLSConstants.MAXRECLEN)
                    len = XLSConstants.MAXRECLEN
                else
                    len = clen

                if (len > 0) {
                    dtx = ByteArray(len)
                } else {
                    dtx = ByteArray(clen)
                    len = clen
                }

                // populate the bytes
                System.arraycopy(dta, start, dtx, 0, len)
                retconts[x] = createContinue(dtx, book!!)
                clen -= len
            }
            return retconts
        }

        /**
         * trims the current rec to MAXRECLEN size
         */
        fun createObjContinues(rec: BiffRec): Int {
            if (rec.length > XLSConstants.MAXRECLEN + 4)
                ContinueHandler.trimRecSize(rec, XLSConstants.MAXRECLEN)
            return 0//objConts.length;

        }

        /**
         * get record size boundaries which determine the
         * break-point of Continue record data
         *
         * @param x the number of boundaries needed
         */
        internal fun getBoundaries(x: Int): IntArray {
            val boundaries = IntArray(x)
            var thisbound = 0
            for (i in 0 until x) {
                thisbound += XLSConstants.MAXRECLEN
                boundaries[i] = thisbound
            }
            return boundaries
        }

        /**
         * trim original rec to max rec size
         * (for records with size-related Continues, not Txos)
         */
        fun trimRecSize(rec: BiffRec, CONTINUESIZE: Int) {
            val dta = rec.data
            val newdata = ByteArray(CONTINUESIZE)
            System.arraycopy(dta, 0, newdata, 0, CONTINUESIZE)
            rec.data = newdata
        }
    }
}