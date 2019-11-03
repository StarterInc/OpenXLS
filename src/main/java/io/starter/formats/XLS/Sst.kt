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

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.OOXMLConstants
import io.starter.formats.OOXML.Ss_rPr
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory

import java.io.*
import java.util.*
import java.util.concurrent.atomic.AtomicInteger


/**
 * **Sst: Shared String Table 0xFCh**<br></br>
 *
 *
 * Sst records contain a table of Strings possibly spanning multiple Continue
 * Records
 *
 *
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       cstTotal    4       Total number of strings in this and the
 * EXTSST record.
 * 8       cstUnique   4       Number of unique strings in this table.
 * 12      rgb         var     Array of unique strings
 *
</pre> *
 *
 *
 * @see Sst
 *
 * @see Labelsst
 *
 * @see Extsst
 */
class Sst : io.starter.formats.XLS.XLSRecord() {
    /**
     * return the total number of strings in the SST
     *
     * @return
     */
    var numTotal = -1
        private set
    /**
     * return the number of unique strings in the SST
     *
     * @return
     */
    var numUnique = -1
        private set
    private val boundincrement = 0

    // continue handling
    /**
     * return # continues
     *
     * @return
     */
    var numContinues = -1
        private set
    private var boundaries: IntArray? = null
    private var grbits: ByteArray? = null
    private var stringvector: MutableList<*> = SstArrayList()
    private var dupeSstEntries = HashSet()
    private var existingSstEntries = HashSet()
    /**
     * set the Extsst rec for this Sst
     */
    internal var extsst: Extsst? = null
    internal var origSstLen = 0

    internal var thiscont: Continue? = null
    internal var datalen = -1
    internal var currbound = 0

    internal var deldata: ByteArray? = null

    val realOriginalSize: Int
        get() = this.originalsize

    override var data: ByteArray?
        get
        set(b) {
            if (data == null) {
                this.originalsize = b!!.size
            }
            super.setData(b)
        }

    /**
     * because the SST comes between the BOUNDSHEET records and all BOUNDSHEET
     * BOFs, the lbPlyPos needs to change for all of them when record size
     * changes.
     */
    val updatesAllBOFPositions: Boolean
        get() = true

    internal var retpos = -1

    internal var cbounds: CompatibleVector? = CompatibleVector()
    internal var sstgrbits: CompatibleVector? = CompatibleVector()
    internal var lastwasbreakable = true
    internal var stringisonbound = false
    internal var laststringwasonbound = false
    internal var islast = false
    internal var thisbounds = WorkBookFactory.MAXRECLEN
    internal var lastbounds = 0
    internal var contcounter = 0
    internal var lastlen = 0
    internal var grbitct = 0
    internal var dl = 0
    internal var leftoverlen = 0
    internal var gr: Byte = 0x0

    /*
     * handle the checking of Continue boundary strings
     */
    internal var stringnumber = 0
    internal var continuenumber = -1
    internal var lastgrbit = 0

    internal var continueDef: Array<Any>

    // Optimization -- don't check UStr on add
    internal var STRING_ENCODING_MODE = Sst.STRING_ENCODING_UNICODE

    /**
     * Returns the String vector
     */
    val stringVector: List<*>
        get() = this.stringvector

    /**
     * Returns all strings that are in the SharedStringTable for this workbook.
     * The SST contains all standard string records in cells, but may not
     * include such things as strings that are contained within formulas. This
     * is useful for such things as full text indexing of workbooks
     *
     * @return Strings in the workbook.
     */
    val allStrings: ArrayList<*>
        get() {
            val al = ArrayList(stringvector.size)
            for (i in stringvector.indices) {
                al.add(stringvector[i].toString())
            }
            return al
        }

    /**
     * Returns the length of this record, including the 4 header bytes
     */
    override// if "hasGrbit" must account for additional size taken up by it
    // see ContinueHandler.createSstContinues
    // Sst grbit is either 0h
    // or 1h, otherwise it's
    // String data
    var length: Int
        get() {
            var len = super.length
            for (i in 0 until sstgrbits!!.size - 1) {
                val b = sstgrbits!![i] as Byte
                if (b != null) {
                    if (b < 0x2 && b >= 0x0)
                        len++
                }
            }
            return len
        }
        set

    /**
     * removes all existing Continues from the Sst
     */
    override fun removeContinues() {
        super.removeContinues()
        this.continues = null
        thiscont = null
    }

    /**
     * initialize the Continue Handling counters
     */
    private fun initContinues() {
        // create the array of record boundary offsets
        // this allows us to detect spanning UNICODE strings in CONTINUE
        // records...
        datalen = this.length
        numContinues = this.continueVect.size
        boundaries = IntArray(numContinues + 1)
        var thisbound = 0
        var ir = 0
        grbits = ByteArray(numContinues)
        // continues = this.getContinueVect();
        val it = this.continues!!.iterator()
        while (it.hasNext()) {
            val ci = it.next() as Continue
            // byte[] b = ci.getData(); // REMOVE
            grbits[ir++] = ci.getGrbit()
            this.streamer!!.removeRecord(ci) // remove existing continues
            // from stream
        }
        thisbound = this.realOriginalSize
        boundaries[0] = thisbound
        var lastcontlen = 0
        for (i in 1 until boundaries!!.size) {
            val cxi = continues!![i - 1] as Continue
            var contlen = cxi.length
            if (cxi.hasGrbit)
                contlen--
            thisbound += contlen
            // if(DEBUGLEVEL > 5)Logger.logInfo( contlen + ",");
            lastcontlen += contlen
            datalen += contlen
            boundaries[i] = thisbound - 4 * i
            cxi.continueOffset = boundaries!![i - 1]
        }
        // if(DEBUGLEVEL > 5) Logger.logInfo("");
        // for(int i = 1;i<boundaries.length;i++){
        // int contlen =continues[i-1].getLength();
        // if(DEBUGLEVEL > 5) Logger.logInfo("0x" + continues[i-1].getGrbit() +
        // ",");
        // }
        // if(DEBUGLEVEL > 5) Logger.logInfo("");
        thisbound = 0
    }

    override fun init() {
        if (originalsize == 0)
            originalsize = reclen
        Sst.init(this)
    }

    /**
     * retrieves the sst string at the location pos and returns the next
     * position
     *
     * @param ustrLen    actual unicode string length, (not including formatting runs,
     * phonetic data or double byte multiplication)
     * @param pos        position in the source data buffer
     * @param ustrStart  start of unicode string within single sst record
     * @param cchExtRst  phonetic data length
     * @param runlen     fomratting run length
     * @param doublebyte true if unicode string data is double byte (and then the size
     * of the unicode string data array is ustrLen * 2)
     * @return
     */
    internal fun initUnicodeString(ustrLen: Int, pos: Int, ustrStart: Int, cchExtRst: Int, runlen: Int, doublebyte: Boolean): Int {
        val bufferBoundary = boundaries!![currbound] // get the current boundary
        val totalStrLen = ustrStart + ustrLen + cchExtRst + runlen        // calculate the total byte length of the unicode string
        val posEnd = pos + totalStrLen // end position -- if > current record length must span and access next record/continues
        val uLen = AtomicInteger(ustrLen) // same as ustrLen but mutable in order to allow changing value in getData method

        // begin checking string against current record buffer boundary
        if (posEnd < bufferBoundary) {// string does not cross current boundary - easy! retrieve totalStringLen bytes and create unicode string
            val newStringBytes = getData(uLen, pos, ustrStart, cchExtRst, runlen, doublebyte, false)
            this.initString(newStringBytes, pos, false)
            return posEnd
        } else if (posEnd == bufferBoundary) {// string is on the boundary - easy!
            if (this.numContinues == 0 || this.numContinues == this.contcounter) {
                if (DEBUGLEVEL > 5) Logger.logInfo("Last String in SST encountered.")
            }
            val newStringBytes = getData(uLen, pos, ustrStart, cchExtRst, runlen, doublebyte, false)
            this.initString(newStringBytes, pos, false)

            /* "If fHighByte is 0x1 and rgb is extended with a Continue record the break
			   MUST occur at the double-byte character boundary."
			*/  // because we ended on a string, there is no grbit on the next continue
            if (this.continues!!.size > this.currbound) {
                thiscont = this.continues!![this.currbound] as Continue
                currbound++
                if (thiscont!!.hasGrbit) {
                    thiscont!!.hasGrbit = false
                    this.shiftBoundaries(1)
                }
            }
            return posEnd
        }

        // spans or crosses the continue boundary
        val newStringBytes = getData(uLen, pos, ustrStart, cchExtRst, runlen, doublebyte, true)    // retrieve the bytes, accounting for spanning (true)
        this.initString(newStringBytes, pos, false)
        return pos + (uLen.toInt() + ustrStart + cchExtRst + runlen)        // in most cases should be same as pos + totalStrLen but it's possible for uLen to be changed in getData
    }

    /**
     * gets the sst string at strpos length allstrlen and returns the next
     * position
     *
     * @param allstrlen
     * @param strpos
     * @param strend
     * @param STATE
     * @return
     *
     *
     * KSC: replacing int getString(int allstrlen, int strpos, int
     * strend, int[] STATE){ byte[] newStringBytes = getData(allstrlen,
     * strpos, STATE); int nextpos = strpos + allstrlen;
     * if(STATE[SI_SPANSTATE]==Sst.STATE_EXRSTSPAN){
     * this.initString(newStringBytes, strpos,true); }else{
     * this.initString(newStringBytes, strpos,false); }
     *
     * if(strend != nextpos) Logger.logWarn(
     * "Sanity Check in Sst initUnicodeString(): strend != nextpos.");
     *
     * return nextpos; }
     */

    /**
     * Adjust the boundary pointers based on whether we need to compensate for
     * grbit anomalies
     *
     *
     * NOT USED
     *
     * void shiftBoundariesX(int x) {
     * int ct = 0;
     * Iterator it = continues.iterator();
     * while (it.hasNext()) {
     * Continue nextcont = (Continue) it.next();
     * if (ct++ >= this.currbound) {
     * nextcont.setContinueOffset(nextcont.getContinueOffset() + x);
     * boundaries[ct] = nextcont.getContinueOffset();
     * if (DEBUGLEVEL > -5)
     * Logger.logInfo("Sst.shiftBoundaries() Updated " + nextcont
     * + " : " + nextcont.getContinueOffset());
     * }
     * }
     * if (boundaries.length == (this.continues.size() + 1)) {
     * boundaries[this.continues.size()] += x;
     * }
     * } */

    /**
     * Adjust the boundary pointers based on whether we need to compensate for
     * grbit anomalies
     */
    internal fun shiftBoundaries(x: Int) {
        // int ret = 0;
        for (t in currbound until this.continues!!.size) {
            val nextcont = this.continues!![t] as Continue
            nextcont.continueOffset = nextcont.continueOffset + x
            boundaries[t] = nextcont.continueOffset
            if (DEBUGLEVEL > 5)
                Logger.logInfo("Sst.shiftBoundaries() Updated " + nextcont
                        + " : " + nextcont.continueOffset)
        }
        if (boundaries!!.size == this.continues!!.size + 1) {
            boundaries[this.continues!!.size] += x
        }
    }

    /**
     * Refactoring Continue data access
     *
     * @param i
     * @return
     */
    internal fun getNextStringDefData(start: Int): ShortArray {
        var start = start
        val ret = shortArrayOf(0x0.toShort(), 0x0.toShort())
        try {
            // int thiscont = -1;
            val end = start + 3
            if (end <= boundaries!![0]) { // it's in the main Sst data
                ret[0] = ByteTools.readShort(getByteAt(start++).toInt(),
                        getByteAt(start++).toInt())
                ret[1] = getByteAt(start).toShort()
                return ret
            }

            // KSC: no need as Sst.getData increments correctly ...
            // this.thiscont = this.getContinue(end);
            val b0 = this.thiscont!!.getByteAt(start++)
            val b1 = this.thiscont!!.getByteAt(start++)
            val b2 = this.thiscont!!.getByteAt(start++)
            ret[0] = ByteTools.readShort(b0.toInt(), b1.toInt())
            ret[1] = b2.toShort()
        } catch (e: Exception) {
            if (DEBUGLEVEL > 0)
                Logger.logWarn("possible problem parsing String table getting next string def data: $e")
        }

        return ret
    }

    /**
     * return the continue that contains up to t length
     *
     * @param t
     * @return
     */
    internal fun getContinue(t: Int): Continue? {
        if (t - 1 == datalen) {
            return this.continues!![this.continues!!.size - 1] as Continue
        }
        for (x in this.boundaries!!.indices.reversed()) {
            if (t > boundaries!![x]) {
                return this.continues!![x] as Continue
            }
        }
        return null
    }

    /**
     * get the string data from the proper place (either this sst record, or one or more continues)
     *
     * <br></br>
     * NOTES: if the string spans a continue, the length in bytes of each part
     * is contained in the first two ints.
     *
     *
     * if the record spans a Continue, we need to see if the border falls within
     * text data or extra data
     *
     *
     * 10 len, really 15 bytes uncomp ||gr comp 2,0,1,0,2,0,3,0,3,0,||0
     * ,4,5,5,7,8
     *
     * @param ustrLen    unicode string length
     * @param pos        position in buffer
     * @param ustrStart  start of unicode string part (after initial length(s))
     * @param cchExtRst  phonetic data length or 0 if none
     * @param runlen     formatting runs length or 0 if none
     * @param doublebyte true of doublebyte
     * @param bSpans     true if spans records
     * @return byte[] defining unicode string
     */
    internal fun getData(ustrLen: AtomicInteger, pos: Int, ustrStart: Int, cchExtRst: Int, runlen: Int, doublebyte: Boolean, bSpans: Boolean): ByteArray? {
        var pos = pos
        var totalStrLen = ustrStart + ustrLen.toInt() + cchExtRst + runlen
        var posEnd = pos + totalStrLen // buffer end position

        if (posEnd <= boundaries!![0]) { // it's in the main Sst data just grab string and return
            return this.getBytesAt(pos, totalStrLen)
        }

        // if it's in the current continues without spanning, just get the bytes and return
        if (!bSpans) { // Simple -- no Span, return data
            pos += thiscont!!.grbitoff
            val thisoff = pos - thiscont!!.continueOffset
            return thiscont!!.getBytesAt(thisoff, totalStrLen)
        }

        // if string spans two or more records, must deal with boundaries and grbits and lots of complications ...
        val bufferBoundary = boundaries!![currbound] // get the current boundary
        if (DEBUGLEVEL > 5)
            Logger.logInfo("Crossing Boundary: " + bufferBoundary
                    + ".  Double-Bytes: " + doublebyte)

        // get ensuing record (previous==thiscont.predecessor)
        if (currbound < continues!!.size) {
            thiscont = continues!![currbound++] as Continue
        }

        // find out where break is
        var currpos = pos + totalStrLen
        var bfoundBreak = false
        var bUnCompress = false // true if string on previous boundary must be uncompressed
        var bUnCompress1 = false // true if string1 must be uncompressed ** this one is confusing but works **

        // check if break is in ExtRst data (==phonetic data)
        if (cchExtRst > 0) {
            currpos -= cchExtRst
            if (currpos <= bufferBoundary) {
                if (DEBUGLEVEL > 5)
                    Logger.logInfo("Continue Boundary in ExtRst data.")
                if (thiscont!!.hasGrbit) {
                    thiscont!!.hasGrbit = false
                    this.shiftBoundaries(1)
                }
                bfoundBreak = true
            }
        }

        // check if break is in formatting run data
        if (runlen > 0) {
            currpos -= runlen
            if (!bfoundBreak && currpos <= bufferBoundary) { // check against japanese!
                if (DEBUGLEVEL > 5)
                    Logger.logInfo("Continue Boundary in Formatting Run data.")
                if (thiscont!!.hasGrbit) {
                    this.shiftBoundaries(1)
                    thiscont!!.hasGrbit = false
                }
                bfoundBreak = true
            }
        }

        // otherwise the break is in unicode stringdata part
        currpos = pos + ustrStart
        if (!bfoundBreak && currpos < bufferBoundary) {
            if (ustrLen.toInt() == 0) { // a ONE BYTE String on the boundary! Add the grbit back to the Continue
                if (DEBUGLEVEL > 5)
                    Logger.logInfo("1 byte length String on the Continue Boundary.")
                boundaries!![boundaries!!.size - 1]++ // increment the last boundary...
            }
        }

        // check if break is within the actual ustring data
        if (currpos <= bufferBoundary && currpos + ustrLen.toInt() > bufferBoundary) { // is break within String portion// ?
            if (DEBUGLEVEL > 5)
                Logger.logInfo("Continue Boundary in String data.")
            if (!thiscont!!.hasGrbit) { // when does this happen???
                thiscont!!.hasGrbit = true
                this.shiftBoundaries(-1)
            }

            val b = thiscont!!.getGrbit()
            // If it changes double --> single or single --> double then adjust accordingly (plus set bUnCompress or bUnCompress1 flags which govern how bytes are accessed)
            if (doublebyte && b.toInt() == 0x0) { // it is in doublebytes but it really should be compressed
                val preBoundaryBytes = bufferBoundary - pos - ustrStart
                var postBoundaryBytes = pos + ustrStart + ustrLen.toInt() - bufferBoundary
                postBoundaryBytes = postBoundaryBytes / 2
                ustrLen.set(preBoundaryBytes + postBoundaryBytes)
                bUnCompress1 = true // dunno what this means but it works ...
            } else if (!doublebyte && b.toInt() == 0x1) { // string portion on prevouos boundary should be uncompressed/converted to doublebyte
                val preBoundaryBytes = bufferBoundary - pos - ustrStart
                var postBoundaryBytes = pos + ustrStart + ustrLen.toInt() - bufferBoundary
                postBoundaryBytes = postBoundaryBytes * 2
                ustrLen.set(preBoundaryBytes + postBoundaryBytes)
                bUnCompress = true
            }
            // ustrLen may have changed above - reset vars
            totalStrLen = ustrStart + ustrLen.toInt() + cchExtRst + runlen
            posEnd = pos + totalStrLen
        }

        // calculate length on current record (=thiscont.predecessor) and length on ensuing continues
        var string1ByteLength = pos - this.thiscont!!.continueOffset // bytes on 1st record or continue
        if (string1ByteLength < 0)
            string1ByteLength *= -1
        var string2ByteLength = totalStrLen - string1ByteLength // bytes on 2nd// or// ensuing// continues// ==// spanned// bytes
        var extraData = cchExtRst + runlen // non-string-data (phonetic info// and/or formatting runs)

        // remove ExtRst and runlen info from 2nd String length
        string2ByteLength -= extraData
        if (string2ByteLength <= 0) {
            // if it spans we want extr just to be the bytes on the second continue
            extraData = string2ByteLength + extraData
            string2ByteLength = 0// all String data is contained in prior // Continue
        }

        var string1bytes: ByteArray? = null
        var string2bytes: ByteArray? = null

        // we need to expand the first section bytes to fit the last one (????)
        if (this.thiscont!!.predecessor is Continue) {
            pos = this.thiscont!!.predecessor!!.length - string1ByteLength
            pos -= 4
        }
        if (this.thiscont!!.hasGrbit) {
            thiscont!!.grbitoff = 1
        } else {
            thiscont!!.grbitoff = 0
        }

        // *********************************************************************************************************
        // handle the part of unicode string on previous continues
        if (!bUnCompress) {
            string1bytes = this.thiscont!!.predecessor!!.getBytesAt(pos, string1ByteLength)
        } else { // portion of string on previous boundary is singlebyte; ensuing portion is doublebyte; must convert previous to doublebyte
            string1bytes = convertCompressedBytesToDoubleBytes(pos, string1ByteLength, ustrStart)
        }

        // *********************************************************************************************************
        // handle part on current (and ensuing, if necessary) continues
        if (string2ByteLength < XLSConstants.MAXRECLEN) { // 99.9% usual case
            if (!bUnCompress1) {
                string2bytes = thiscont!!.getBytesAt(0 + thiscont!!.grbitoff, string2ByteLength)
            } else { // Expand the second string bytes
                string2ByteLength *= 2
                string2bytes = ByteArray(string2ByteLength)
                for (t in 0 until string2ByteLength / 2) {
                    string2bytes[t * 2] = this.thiscont!!.getByteAt(t + this.thiscont!!.continueOffset)
                }
            }
            // since we've accessed the last bytes of the prior Continue, blow
            // it out!
            if (this.thiscont!!.predecessor is Continue) this.thiscont!!.predecessor!!.data = null
        } else { // string2ByteLength spans continue(s) ************************************************* see infoteria/cannotread824315.xls
            var blen = string2ByteLength
            var idx = 0
            var start = 0
            string2bytes = ByteArray(blen)
            while (blen > 0) {                // loop thru ensuing continues until correct length is read in
                var curlen = Math.min(start + thiscont!!.length - thiscont!!.grbitoff, blen)
                if (!bUnCompress1) {
                    val tmp = thiscont!!.getBytesAt(start + thiscont!!.grbitoff, curlen)
                    System.arraycopy(tmp!!, 0, string2bytes, idx, curlen)
                } else { // Expand the second string bytes - NOTE: This has not been hit so hasn't been tested ...
                    curlen *= 2
                    val tmp = ByteArray(curlen)
                    for (t in 0 until curlen / 2) {
                        tmp[t * 2] = this.thiscont!!.getByteAt(t + this.thiscont!!.continueOffset)
                    }
                    System.arraycopy(tmp, 0, string2bytes, idx, curlen)
                }
                // since we've accessed the last bytes of the prior Continue, blow it out!
                if (this.thiscont!!.predecessor is Continue) this.thiscont!!.predecessor!!.data = null
                if (curlen >= thiscont!!.length - thiscont!!.grbitoff) { // finished this one, get next continue
                    if (currbound < continues!!.size)
                        thiscont = continues!![currbound++] as Continue
                    if (this.thiscont!!.hasGrbit)
                        thiscont!!.grbitoff = 1
                    else
                        thiscont!!.grbitoff = 0
                    start = 4        // don't understand this but, hey, it works ...
                } else {    // we are done
                    // get current length in current continues only (==start postion for extra data, if any)
                    string2ByteLength = curlen + start
                    break
                }
                idx += curlen
                blen -= curlen
            }
        }

        // ***********************************************************************************************************
        // now put together the string bytes - string1bytes and string2bytes are the ustring only, excluding extraData (formatting or phoentic data)
        val returnstringbytes = ByteArray(string1bytes!!.size + string2bytes!!.size + extraData)
        System.arraycopy(string1bytes, 0, returnstringbytes, 0, string1bytes.size)
        System.arraycopy(string2bytes, 0, returnstringbytes, string1bytes.size, string2bytes.size)

        // does it have ExtRst or Formatting Runs?
        if (extraData > 0) {
            if (posEnd <= boundaries!![currbound]) { // usual case!!
                var startpos = string2ByteLength
                if (bUnCompress1) startpos /= 2
                if (string2ByteLength != 0) startpos += this.thiscont!!.grbitoff        // dunno why but it works ...
                val rx2 = thiscont!!.getBytesAt(startpos, extraData)
                System.arraycopy(rx2!!, 0, returnstringbytes, string1bytes.size + string2bytes.size, extraData)
            } else {    // extraData spans continues ...
                // have to get portion on prev continue and rest on next continue ... sigh ....
                var startpos = string2ByteLength
                if (bUnCompress1) startpos /= 2
                val rx2 = ByteArray(extraData)
                string1ByteLength = this.thiscont!!.length - startpos // bytes on 1st record or continue
                string2ByteLength = extraData - string1ByteLength // bytes on 2nd// or// ensuing// continues// ==// spanned// bytes
                if (currbound < continues!!.size) {
                    thiscont = continues!![currbound++] as Continue
                    if (this.thiscont!!.hasGrbit) {
                        thiscont!!.grbitoff = 1
                    } else {
                        thiscont!!.grbitoff = 0
                    }
                }
                pos = this.thiscont!!.predecessor!!.length - string1ByteLength
                pos -= 4
                val start = 4    // why?????
                System.arraycopy(this.thiscont!!.predecessor!!.getBytesAt(pos, string1ByteLength), 0, rx2, 0, string1ByteLength)
                System.arraycopy(thiscont!!.getBytesAt(start + thiscont!!.grbitoff, string2ByteLength)!!, 0, rx2, string1ByteLength, string2ByteLength)
                System.arraycopy(rx2, 0, returnstringbytes, string1bytes.size + string2bytes.size, extraData)
                ustrLen.set(ustrLen.get() - 1)    // ???? DO NOT UNDERSTAND THIS BUT IT APPEARS TO WORK - hits on infoteria/cannotread824315.xls
            }
        }

        if (DEBUGLEVEL > 23)
            Logger.logInfo("Total Length from Continue: " + returnstringbytes.size)
        return returnstringbytes
    }

    /**
     * for rare occurrences where string portion on previous boundary is flagged singlebyte/compressed,
     * and the ensuing continue is flagged doublebyte; in these cases the unicode-string-portion on the
     * previous boundary must be converted to doublebyte
     *
     * @param pos       positon on previous boundary
     * @param totallen  total string length on previous boudary
     * @param uStrStart start of unicode string portion
     * @return
     */
    internal fun convertCompressedBytesToDoubleBytes(pos: Int, totallen: Int, uStrStart: Int): ByteArray {
        val uLenOnPrevious = totallen - uStrStart    // unicode string portion on previous boundary
        val converted = ByteArray(uStrStart + uLenOnPrevious * 2)
        System.arraycopy(this.thiscont!!.predecessor!!.getBytesAt(pos, uStrStart), 0, converted, 0, uStrStart)
        val ustr = this.thiscont!!.predecessor!!.getBytesAt(pos + uStrStart, uLenOnPrevious)
        converted[2] = (converted[2] or 0x1).toByte()    // flag as doublebyte/uncompressed for unicode string processing
        for (i in 0 until uLenOnPrevious) {    // copy rest of unicode string portion on prev boundary as doublebyte
            converted[uStrStart + i * 2] = ustr[i]
        }
        return converted

    }

    /**
     * given unicode bytes, create a Unicodestring and add it to the string vector
     */
    internal fun initString(newStringBytes: ByteArray?, strpos: Int,
                            extrstbrk: Boolean): Unicodestring {
        // create a new Unicodestring, set its data
        val newString = Unicodestring()
        newString.sstPos = strpos
        newString.init(newStringBytes!!, extrstbrk)

        // add the new String to the String table and return the new pointer
        if (DEBUGLEVEL > 5)
            Logger.logInfo(" val: $newString")
        if (newString.len == 0) {
            Logger.logInfo("Adding zero-length string!")
        } else {
            this.putString(newString)
        }
        return newString
    }

    private fun putString(newString: Unicodestring): Int {
        ++retpos
        (stringvector as SstArrayList).put(newString, Integer.valueOf(
                retpos))
        return retpos
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        cbounds!!.removeAllElements()
        sstgrbits!!.removeAllElements()
        stringvector.clear()
        stringvector = SstArrayList()
        dupeSstEntries.clear()
        dupeSstEntries = HashSet()
        existingSstEntries.clear()
        existingSstEntries = HashSet()
    }

    /**
     * call this method after changing the value of an SST Unicode string to
     * update the underlying SST byte array.
     */
    internal fun updateUnicodestrings() {
        // TODO: OPTIMIZE: check that the Sst has changed
        // reset defaults
        cbounds = CompatibleVector()
        sstgrbits = CompatibleVector()
        lastwasbreakable = true
        stringisonbound = false
        laststringwasonbound = false
        islast = false
        thisbounds = WorkBookFactory.MAXRECLEN
        lastbounds = 0
        contcounter = 0
        lastlen = 0
        grbitct = 0
        dl = 0
        leftoverlen = 0
        gr = 0x0

        // loop through the strings and copy their
        // bytes to the SST byte array.
        // byte[] tmp = new byte[0];
        val cstot = ByteTools.cLongToLEBytes(numTotal)
        val cstun = ByteTools.cLongToLEBytes(numUnique)

        // TODO: OPTIMIZE!!
        val out = ByteArrayOutputStream()
        try {
            out.write(cstot)
            out.write(cstun)
        } catch (e: IOException) {
            Logger.logInfo("Exception getting String bytes: $e")
        }

        if (this.stringvector.size > 0) {
            // now get the continue boundaries
            var thispos = 8
            var lastpos = 0
            cbounds!!.removeAllElements()
            sstgrbits!!.removeAllElements()
            var strb: ByteArray? = null
            val it = stringvector.iterator()
            while (it.hasNext()) {
                val ob = it.next()
                val str = ob as Unicodestring

                // from updateUnicodeStrings()
                str.sstPos = thispos

                strb = str.read()
                try {
                    out.write(strb!!)
                } catch (e: IOException) {
                    Logger.logInfo("Exception getting String bytes: $e")
                }

                lastpos = thispos
                thispos = lastpos + strb!!.size
                checkOnBoundary(str, lastpos, thispos + 4, strb) // add 4
                // because 4
                // added to
                // boundaries
            }

            if (leftoverlen > 0) {// there was leftover data!
                cbounds!!.add(Integer.valueOf(leftoverlen))
                sstgrbits!!.add(gr)
            }
            if (cbounds!!.size > 0)
                numContinues = cbounds!!.size - 1
            else
                numContinues = 0

            val bb = out.toByteArray()
            if (sanityCheck(bb.size)) {
                this.data = bb
            } else {
                this.datalen = bb.size + 4
                this.updateUnicodestrings()
            }

        }
        if (DEBUGLEVEL > 15 && cbounds != null) {
            for (t in cbounds!!.indices)
                Logger.logInfo((cbounds!![t] as Int).toInt().toString() + ",")
            Logger.logInfo("")
        }
        if (DEBUGLEVEL > 150 && sstgrbits != null) {
            for (t in sstgrbits!!.indices)
                Logger.logInfo("0x" + (sstgrbits!![t] as Byte).toByte()
                        + ",")
            Logger.logInfo("")
        }
    }

    /**
     * Checks that the continues make sense. In some cases the continue lengths
     * will be wrong due to the datalen being off. Datalen is created via
     * offsets, not absolutes, so this can occur. If so, we will reset the
     * datalen to the correct number.
     *
     * @return
     */
    private fun sanityCheck(realLen: Int): Boolean {
        var contLens: Long = 0
        for (i in cbounds!!.indices) {
            val intgr = cbounds!![i] as Int
            contLens += intgr.toLong()
        }
        if (datalen - 4 - contLens > 8223) {
            if (DEBUGLEVEL > 1)
                Logger.logWarn("SST continue lengths not correct, regenerating")
            return false
        }
        return true
    }

    internal fun checkOnBoundary(str: Unicodestring, lastpos: Int, thispos: Int,
                                 strb: ByteArray) {
        var lastpos = lastpos
        if (false)
            Logger.logInfo("Checking Sst boundary: " + lastpos + "/" + thispos
                    + ":" + thisbounds + " ContinueNumber: " + continuenumber
                    + " StringNumber:" + stringnumber++ + "numboundaries"
                    + cbounds!!.size)

        while (thispos >= thisbounds) {
            continuenumber++
            if (DEBUGLEVEL > 5)
                Logger.logInfo(thisbounds.toString())

            // check whether the string can safely be split
            val breaksok = str.isBreakable(thisbounds)
            var contlen = 0

            // get the Continue grbit
            gr = strb[0] // default is a non-grbit -- if it doesn't break, we
            // don't want it

            if (breaksok)
                gr = this.getContinueGrbitFromString(str)
            if (DEBUGLEVEL > 5)
                Logger.logInfo(" String @: " + thispos + " is breakable: "
                        + breaksok)

            // deal with string break subtleties
            contlen = WorkBookFactory.MAXRECLEN // the default
            if (islast) {
                contlen = leftoverlen
                leftoverlen = 0
                contlen++
                if (!lastwasbreakable)
                    contlen++
            } else if (!breaksok) {
                stringisonbound = true // we are
                contlen = lastpos - lastbounds
                lastbounds = lastpos
            } else {
                // check if it's double byte, if so, make sure that the break is
                // not in the middle of a character.
                if (breaksok && gr.toInt() == 1) {
                    if (str.charBreakOnBounds(thisbounds + lastgrbit)) {
                        contlen--
                    }
                }
            }

            // set continue length
            if (!laststringwasonbound && lastwasbreakable
                    && contcounter > 0) { // normal w/grbit
                if (!breaksok) {
                    cbounds!!.add(Integer.valueOf(contlen))
                } else {
                    if (!islast)
                        thisbounds--
                    cbounds!!.add(Integer.valueOf(contlen - 1))
                }
            } else {
                cbounds!!.add(Integer.valueOf(contlen))
            }

            // set grbit add null if the Continue should not have a grbit
            if (str.cch < 2) {
                sstgrbits!!.add(null)
                lastgrbit = 0
            } else if (!breaksok && gr < 0x2 && gr >= 0x0) {
                sstgrbits!!.add(null)
                lastgrbit = 0
            } else {
                sstgrbits!!.add(gr)
                lastgrbit = 1
            }

            contcounter++

            // reset stuff
            lastwasbreakable = breaksok
            laststringwasonbound = stringisonbound
            stringisonbound = false

            lastlen = contlen
            if (breaksok)
                lastbounds = thisbounds

            // datalen will be smaller than reclen
            // if continues were not created
            if (reclen > datalen)
                dl = this.reclen
            else
                dl = this.datalen
            // 20060518 KSC: handle segments that fall between the extra 4 added
            // to the boundary ...
            if (thisbounds + contlen + 4 < dl) { // not the last one
                thisbounds += contlen
                lastpos += contlen // 20090407 KSC: If !breaksok but still
                // loops, must increment lastpos or infinite
                // loops [BUGTRACKER 2355 Infoteria OOM]
            } else if (!islast) {
                leftoverlen = dl + 4 - lastlen
                if (!lastwasbreakable && leftoverlen > 0)
                    leftoverlen++
                thisbounds = dl
                islast = true
            } else {
                thisbounds += contlen
            }

        }
    }

    /**
     * This returns the Continue record grbit which is either 0 or 1 -- NOT the
     * string's grbit which determines much more...
     */
    internal fun getContinueGrbitFromString(str: Unicodestring): Byte {
        var grb: Byte = 0x0
        when (str.grbit) {
            0x1 -> grb = 0x1
            0x5 -> grb = 0x1
            0x9 -> grb = 0x1
            0xd -> grb = 0x1
        }
        return grb
    }

    /**
     * Called from LabelSst on initialization from a new workbook, this
     * pre-populates the list of strings that are currently shared.
     */
    internal fun initSharingOnStrings(isst: Int) {
        val iSst = Integer.valueOf(isst)
        if (existingSstEntries.contains(iSst)) {
            // really is just a switch -doesn't track # times string is shared ...
            dupeSstEntries.add(iSst)
        } else {
            existingSstEntries.add(iSst)
        }
    }

    fun setStringEncodingMode(mode: Int) {
        this.STRING_ENCODING_MODE = mode
    }

    /**
     * remove a Unicodestring from the table
     *
     * @param idx
     */
    internal fun removeUnicodestring(str: Unicodestring) {
        this.stringvector.remove(idx)
        this.retpos--
        this.reclen -= str.len
    }

    /**
     * used when modifying existing sst entry update data + rec lens
     *
     * @param delta amt of adjustment
     */
    internal fun adjustSstLength(delta: Int) {
        this.reclen += delta
        this.datalen += delta
    }

    /**
     * insert a new Unicodestring into the array of strings composing this
     * String Table
     */
    internal fun insertUnicodestring(us: Unicodestring): Int {
        var retpos = -1
        numTotal++
        val isuni = false
        // get the existing position of this string
        // but only if we're not ignoring dupes
        if (this.workBook!!.isSharedupes) {
            retpos = (stringvector as SstArrayList).find(us) // indexOf will not
            // match entire
            // unicode
            // string
            // (including
            // formatting)
        }
        if (retpos == -1) { // unicode string isn't in yet
            numUnique++
            val strlen = us.len
            reclen += strlen + if (us.isRichString) 5 else 3
            datalen += strlen + if (us.isRichString) 5 else 3
            if (isuni) {
                reclen += strlen // utf double encoding.
                datalen += strlen
            }
            retpos = this.putString(us)
        } else {
            // this is a duplicate string, track it!
            dupeSstEntries.add(Integer.valueOf(retpos))
        }

        return retpos
    }

    /**
     * create a new unicode string from string and formatting information and
     * add it to the Sst string array formatting runs, if present, contain list
     * of short[] {char index, font index} where char index is start index in
     * the string to apply font at font index
     *
     * @param s
     * @param formattingRuns
     * @return
     */
    internal fun addUnicodestring(s: String, formattingRuns: ArrayList<*>?): Int {
        numTotal++
        numUnique++
        val str = createUnicodeString(s, formattingRuns,
                STRING_ENCODING_MODE)

        reclen += str!!.len
        datalen += str.len
        retpos = this.putString(str)
        return retpos
    }

    /**
     * insert a new Unicodestring into the array of strings composing this
     * String Table
     */
    internal fun insertUnicodestring(s: String): Int {
        var retpos = -1
        // get the existing position of this string
        // but only if we're not ignoring dupes
        if (this.workBook!!.isSharedupes) {
            retpos = stringvector.indexOf(s)
            if (retpos > -1) {
                val str = stringvector[retpos] as Unicodestring
                if (str.hasFormattingRuns())
                    retpos = -1 // do not match if there are formatting runs
                // embedded
            }
        }

        if (retpos == -1) { // it's a new string
            retpos = addUnicodestring(s, null) // add with no formatting
            // information
        } else {
            numTotal++
            // this is a duplicate string, track it!
            dupeSstEntries.add(Integer.valueOf(retpos))
        }

        return retpos
    }

    /**
     * Determine if the isst passed in is for a duplicate string or not.
     */
    internal fun isSharedString(sstLoc: Int): Boolean {
        return dupeSstEntries.contains(Integer.valueOf(sstLoc))
    }

    /**
     * Return the Unicodestring at the corresponding index
     */
    internal fun getUStringAt(i: Int): Unicodestring {
        return stringvector[i] as Unicodestring
    }

    /**
     * find this unicode string (including formatting) in stringarray
     *
     * @param us
     * @return
     */
    internal fun find(us: Unicodestring): Int {
        return (stringvector as SstArrayList).find(us)
    }

    /**
     * we need to override stream to update changes to the byte array
     */
    override fun preStream() {
        this.updateUnicodestrings()
    }

    // For debugging purposes
    override fun toString(): String {
        val sb = StringBuffer()
        sb.append("cstTotal:" + numTotal + " cstUnique:" + numUnique
                + " numConts:" + numContinues)
        for (i in stringvector.indices) {
            sb.append("\n " + stringvector[i])
        }
        return sb.toString()
    }

    /**
     * Override ArrayList to allow matching based on .toString. Required because
     * we call ArrayList.indexOf(String) when ArrayList contains UnicodeStrings.
     */
    private inner class SstArrayList : ArrayList<*>() {
        private val container = HashMap()

        fun put(o: Any, isst: Int?): Boolean {
            container.put((o as Unicodestring).toCachingString(), isst)
            return super.add(o)
        }

        override fun indexOf(o: Any): Int {
            val oo = container.get(o.toString()) ?: return -1
            return (oo as Int).toInt()
        }

        override fun remove(o: Any?): Boolean {
            Logger.logWarn("String being removed from SST array, Indexing may be off")
            container.remove((o as Unicodestring).toCachingString())
            return super.remove(o)
        }

        /**
         * find this particular unicode string, including formatting
         *
         * @param us
         * @return
         */
        fun find(us: Unicodestring): Int {
            return super.indexOf(us)
        }

        companion object {
            /**
             * serialVersionUID
             */
            private val serialVersionUID = 7904551471519095640L
        }
    }

    /**
     * generate the OOXML necessary to describe this string table, also fill
     * sststrings list with unique sststrings
     *
     * @param sststrings
     * @return sstooxml
     * @throws IOException
     */
    @Throws(IOException::class)
    fun writeOOXML(zip: Writer) {
        val sstooxml = StringBuffer()

        zip.write(OOXMLConstants.xmlHeader)
        zip.write("\r\n")
        zip.write("<sst xmlns=\"" + OOXMLConstants.xmlns + "\" count=\""
                + numTotal + "\" uniqueCount=\"" + numUnique + "\">")
        zip.write("\r\n")
        for (i in 0 until this.stringVector.size) {
            val us = this.stringVector[i] as Unicodestring
            val frs = us.formattingRuns
            var s = us.stringVal
            s = OOXMLAdapter.stripNonAscii(s).toString()
            // sststrings.add(OOXMLAdapter.stripNonAscii(s));// zip.write(s); //
            // used as an index for cell values in parsing sheet ooxml

            // TODO: below should be in Unicodestring as .getOOXML?
            zip.write("<si>")
            zip.write("\r\n")

            if (frs == null) { // no intra-string formattingz
                if (s.indexOf(" ") == 0 || s.lastIndexOf(" ") == s.length - 1) {
                    zip.write("<t xml:space=\"preserve\">$s</t>")
                } else {
                    zip.write("<t>$s</t>")
                }
                zip.write("\r\n")
            } else { // have formatting runs which split up string into areas
                // with separate formats applied
                var begIdx = 0
                for (j in frs.indices) {
                    val idxs = frs[j] as ShortArray
                    if (idxs[0] > begIdx) { // +1!!
                        if (j == 0) {
                            zip.write("<r>") // new rich text run
                            zip.write("<t xml:space=\"preserve\">"
                                    + OOXMLAdapter.stripNonAscii(s.substring(
                                    begIdx, idxs[0].toInt())) + "</t>")
                            zip.write("</r>")
                            zip.write("\r\n")
                        } else {
                            zip.write("<t xml:space=\"preserve\">"
                                    + OOXMLAdapter.stripNonAscii(s.substring(
                                    begIdx, idxs[0].toInt())) + "</t>")
                            zip.write("</r>")
                            zip.write("\r\n")
                        }
                        begIdx = idxs[0].toInt()
                    }
                    zip.write("<r>") // new rich text run
                    val rp = Ss_rPr.createFromFont(this.workBook!!
                            .getFont(idxs[1].toInt()))
                    zip.write(rp.ooxml)
                }
                if (begIdx < s.length)
                // output remaining string
                    s = s.substring(begIdx)
                else
                    s = ""
                zip.write("<t xml:space=\"preserve\">"
                        + OOXMLAdapter.stripNonAscii(s) + "</t>")
                zip.write("\r\n")
                zip.write("</r>")
            }
            zip.write("</si>")
            zip.write("\r\n")
        }
        zip.write("</sst>")
        // return sstooxml.toString();
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6966063306230877101L

        /**
         * Initializes the sst as well as initializing the UnicodeStrings contained
         * within
         */
        fun init(sst: Sst) {
            sst.origSstLen = sst.length
            sst.currbound = 0
            sst.stringvector.clear()

            // init the string cache for fast access of init vals

            // get the row, col and ixfe information
            sst.numTotal = ByteTools.readInt(sst.getByteAt(0), sst.getByteAt(1),
                    sst.getByteAt(2), sst.getByteAt(3))
            sst.numUnique = ByteTools.readInt(sst.getByteAt(4), sst.getByteAt(5),
                    sst.getByteAt(6), sst.getByteAt(7))
            var strlen = 0
            var strpos = 8

            if (sst.DEBUGLEVEL > 5)
                Logger.logInfo("INFO: initializing Sst: " + sst.numTotal
                        + " total Strings, " + sst.numUnique + " unique Strings.")
            // Initialize continues records
            sst.initContinues()

            // initialize the Unicodestrings from the byte array
            for (d in 0 until sst.numUnique) {
                // Unicodestring values
                var numruns = 0
                var runlen = 0
                // the number of formatting runs each one adds 4 bytes
                var basereclen = 3 // the base length of the ustring being created
                var cchExtRst = 0 // the length of any Extended string data
                var doubleByte = false // whether this is a double-byte string
                var grbit: Byte = 0x0
                // the grbit tells us what kind of Unicodestring this is
                if (sst.DEBUGLEVEL > 30)
                    Logger.logInfo("Initializing String: " + d
                            + "/" + sst.numTotal)
                // figure out the boundary offsets
                var offr = sst.boundaries!!.size
                if (offr < 1)
                    offr = 0
                else
                    offr = 1
                if (strpos >= sst.boundaries!![sst.boundaries!!.size - offr])
                    break
                val recdef = sst.getNextStringDefData(strpos)
                // get the length of the Unicodestring
                strlen = recdef[0].toInt()
                grbit = recdef[1].toByte()

                // we only want the bottom 4 bytes of the grbit, & not bit 2.. other
                // stuff is junk
                // grbit = (byte)(0xD & grbit);

                // init the string cache for fast access of init vals
                // st.initCacheBytes(strpos, 10); // commented out as it causes
                // array errors when short strings at end of continue boundary

                var currec: XLSRecord? = sst
                if (sst.DEBUGLEVEL > 5)
                    Logger.logInfo("INFO: StrLen:" + strlen + " Strpos:" + strpos + " bound:" + sst.boundaries!![sst.currbound])

                if (strpos >= sst.boundaries!![0])
                    currec = sst.thiscont
                when (grbit) {

                    0x1 // non-rich, double-byte string
                    -> doubleByte = true

                    0x4 // non-rich, single byte string
                    -> {
                        cchExtRst = ByteTools.readInt(currec!!.getByteAt(strpos + 3),
                                currec.getByteAt(strpos + 4),
                                currec.getByteAt(strpos + 5),
                                currec.getByteAt(strpos + 6))
                        basereclen = 7
                        doubleByte = false
                    }

                    0x5 // extended, non-rich, double-byte string
                    -> {
                        cchExtRst = ByteTools.readInt(currec!!.getByteAt(strpos + 3),
                                currec.getByteAt(strpos + 4),
                                currec.getByteAt(strpos + 5),
                                currec.getByteAt(strpos + 6))
                        basereclen = 7
                        doubleByte = true
                    }

                    0x8 // rich single-byte UNICODE string
                    -> {
                        numruns = ByteTools.readShort(currec!!.getByteAt(strpos + 3).toInt(),
                                currec.getByteAt(strpos + 4).toInt()).toInt()
                        runlen = numruns * 4
                        basereclen = 5
                        doubleByte = false
                    }

                    0x9 // rich double-byte UNICODE string
                    -> {
                        numruns = ByteTools.readShort(currec!!.getByteAt(strpos + 3).toInt(),
                                currec.getByteAt(strpos + 4).toInt()).toInt()
                        runlen = numruns * 4
                        basereclen = 5
                        doubleByte = true
                    }

                    0xc // rich single-byte eastern string
                    -> {
                        numruns = ByteTools.readShort(currec!!.getByteAt(strpos + 3).toInt(),
                                currec.getByteAt(strpos + 4).toInt()).toInt()
                        cchExtRst = ByteTools.readInt(currec.getByteAt(strpos + 5),
                                currec.getByteAt(strpos + 6),
                                currec.getByteAt(strpos + 7),
                                currec.getByteAt(strpos + 8))
                        runlen = numruns * 4
                        basereclen = 9
                        doubleByte = false
                    }

                    0xd // rich double-byte eastern string
                    -> {
                        numruns = ByteTools.readShort(currec!!.getByteAt(strpos + 3).toInt(),
                                currec.getByteAt(strpos + 4).toInt()).toInt()
                        cchExtRst = ByteTools.readInt(currec.getByteAt(strpos + 5),
                                currec.getByteAt(strpos + 6),
                                currec.getByteAt(strpos + 7),
                                currec.getByteAt(strpos + 8))
                        runlen = numruns * 4
                        basereclen = 9
                        doubleByte = true
                    }

                    else -> {
                        doubleByte = false
                        cchExtRst = 0
                        basereclen = 3
                        if (grbit.toInt() != 0x0) {
                            // if(st.DEBUGLEVEL > 10)
                            Logger.logWarn("ERROR: Invalid Unicodestring grbit:$grbit")
                        }
                    }
                }
                // create the String
                if (strlen == 0) {
                    if (sst.DEBUGLEVEL > 10)
                        Logger.logWarn("WARNING: Attempt to initialize Zero-length String.")
                }
                if (doubleByte)
                    strlen *= 2
                // it's a double-byte string so total size is *2
                try {
                    strpos = sst.initUnicodeString(strlen, strpos, basereclen, cchExtRst, runlen, doubleByte)

                    // Logger.logInfo("SST Currbound: " + sst.currbound +" strpos: "
                    // + strpos);
                    // if(st.DEBUGLEVEL > 5)Logger.logInfo("numruns: "
                    // +String.valueOf(numruns)+" @"+String.valueOf(strpos)
                    // +" len: " + String.valueOf(strlen) + " gr: "
                    /// +String.valueOf(grbit) +
                    // " base: "+String.valueOf(basereclen)+
                    // " cchExtRst: "+String.valueOf(cchExtRst));
                } catch (e: Exception) {
                    Logger.logWarn("ERROR: Error Reading String @ $strpos$e Skipping...")
                    strpos += strlen + basereclen + runlen
                }

            }
            if (sst.DEBUGLEVEL > 5)
                Logger.logInfo("Done reading SST.")
        }

        /**
         * return the sizes of Continue records for an Sst caches the read if
         * neccesary
         */
        fun getContinueDef(rec: Sst, cached: Boolean): Array<Any> {
            if (cached) {
                return rec.continueDef
            } else {
                val cbs = arrayOfNulls<Int>(rec.cbounds!!.size)
                val sstgrs = arrayOfNulls<Byte>(rec.sstgrbits!!.size)
                for (t in cbs.indices) {
                    cbs[t] = rec.cbounds!![t] as Int
                }

                for (t in sstgrs.indices) {
                    sstgrs[t] = rec.sstgrbits!![t] as Byte
                }

                rec.continueDef = arrayOfNulls<Any>(2)
                rec.continueDef[0] = cbs
                rec.continueDef[1] = sstgrs
                return rec.continueDef
            }

        }

        /**
         * Create a unicode string
         *
         * @param s
         * @param formattingRuns
         * @param ENCODINGMODE
         * @return
         */
        fun createUnicodeString(s: String,
                                formattingRuns: ArrayList<*>?, ENCODINGMODE: Int): Unicodestring? {
            try {
                var isuni = false
                if (ENCODINGMODE == WorkBook.STRING_ENCODING_AUTO)
                    isuni = ByteTools.isUnicode(s)
                else if (ENCODINGMODE == WorkBook.STRING_ENCODING_COMPRESSED)
                    isuni = false
                else if (ENCODINGMODE == WorkBook.STRING_ENCODING_UNICODE)
                    isuni = true
                if (formattingRuns != null)
                    isuni = true

                var charbytes = s.toByteArray(charset(XLSConstants.DEFAULTENCODING))
                var strlen = charbytes.size // .length();
                var strbytes: ByteArray? = null

                // handle string sizes
                if (strlen * 2 > java.lang.Short.MAX_VALUE)
                    isuni = false // can't fit larger than Short String length
                if (strlen > java.lang.Short.MAX_VALUE - 3) { // if strlen is greater than
                    // the maximum value for
                    // excel cells, truncate
                    strlen = java.lang.Short.MAX_VALUE - 3 // maximum value
                    charbytes = ByteArray(strlen)
                    System.arraycopy(s.toByteArray(charset(XLSConstants.DEFAULTENCODING)), 0,
                            charbytes, 0, strlen)
                }

                if (formattingRuns != null)
                    isuni = true

                if (isuni) { // encode string bytes
                    try {// if you use a string here for the encoding rather than a
                        // reference to a static String, performance in JDK 4.2
                        // will suffer. Why? Dunno, but it's bad!
                        charbytes = s.toByteArray(charset(WorkBook.UNICODEENCODING))
                    } catch (e: UnsupportedEncodingException) {
                        Logger.logWarn("error encoding string: " + e
                                + " with default encoding 'UnicodeLittleUnmarked'")
                    }

                    if (formattingRuns == null)
                        strbytes = ByteArray(charbytes.size + 3)
                    else
                        strbytes = ByteArray(charbytes.size + 5) // need 2 extra
                    // bytes to
                    // store
                    // formatting
                    // run info
                } else
                    strbytes = ByteArray(charbytes.size + 3)

                // given info, create strbytes for Unicode init
                var pos = 0
                val encodedlen = charbytes.size
                val lenbytes = ByteTools.shortToLEBytes(strlen.toShort())
                strbytes[pos++] = lenbytes[0] // cch bytes 0 & 1
                strbytes[pos++] = lenbytes[1]
                if (!isuni) {
                    strbytes[pos++] = 0x0.toByte() // grbit byte 2
                } else {
                    strbytes[pos++] = 0x1.toByte() // grbit byte 2
                    if (formattingRuns != null) { //
                        strbytes[pos - 1] = strbytes[pos - 1] or 0x8 // set Rich Text attribute
                        val fr = ByteTools.shortToLEBytes(formattingRuns
                                .size.toShort())
                        strbytes[pos++] = fr[0] // # formatting runs bytes 3 & 4
                        strbytes[pos++] = fr[1]
                    }
                }
                System.arraycopy(charbytes, 0, strbytes, pos, encodedlen)

                if (formattingRuns != null) {
                    // formatting runs (charindex, fontindex)*n after string data
                    val frs = ByteArray(formattingRuns.size * 4)
                    for (i in formattingRuns.indices) {
                        val o = formattingRuns[i] as ShortArray
                        val charIndex = ByteTools.shortToLEBytes(o[0])
                        val fontIndex = ByteTools.shortToLEBytes(o[1])
                        System.arraycopy(charIndex, 0, frs, i * 4, 2)
                        System.arraycopy(fontIndex, 0, frs, i * 4 + 2, 2)
                    }
                    // Append frs to end of strbytes
                    val newdata = ByteArray(strbytes.size + frs.size)
                    System.arraycopy(strbytes, 0, newdata, 0, strbytes.size)
                    System.arraycopy(frs, 0, newdata, strbytes.size, frs.size)
                    strbytes = newdata
                }
                // create a new one, set its data
                val str = Unicodestring()
                str.init(strbytes, false)
                return str
            } catch (e: UnsupportedEncodingException) {
                Logger.logWarn("error encoding string: $e")
            }

            return null
        }

        /**
         * given SharedStrings.xml OOXML inputstream, read in string and formatting
         * data, if any and parse into ArrayList for later use in parseSheetOOXML
         *
         * @param bk WorkBookHandle
         * @param ii InputStream
         * @return String ArrayList return list of shared strings
         * @see parseSheetOOXML
         */
        fun parseOOXML(bk: WorkBookHandle, ii: InputStream): ArrayList<*> {
            // NOTE:
            // apparently can have dup entries in sharedstring.xml
            // index of string links to cell value so must keep dups here
            // reset after parsing
            var shareDups = false
            if (bk.workBook!!.isSharedupes) {
                bk.workBook!!.isSharedupes = false
                shareDups = true
            }
            try {
                val factory = XmlPullParserFactory.newInstance()
                factory.isNamespaceAware = true
                val xpp = factory.newPullParser()

                xpp.setInput(ii, "UTF-8") // using XML 1.0 specification
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "si") { // parse si single string table
                            // entry
                            var s = ""
                            var formattingRuns: ArrayList<*>? = null
                            while (eventType != XmlPullParser.END_DOCUMENT) {
                                if (eventType == XmlPullParser.START_TAG) {
                                    if (xpp.name == "rPr") { // intra-string
                                        // formatting
                                        // properties
                                        val idx = s.length // index into
                                        // character string
                                        // to apply
                                        // formatting to
                                        val rp = Ss_rPr.parseOOXML(xpp, bk)
                                                .cloneElement() as Ss_rPr
                                        val f = rp.generateFont(bk) // NOW CONVERT
                                        // ss_rPr to
                                        // a font!!
                                        var fIndex = bk.workBook!!.getFontIdx(f) // index
                                        // for
                                        // specific
                                        // font
                                        // formatting
                                        if (fIndex == -1)
                                        // must insert new font
                                            fIndex = bk.workBook!!.insertFont(f) + 1
                                        if (formattingRuns == null)
                                            formattingRuns = ArrayList()
                                        formattingRuns!!.add(shortArrayOf(Integer.valueOf(idx)!!.toShort(), Integer.valueOf(fIndex)!!.toShort()))
                                    } else if (xpp.name == "t") {
                                        /*
                                     * boolean bPreserve= false; if
                                     * (xpp.getAttributeCount()>0) { if
                                     * (xpp.getAttributeName(0).equals("space")
                                     * &&
                                     * xpp.getAttributeValue(0).equals("preserve"
                                     * )) bPreserve= true; }
                                     */
                                        eventType = xpp.next()
                                        while (eventType != XmlPullParser.END_DOCUMENT
                                                && eventType != XmlPullParser.END_TAG
                                                && eventType != XmlPullParser.TEXT) {
                                            eventType = xpp.next()
                                        }
                                        if (eventType == XmlPullParser.TEXT) {
                                            s += xpp.text
                                        }
                                    }
                                } else if (eventType == XmlPullParser.END_TAG && xpp.name == "si") {
                                    bk.workBook!!.sharedStringTable!!
                                            .addUnicodestring(s, formattingRuns) // create
                                    // a
                                    // new
                                    // unicode
                                    // string
                                    // with
                                    // formatting
                                    // runs
                                    break
                                }
                                eventType = xpp.next()
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("SST.parseXML: $e")
            }

            if (shareDups)
                bk.workBook!!.isSharedupes = true

            return bk.workBook!!.sharedStringTable!!
                    .stringVector as ArrayList<*>
        }
    }
}
