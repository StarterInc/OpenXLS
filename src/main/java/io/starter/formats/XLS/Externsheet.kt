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

import io.starter.formats.XLS.formulas.IxtiListener
import io.starter.formats.XLS.formulas.Ptg
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

/**
 * **Externsheet: External Sheet Record (17h)**<br></br>
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       cXTI        2       Number of XTI Structures
 * 6       rgXTI       var     Array of XTI Structures
 *
 * XTI
 * offset  name        size    contents
 * ---
 * 0       iSUPBOOK    2       0-based index to table of SUPBOOK records
 * 2       itabFirst   2       0-based index to first sheet tab in reference
 * 4       itabLast    2       0-based index to last sheet tab in reference
 *
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Supbook
 */

class Externsheet : io.starter.formats.XLS.XLSRecord() {
    internal var cXTI: Short = 0
    // int DEBUGLEVEL = 10;
    internal var rgs = CompatibleVector()

    override var workBook: WorkBook?
        get
        set(bk) {
            super.workBook = bk
            val it = rgs.iterator()
            while (it.hasNext())
                (it.next() as rgxti).setWorkBook(bk)
        }

    /**
     * Certain records require a virtual reference, this is not a real reference to a sheet,
     * rather an entry that is used by add in formulas, values are FE FF FE FF
     *
     *
     * This method either finds the existing reference, or creates a new one and returns
     * the pointer
     *
     * @return
     */
    /*null*//*null*/ val virtualReference: Int
        get() {
            for (i in rgs.indices) {
                val thisXti = rgs.elementAt(i) as rgxti
                if (thisXti.sheet1num == 0xFFFE.toShort() && thisXti.sheet2num == 0xFFFE.toShort())
                    return i
            }
            val bts = ByteArray(6)
            if (this.wkbook != null) {
                val sb = this.wkbook!!.supBooks
                bts[0] = getAddInIndex(sb)
            }
            val newcXTI = rgxti(bts)
            newcXTI.setWorkBook(this.wkbook)
            newcXTI.sheet1 = 0xFFFE
            newcXTI.sheet2 = 0xFFFE
            rgs.add(newcXTI)
            this.cXTI++
            this.update()
            return rgs.size - 1
        }

    /**
     * In some cases, we need to have an Xti reference to a non-existing sheet.  For instance, a chart
     * or formula with a ptgRef3d that referrs to a missing sheet.  Externsheet handles this by having
     * an internal record populated with -1's   This method searches for a non-existing record, and if
     * that doesn't exist it creates on and passes the reference back.
     *
     * @return pointer to broken (-1) reference
     */
    /*getSheet1Num()*//*getSheet2Num()*/ val brokenXtiReference: Int
        get() {
            for (i in rgs.indices) {
                val thisXti = rgs.elementAt(i) as rgxti
                if (thisXti.sheet1num.toInt() == 0xFFFF && thisXti.sheet2num.toInt() == 0xFFFF)
                    return i
            }
            val bts = ByteArray(6)
            if (this.wkbook != null) {
                val sb = this.wkbook!!.supBooks
                bts[0] = getGlobalSupBookIndex(sb)
            }
            val newcXTI = rgxti(bts)
            newcXTI.setWorkBook(this.wkbook)
            newcXTI.sheet1 = 0xFFFF
            newcXTI.sheet2 = 0xFFFF
            rgs.add(newcXTI)
            this.cXTI++
            this.update()
            return rgs.size - 1
        }

    override fun preStream() {
        this.update()
    }

    @Throws(WorkSheetNotFoundException::class)
    fun addPtgListener(p: IxtiListener) {
        val ix = p.ixti
        // if(ix>0)ix--;
        if (this.rgs.size > ix) {
            val rg = this.rgs[ix.toInt()] as rgxti
            rg.addListener(p)
        } else {
            val rg = rgxti()
            rg.setWorkBook(this.wkbook)    // 20080306 KSC
            rg.sheet1 = ix   // 20080306 KSC: use actual sheet# this.wkbook.getWorkSheetByNumber(ix));
            rg.sheet2 = ix   // ""  this.wkbook.getWorkSheetByNumber(ix));
            this.rgs.add(rg)
            rg.addListener(p)
            this.update()
        }
    }


    fun getBoundSheets(cLoc: Int): Array<Boundsheet>? {
        var cLoc = cLoc

        if (rgs.size == 0)
            return null

        if (cLoc > rgs.size - 1)
            cLoc = rgs.size - 1

        val rg = rgs[cLoc] as rgxti

        val first = rg.sheet1num
        var last = rg.sheet2num
        if (first == 0xFFFE.toShort())
        // associated with a Name record
            return null
        if (first == 0xFFFF.toShort())
        // 20080212 KSC - should be a ref to a deleted or unfound sheet
            return null    // error trap trying to get virtual sheet

        var numshts = ++last - first
        if (numshts < 1) numshts = 1
        val bs = arrayOfNulls<Boundsheet>(numshts)
        var p = 0
        for (t in first until last) {
            try {
                bs[p++] = this.wkbook!!.getWorkSheetByNumber(t)
            } catch (e: WorkSheetNotFoundException) {
                // don't error out on the external workbook sheet references
                if (DEBUGLEVEL > 1 && t != 65535 && !rg.isExternal)
                // 20080306 KSC: add external check ...
                    Logger.logWarn("Attempt to access Externsheet reference for sheet failed: $e")
            }

        }
        return bs
    }

    /**
     * returns array of referenced sheet names, including external references ...
     *
     * @param cLoc
     * @return
     */
    fun getBoundSheetNames(cLoc: Int): Array<String>? {
        var cLoc = cLoc
        if (rgs.size == 0)
            return null
        if (cLoc > rgs.size - 1)
            cLoc = rgs.size - 1

        val rg = rgs[cLoc] as rgxti

        var first = rg.sheet1num
        var last = rg.sheet2num

        if (first == 0xFFFE.toShort())
        // associated with a Name record (=Add-in)
            return arrayOf("AddIn")

        if (first == 0xFFFF.toShort())
        // is a ref to a deleted or un-found sheet
            return arrayOf("#REF!")

        var numshts = ++last - first
        if (numshts < 1) numshts = 1
        if (first < 0)
            first = 1

        val sheets = arrayOfNulls<String>(numshts)
        if (first == last)
            return arrayOf("#REF!")
        var p = 0
        for (t in first until last) {
            try {
                sheets[p++] = rg.getSheetName(t)   // should successfully retrieve External Sheetnames
            } catch (we: WorkSheetNotFoundException) {
                if (DEBUGLEVEL > 1)
                    Logger.logWarn("Attempt to access Externsheet reference for sheet failed: $we")
            }

        }
        return sheets
    }

    /**
     * returns true if the passed in sheet number is an
     * external link (i.e. an external sheet reference)
     *
     * @param loc external sheet number
     * @return
     */
    fun getIsExternalLink(loc: Int): Boolean {
        if (rgs.size == 0)
            return false

        val rg = rgs[loc] as rgxti
        return rg.isExternal
    }

    /**
     * get the number of refs in this Externsheet rec
     */
    fun getcXTI(): Int {
        return cXTI.toInt()
    }

    override fun init() {
        super.init()
        cXTI = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        var pos = 2
        for (t in 0 until cXTI) {
            try {
                val bts = this.getBytesAt(pos, 6) // System.arraycopy(rkdata,pos,bts,0,6);
                val rg = rgxti(bts)
                rgs.add(rg)
                if (wkbook != null)
                    rg.setWorkBook(wkbook)
            } catch (e: Exception) {
                if (DEBUGLEVEL > 10) Logger.logWarn("init of Externsheet record failed: $e")
            }

            pos += 6
        }
        if (DEBUGLEVEL > 10) Logger.logInfo("Done Creating Externsheet")
    }

    /**
     * update the underlying bytes by recreating from the properties
     */
    internal fun update() {
        // init a new byte array
        var blen = rgs.size * 6
        blen += 2
        val newbytes = ByteArray(blen)
        this.cXTI = (rgs.size - 1).toShort()
        // get the number of structs
        val cx = ByteTools.shortToLEBytes(rgs.size.toShort())
        System.arraycopy(cx, 0, newbytes, 0, 2)

        // iterate the structs and get the bytes
        var pos = 2
        val it = rgs.iterator()
        while (it.hasNext()) {
            val btx = (it.next() as rgxti).bytes
            System.arraycopy(btx, 0, newbytes, pos, 6)
            pos += 6
        }

        // set the data
        this.setData(newbytes)
    }

    /**
     * Remove a sheet from this reference table,
     *
     *
     * Also updates the ixti listeners, which appear to be records that contain an internal
     * reference to a sheet number.
     */
    @Throws(WorkSheetNotFoundException::class)
    internal fun removeSheet(sheetnum: Int) {
        if (DEBUGLEVEL > 10) Logger.logInfo("Removing Sheet from Externsheet")
        // iterate the cXTI and check if this sheet is contained
        val it = rgs.iterator()
        while (it.hasNext()) {
            val rg = it.next() as rgxti
            val first = rg.sheet1num.toInt()
            val last = rg.sheet2num.toInt()
            if (sheetnum == first && sheetnum == last) {
                // reference should not be removed, rather set to -1;
                rg.sheet1 = 0xFFFF
                rg.sheet2 = 0xFFFF
                this.cXTI--
            } else if (sheetnum >= first && sheetnum <= last) { // it's contained in the ref
                rg.sheet2 = rg.sheet2 - 1
            } else if (sheetnum <= first) {
                rg.sheet1 = rg.sheet1 - 1
                rg.sheet2 = rg.sheet2 - 1
            }
        }
        this.update()
    }

    /* 20100506 KSC: this is not necessary any longer since
     * rg entries are not removed upon deletes or moved in any operations
     * therefore the ixti will stay constant
     * -- for deleted sheets, the sheet reference will be set to the 0xFFFE deleted
     * sheet reference
    public void notifyIxtiListeners() {
        Iterator it = rgs.iterator();
        while(it.hasNext()) {
            ((rgxti)it.next()).notifyListeners();
        }
    }
    */

    /**
     * add a sheet to this reference table
     *
     *
     * why doesn't this return the new referenced int?  NR
     */
    @Throws(WorkSheetNotFoundException::class)
    internal fun addSheet(sheetnum: Int) {
        this.addSheet(sheetnum, sheetnum)
    }

    private fun getAddInIndex(sb: Array<Supbook>): Byte {
        var i = 0
        while (i < sb.size) {
            if (sb[i].isAddInRecord)
                return i.toByte()
            i++
        }
        return (-1).toByte()
    }

    private fun getGlobalSupBookIndex(sb: Array<Supbook>): Byte {
        var i = 0
        while (i < sb.size) {
            if (sb[i].isGlobalRecord)
                return i.toByte()
            i++
        }
        return (-1).toByte()
    }

    /**
     * add a sheet range this reference table
     */
    @Throws(WorkSheetNotFoundException::class)
    internal fun addSheet(firstSheet: Int, lastSheet: Int) {
        if (DEBUGLEVEL > 10) Logger.logInfo("Adding new Sheet to Externsheet")

        // KSC: Added logic to set correct supbook index for added XTI
        val bts = ByteArray(6)
        if (this.wkbook != null) {   // should never happen!
            val sb = this.wkbook!!.supBooks    // must have SUPBOOK records when have an EXTERNSHEET!
            if (firstSheet == 0xFFFE)
            // then link to ADD-IN SUPBOOK
                bts[0] = getAddInIndex(sb)
            else {// link to global SUPBOOK record
                bts[0] = getGlobalSupBookIndex(sb)
            }
        }
        val newcXTI = rgxti(bts)
        newcXTI.setWorkBook(this.wkbook)      // 20080306 KSC
        if (firstSheet == 0xFFFE)
        // it's a virtual sheet range for Add-ins
            newcXTI.isAddIn = true

        if (!newcXTI.isAddIn) {
            newcXTI.sheet1 = firstSheet
            newcXTI.sheet2 = lastSheet
        }
        rgs.add(newcXTI)

        this.cXTI++
        this.update()
    }

    fun addExternalSheetRef(externalWorkbook: String, externalSheetName: String): Short {
        // get the external supbook record for this external workbook, creates if not present
        val sb = this.wkbook!!.getExternalSupbook(externalWorkbook, true)
        val sheetRef = sb!!.addExternalSheetReference(externalSheetName)
        val sbRef = this.wkbook!!.getSupbookIndex(sb).toShort()

        /* see if external ref  exists already */
        val it = rgs.iterator()
        var i = 0
        while (it.hasNext()) {
            val rg = it.next() as rgxti
            if (rg.sheet1num == sheetRef && rg.sheet2num == sheetRef && rg.sbs == sbRef)
                return i.toShort()
            i++
        }

        val bts = ByteArray(6)
        System.arraycopy(ByteTools.shortToLEBytes(sbRef), 0, bts, 0, 2)   // input SUPBOOK ref #
        System.arraycopy(ByteTools.shortToLEBytes(sheetRef), 0, bts, 2, 2)    // input Sheet ref #
        System.arraycopy(ByteTools.shortToLEBytes(sheetRef), 0, bts, 4, 2)    // input Sheet ref #
        val newcXTI = rgxti(bts)
        newcXTI.setIsExternalRef(true)        // flag don't look up worksheets in this workbook
        newcXTI.setWorkBook(this.wkbook)      // 20080306 KSC
        rgs.add(newcXTI)

        this.cXTI++
        this.update()
        return this.cXTI
    }

    /**
     * Insert location checks if a specific boundsheet range already has a reference.
     *
     *
     * If the range already exists within the externsheet the index to the
     * range is returned.  Else, it adds the range to the externsheet and returns the index.
     *
     * @param firstBound
     * @param lastBound
     * @return
     */
    @Throws(WorkSheetNotFoundException::class)
    fun insertLocation(firstBound: Int, lastBound: Int): Int {
        val it = rgs.iterator()
        var i = 0
        while (it.hasNext()) {
            val rg = it.next() as rgxti
            val first = rg.sheet1num.toInt()   // 20080306 KSC: getSheet1Num();
            val last = rg.sheet2num.toInt()    // getSheet2Num();

            if (first == firstBound && last == lastBound && rg.sb!!.isGlobalRecord) {
                return i
            }
            i++
        }
        this.addSheet(firstBound, lastBound)
        return rgs.size - 1
    }

    /**
     * Gets the xti reference for a boundsheet name
     *
     * @return xti reference, -1 if not located.
     */
    fun getXtiReference(firstSheet: String, secondSheet: String): Int {
        for (i in rgs.indices) {
            val thisXti = rgs.elementAt(i) as rgxti
            try {
                if (thisXti.getSheetName(thisXti.sheet1num.toInt()).equals(firstSheet, ignoreCase = true) && thisXti.getSheetName(thisXti.sheet2num.toInt()).equals(secondSheet, ignoreCase = true)) {
                    return i
                }
            } catch (we: WorkSheetNotFoundException) {
                if (DEBUGLEVEL > 10)
                    Logger.logWarn("Externsheet.getXtiReference:  Attempt to find Externsheet reference for sheet failed: $we")
            }

        }
        return -1
    }

    override fun close() {
        while (rgs.size > 0)
            rgs.removeAt(0)
    }

    /**
     * Internal structure tracks Sheet references
     */
    internal inner class rgxti : java.io.Serializable {
        var sbs: Short = 0              // supbook #
        var sheet1num: Short = 0
        var sheet2num: Short = 0 // store numbers as sometimes sheets are not in workbook ...
        var sb: Supbook? = null
        var wkbook: WorkBook? = null
        var bts: ByteArray? = null
        var isAddIn = false   // flag if this is "iSupbook" Externsheet
        /**
         * return if this rgi references an external SUPBOOK record
         * i.e. one that references an External Workbook
         * (therefore sheet references are not found in the current workbook)
         */
        var isExternal = false
            private set // flag if this is an external ref.
        // contains references to virtual 0xFFFE sheet #

        var listeners = CompatibleVector()

        /**
         * @return Returns the bts.
         */
        val bytes: ByteArray
            get() {
                if (bts == null) bts = ByteArray(6)
                var shtbt = ByteTools.shortToLEBytes(sheet1num)
                System.arraycopy(shtbt, 0, bts!!, 2, 2)
                shtbt = ByteTools.shortToLEBytes(sheet2num)
                System.arraycopy(shtbt, 0, bts!!, 4, 2)
                return bts
            }

        /**
         * set the sheet# referenced
         *
         * @param int sheet #
         */
        var sheet1: Int
            get() = this.sheet1num.toInt()
            set(sh1) {
                this.sheet1num = sh1.toShort()
            }

        /**
         * set the sheet# referenced
         *
         * @param int sheet #
         */
        var sheet2: Int
            get() = this.sheet2num.toInt()
            set(sh2) {
                this.sheet2num = sh2.toShort()
            }

        fun addListener(p: IxtiListener) {
            listeners.add(p)
        }

        fun notifyListeners() {
            val it = listeners.iterator()
            val locp = rgs.indexOf(this)
            while (it.hasNext()) {
                val p = it.next() as IxtiListener
                if (locp == -1)
                    (p as Ptg).parentRec.remove(true) // this reference has been deleted, rec not valid
                else
                    p.ixti = locp.toShort()
            }
        }

        /**
         * default constructor
         */
        constructor() {
            // do something?
        }

        fun setWorkBook(bk: WorkBook?) {
            this.wkbook = bk
            sbs = ByteTools.readShort(bts!![0].toInt(), bts!![1].toInt())           // supbook #
            sb = wkbook!!.supBooks[sbs]                      // get supbook referenced by sbs

            sheet1num = ByteTools.readShort(bts!![2].toInt(), bts!![3].toInt())
            sheet2num = ByteTools.readShort(bts!![4].toInt(), bts!![5].toInt())

            if (sheet1num.toInt() == 0xFFFE)
                isAddIn = true

            if (sb!!.isExternalRecord)
                isExternal = true

        }

        /**
         * constructor used to init from bytes without a book
         *
         * @param initbytes
         */
        constructor(initbytes: ByteArray) {
            this.bts = initbytes
        }

        /**
         * return the sheet name for the sheetNum in the associated supbook
         */
        @Throws(WorkSheetNotFoundException::class)
        fun getSheetName(sheetNum: Int): String {
            if (isAddIn)
            // no sheets assoc; return virtual sheet #
                return "Virtual Sheet Range 0xFFFE - 0xFFFE"
            if (isExternal)
                return sb!!.getExternalSheetName(sheetNum)
            return if (sheetNum == 0xFFFF) "Deleted Sheet" else wkbook!!.getWorkSheetByNumber(sheetNum).sheetName
            //otherwise, try and get sheetname

            /*
            if(this.sheet1==null)   // if sheet ref is deleted, return 0xFFFF
                return "null";  //-1;
            return this.sheet1.getSheetName();
            */
        }

        /**
         * set if this rgi references an external SUPBOOK record
         * i.e. one that references an External Workbook
         * (therefore sheet references are not found in the current workbook)
         *
         * @param bIsExternal
         */
        fun setIsExternalRef(bIsExternal: Boolean) {
            this.isExternal = bIsExternal
        }

        override fun toString(): String {

            try {
                return "rgxti range: " + this.getSheetName(sheet1num.toInt()) + "-" + this.getSheetName(sheet2num.toInt())
            } catch (we: WorkSheetNotFoundException) {
            }

            return "rgxti range: sheets not initialized"
        }

        companion object {
            /**
             * serialVersionUID
             */
            private const val serialVersionUID = -1591367957030959727L
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -4460757130836967839L

        /**
         * Constructor
         */
        protected// put a reference to sheet1 in as initial value
        val prototype: XLSRecord?
            get() {
                val x = Externsheet()
                x.length = 8.toShort()
                x.opcode = XLSConstants.EXTERNSHEET
                val dta = ByteArray(8)
                dta[0] = 0x1.toByte()
                x.setData(dta)
                x.originalsize = 8
                x.init()
                return x
            }

        /**
         * Add new Externsheet record and set sheet
         *
         * @param sheetNum1
         * @param sheetNum2
         * @param bk
         * @return
         */
        fun getPrototype(sheetNum1: Int, sheetNum2: Int, bk: WorkBook): XLSRecord? {
            val x = prototype as Externsheet?
            try {
                x!!.cXTI--
                x.rgs.removeAt(0)
                x.workBook = bk      // must, for addSheet
                //            x.addSheet(sheetNum1, sheetNum2);
            } catch (e: Exception) {
                Logger.logWarn("ExternSheet.getPrototype error:$e")
            }

            return x
        }
    }
}