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

import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.nio.charset.StandardCharsets
import java.util.Arrays


/**
 * **Supbook: Supporting Workbook (1AEh)**<br></br>
 *
 *
 * Supbook records store information about a supporting
 * external workbook
 *
 *
 * <pre>
 * for ADD-IN SUPBOOK record (occurs before Global SUPBOOK record)
 * offset  name        size    contents
 * ---
 * 0					2		01 00 (0001H)
 * 2					2		01 58 (0x3A) = Add-In supbook record
 *
 * for global SUPBOOK record (SUPBOOK 0) = Internal 3d References
 * offset  name        size    contents
 * ---
 * 0		nSheets		2		number of sheets in the workbook
 * 2					2		01 04 = global or own SUPBOOK record
 *
 *
 * SUPBOOK records for External References:
 * offset  name        size    contents
 * ---
 * 4       ctab        2       number of tabs in the workbook
 * 6       StVirtPath  var     Encoded file name (unicode)
 * var     Rgst        var     An array of tab sheet names (unicode)
 *
 * SUPBOOK records for OLE Objects/DDE
 * offset  name        size    contents
 * ---
 * 0					2		0000
 * 2					var		Encoded document name
 *
 *
</pre> *
 *
 * @see Boundsheet
 *
 * @see Externsheet
 */

/*
      * Encoded File URLS:
      * 1st char determines type of encoding:
      * 0x1 = Encoded URL follows
      * 0x2 = Reference to a sheet in the own document (sheetname follows)
      * 01H An MS-DOS drive letter will follow, or @ and the server name of a UNC path
  02H Start path name on same drive as own document
  03H End of subdirectory name
  04H Start path name in parent directory of own document (may occur repeatedly)
  05H Unencoded URL. Followed by the length of the URL (1 byte), and the URL itself.
  06H Start path name in installation directory of Excel
  08H Macro template directory in installation directory of Excel
  examples:
  =[ext.xls]Sheet1!A1           <01H>[ext.xls]Sheet1
  ='sub\[ext.xls]Sheet1'!A1         <01H>sub<03H>[ext.xls]Sheet1
  ='\[ext.xls]Sheet1'!A1            <01H><02H>[ext.xls]Sheet1
  ='\sub\[ext.xls]Sheet1'!A1        <01H><02H>sub<03H>[ext.xls]Sheet1
  ='\sub\sub2\[ext.xls]Sheet1'!A1 <01H><02H>sub<03H>sub2<03H>[ext.xls]Sheet1
  ='D:\sub\[ext.xls]Sheet1'!A1  <01H><01H>Dsub<03H>[ext.xls]Sheet1
  ='..\sub\[ext.xls]Sheet1'!A1  <01H><04H>sub<03H>[ext.xls]Sheet1
  ='\\pc\sub\[ext.xls]Sheet1'!A1    <01H><01H>@pc<03H>sub<03H>[ext.xls]Sheet1
  ='http://www.example.org/[ext.xls]Sheet1'!A1   <01H><05H><26H>http://www.example.org/[ext.xls]Sheet1
  (the length of the URL (38 = 26H) follows the 05H byte)
      */

class Supbook : io.starter.formats.XLS.XLSRecord() {
    internal var cstab = -1
    internal var nSheets = -1
    internal var tabs = CompatibleVector()
    private var filename: String? = null    // for EXTERNAL references

    /**
     * is this an Add-in Supbook record
     */
    val isAddInRecord: Boolean
        get() {
            val supBookCode = ByteArray(4)
            System.arraycopy(this.getData()!!, 0, supBookCode, 0, 4)
            return Arrays.equals(supBookCode, AddInProto)

        }

    /**
     * is this the global or default supbook record,
     * delineating the total number of sheets in this workbook
     *
     * @return
     */
    val isGlobalRecord: Boolean
        get() {
            val supBookCode = ByteArray(4)
            System.arraycopy(this.getData()!!, 2, supBookCode, 2, 2)
            return Arrays.equals(supBookCode, protocXTI)
        }

    /**
     * returns if this supbook record is the external supbook record,
     * linking sheets in an external workbook
     *
     * @return
     */
    val isExternalRecord: Boolean
        get() {
            val supBookCode = ByteArray(4)
            System.arraycopy(this.getData()!!, 2, supBookCode, 2, 2)
            return !Arrays.equals(supBookCode, protocXTI) && !Arrays.equals(supBookCode, AddInProto)
        }

    /**
     * return the External WorkBook represented by this External Supbook Record
     * or null if it's not an external supbook
     *
     * @return
     */
    val externalWorkBook: String?
        get() = if (!isExternalRecord) null else filename

    override fun init() {
        super.init()
        this.getData()
        // KSC: Interpret SUPBOOK code
        if (isGlobalRecord)
            nSheets = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        else if (!isAddInRecord) {    // then it's an External Reference
            // 20080122 KSC: Ressurrect + get code to work
            try {
                cstab = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()    // number of tabs/sheetnames referenced
                var ln = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()    // len of bytes of filename "encoded URL without sheetname"
                var compression = this.getByteAt(4).toInt()    // TODO: ??? 0= compressed 8 bit 1= uncompressed 16 bit ... asian, rtf ...
                val encoding = this.getByteAt(5).toInt()    // MUST = 0x1, if not it's an internal ref
                var pos = 6

                if (compression == 0) { //non-unicode
                    /* this is whacked.  invalid code in several use cases, its trying to get the length of the string
                     * from bytes that are internal in the string itself;
		        	while (this.getByteAt(pos)<0x8) {	// get control chars ...
		        		if (this.getByteAt(pos)==0x5) { // unencoded, file length follows
		        			ln= this.getByteAt(++pos)+2;
		        		}
		        		pos++;
		        		ln--;
		        	}**/
                    val f = this.getBytesAt(pos, ln - 1)
                    filename = String(f!!)
                    pos += ln - 1
                } else {    // unicode
                    val f = this.getBytesAt(pos, ln * 2 - 1)
                    filename = String(f!!, StandardCharsets.UTF_16LE)
                    pos += ln * 2 - 1
                }
                if (DEBUGLEVEL > 5)
                    Logger.logInfo("Supbook File: " + filename!!)
                // now get sheetnames
                for (i in 0 until cstab) {
                    ln = ByteTools.readShort(this.getByteAt(pos).toInt(), this.getByteAt(pos + 1).toInt()).toInt()
                    compression = this.getByteAt(pos + 2).toInt()    // TODO: ??? 0= compressed 8 bit 1= uncompressed 16 bit ... asian, rtf ...
                    pos += 3
                    var sheetname = ""
                    if (compression == 0) {
                        val f = this.getBytesAt(pos, ln)
                        sheetname = String(f!!)
                        pos += ln
                    } else {
                        val f = this.getBytesAt(pos, ln * 2)
                        sheetname = String(f!!, StandardCharsets.UTF_16LE)
                        pos += ln * 2
                    }
                    tabs.add(sheetname)
                    if (DEBUGLEVEL > 5)
                        Logger.logInfo("Supbook Sheet Reference: $sheetname")
                }
            } catch (e: Exception) {
                Logger.logErr("Supbook.init External Record: $e")
            }

        }
    }


    /**
     * add an external sheet reference to this External Supbook record
     *
     * @param sheetName
     */
    fun addExternalSheetReference(externalSheet: String): Short {
        // see if sheet exists here already
        for (i in 0 until cstab) {
            if ((tabs[i] as String).equals(externalSheet, ignoreCase = true))
                return i.toShort()
        }
        cstab++    // increment # sheets
        System.arraycopy(ByteTools.shortToLEBytes(cstab.toShort()), 0, this.getData()!!, 0, 2)
        // Add new sheet reference to this SUPBOOK
        val f = externalSheet.toByteArray()    // get bytes of new sheet to add
        val ln = f.size/*+1*/                    // and the length
        val encoding = 0                        // default = no unicode // TODO: is this correct?
        val pos = this.getData()!!.size        // start at the end of the sheet refs
        val newData = ByteArray(pos + ln + 3)
        System.arraycopy(this.getData()!!, 0, newData, 0, pos)
        System.arraycopy(ByteTools.shortToLEBytes(ln.toShort()), 0, newData, pos, 2)
        newData[pos + 2] = encoding.toByte()
        System.arraycopy(f, 0, newData, pos + 3, ln/*-1*/)
        this.setData(newData)
        // add newest
        tabs.add(externalSheet)
        if (DEBUGLEVEL > 5)
            Logger.logInfo("Supbook Sheet Reference: $externalSheet")
        return (cstab - 1).toShort()
    }

    /**
     * FOR EXTERNAL SUPBOOKS, returns the i-th sheetname
     *
     * @param i
     * @return
     */
    fun getExternalSheetName(i: Int): String {
        return if (i < 0 || i > tabs.size - 1) "#REF" else tabs[i] as String
    }

    /**
     * If the record is an addinrecord, update the number of sheets referred to.
     *
     * @see io.starter.formats.XLS.XLSRecord.preStream
     */
    override fun preStream() {
        if (this.isGlobalRecord) {
            val b = this.workBook
            if (b != null) {
                nSheets = b.numWorkSheets
                val bite = ByteTools.shortToLEBytes(nSheets.toShort())
                val rkdata = this.getData()
                rkdata[0] = bite[0]
                rkdata[1] = bite[1]
                this.setData(rkdata)
            }

        }
    }

    /**
     * return String representation
     */
    override fun toString(): String {
        if (nSheets != -1)
            return "SUPBOOK: Number of Sheets: $nSheets"
        var ret = ""
        if (filename != null) {    // assume it's a valid external ref
            ret = "SUPBOOK: External Ref: $filename Sheets:"
            for (i in 0 until cstab) {
                ret += " " + tabs[i]
            }
        } else
        // assume it's an add-in
            return "SUPBOOK: ADD-IN"
        return ret
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5010774855281594364L

        internal var protocXTI = byteArrayOf(0x0, 0x0, 0x1, 0x4)
        // KSC: different byte sequence for Add-ins:
        internal var AddInProto = byteArrayOf(0x1, 0x0, 0x1, 0x3A)


        internal var defaultsize = 4

        /**
         * Constructor
         */
        fun getPrototype(numtabs: Int): XLSRecord {
            val x = Supbook()
            x.length = defaultsize.toShort()
            x.opcode = XLSConstants.SUPBOOK
            // x.setLabel("SUPBOOK");
            val dta = ByteArray(defaultsize)
            System.arraycopy(protocXTI, 0, dta, 0, protocXTI.size)
            dta[0] = numtabs.toByte()
            x.setData(dta)
            x.originalsize = defaultsize
            x.init()
            return x
        }

        /**
         * SupBook record for add-in is different
         */
        val addInPrototype: XLSRecord
            get() {
                val x = Supbook()
                x.length = defaultsize.toShort()
                x.opcode = XLSConstants.SUPBOOK
                val dta = ByteArray(AddInProto.size)
                System.arraycopy(AddInProto, 0, dta, 0, AddInProto.size)
                x.setData(dta)
                x.originalsize = defaultsize
                x.init()
                return x
            }

        /**
         * creates a new External Supbook Record for the externalWorkbook
         *
         * @param externalWorkbook String
         * @return
         */
        fun getExternalPrototype(externalWorkbook: String): XLSRecord {
            val x = Supbook()
            x.length = defaultsize.toShort()
            x.opcode = XLSConstants.SUPBOOK
            x.originalsize = defaultsize
            // length of externalWorkBook + 4 (cstab + len) + number of encoding options (2 +)
            val compression = 0    //=non-unicode??
            val encoding = 1    //=Encoded URL follows
            val n = 2    //4;	// number of encoding chars
            val f = externalWorkbook.toByteArray()
            val ln = f.size + 1
            val dta = ByteArray(4 + n + ln - 1)
            System.arraycopy(ByteTools.shortToLEBytes(0.toShort()), 0, dta, 0, 2)   // cstab
            System.arraycopy(ByteTools.shortToLEBytes(ln.toShort()), 0, dta, 2, 2)  // ln
            var pos = 4
            dta[pos + 1] = compression.toByte()
            dta[pos + 2] = encoding.toByte()
            //dta[pos++]= 0x5;	// means unencoded URL follows
            //dta[pos++]= (byte) (ln-1);
            pos += n
            System.arraycopy(f, 0, dta, pos, ln - 1)
            x.setData(dta)
            x.init()
            return x
        }
    }
}