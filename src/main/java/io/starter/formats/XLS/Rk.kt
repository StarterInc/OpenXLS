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

import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.text.NumberFormat


/**
 * **RK: RK Number (7Eh)**<br></br>
 * This record stores an internal numeric type.  Stores data in one of four
 * RK 'types' which determine whether it is an integer or an IEEE floating point
 * equivalent.
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number of the RK record
 * 8       ixfe        2       Index to XF cell format record
 * 10      rk          4       RK number
</pre> *
 *
 * @see MULRK
 *
 * @see NUMBER
 */

class Rk : XLSCellRecord, Mulled {
    var type: Int = 0
        internal set
    internal var Rkdouble: Double = 0.toDouble()
    internal var RKint: Int = 0
    internal var mymul: Mulrk? = null

    var DEBUG = false

    override var myMul: Mul?
        get() = mymul
        set(m) {
            mymul = m as Mulrk
        }

    /**
     * returns the position of this record in the array of records
     * making up this file.
     */
    override// this is a MulRk
    // throw new InvalidRecordException("Rk without a recidx nor a Mulrk.");
    // standalone RK
    val recordIndex: Int
        get() = if (super.recordIndex < 0) {
            if (mymul == null) -1 else mymul!!.recordIndex
        } else
            super.recordIndex

    /**
     * @see io.starter.formats.XLS.XLSRecord.setIntVal
     */
    override var intVal: Int
        @Throws(RuntimeException::class)
        get() {
            if (isFPNumber) {
                val l = Rkdouble.toLong()
                if (l > Integer.MAX_VALUE) {
                    throw NumberFormatException("Cell value is larger than the maximum java signed int size")
                }
                if (l < Integer.MIN_VALUE) {
                    throw NumberFormatException("Cell value is smaller than the minimum java signed int size")
                }
                return Rkdouble.toInt()
            }
            return RKint
        }
        set(f) = try {
            this.setRKVal(f.toDouble())
        } catch (x: Exception) {
            Logger.logWarn("Rk.setIntVal() problem.  Fallback to floating point Number.")
            Rk.convertRkToNumber(this, f.toDouble())
        }

    override val dblVal: Double
        get() = if (isIntNumber) RKint.toDouble() else Rkdouble

    /**
     * @see io.starter.formats.XLS.XLSRecord.setFloatVal
     */
    override var floatVal: Float
        get() = if (isIntNumber) RKint.toFloat() else Rkdouble.toFloat()
        set(f) = try {
            this.setRKVal(f.toDouble())
        } catch (x: Exception) {
            Logger.logWarn("Rk.setFloatVal() problem.  Fallback to floating point Number.")
            Rk.convertRkToNumber(this, f.toDouble())
        }

    /**
     * Return the string value.  If it is over 99999999999 or under -99999999999
     * then return a format using scientific notation.  Note this is *not* significant digits,
     * rather the actual size of the number.  Emulates Excel functionality and display
     */
    override// 20080211 KSC: Double.valueOf(s).doubleValue();
    var stringVal: String?
        get() = if (isIntNumber) {
            RKint.toString()
        } else {
            ExcelTools.getNumberAsString(Rkdouble)
        }
        set(s) = try {
            if (s.indexOf(".") > -1) {
                val f = Double(s)
                this.setDoubleVal(f)
            } else {
                val i = Integer.parseInt(s)
                this.intVal = i
            }
        } catch (f: java.lang.NumberFormatException) {
            Logger.logWarn("in Rk $s is not a number.")
        }

    val typeName: String
        get() = "Rkdouble"

    override fun setNoMul() {
        mymul = null
    }

    /**
     * default constructor
     */
    constructor() : super() {}


    /**
     * Provide constructor which automatically
     * sets the body data and header info.  This
     * is needed by MULRK which creates the RKs without
     * the benefit of WorkBookFactory.parseRecord().
     */
    internal constructor(b: ByteArray, r: Int, c: Int) {
        rw = r
        colNumber = c.toShort()
        setData(b)
        opcode = XLSConstants.RK
        length = 10.toShort()
        this.init(b)
    }

    /**
     * This init method
     * is needed by MULRK which creates the RKs without
     * the benefit of WorkBookFactory.parseRecord().
     */
    // called by Mulrk.init
    internal fun init(b: ByteArray, r: Int, c: Int) {
        rw = r
        colNumber = c.toShort()
        val rwbt = ByteTools.shortToLEBytes(r.toShort())
        val colbt = ByteTools.shortToLEBytes(c.toShort())
        val newData = ByteArray(10)
        newData[0] = rwbt[0]
        newData[1] = rwbt[1]
        newData[2] = colbt[0]
        newData[3] = colbt[1]
        System.arraycopy(b, 0, newData, 4, b.size)
        setData(newData)
        opcode = XLSConstants.RK
        this.init(b)
    }

    /**
     * This init method pulls out the record header information,
     * then sends the as-yet unmodded rkdata record across to the
     * rktranslate method
     */
    internal fun init(rkdata: ByteArray) {
        super.init()
        val s: Short
        val rknum = ByteArray(4)
        // if this is a 'standalone' RK number, then the byte array
        // contains row, col and ixfe data as well as the number value.
        if (rkdata.size > 6) {
            // get the row information
            super.initRowCol()
            s = ByteTools.readShort(rkdata[4].toInt(), rkdata[5].toInt())
            ixfe = s.toInt()
            System.arraycopy(rkdata, 6, rknum, 0, 4)
        } else {
            // get the ixfe information
            s = ByteTools.readShort(rkdata[0].toInt(), rkdata[1].toInt())
            ixfe = s.toInt()
            System.arraycopy(rkdata, 2, rknum, 0, 4)
        }
        this.translateRK(rknum)
    }

    /**
     * This init method pulls out the record header information,
     * then sends the as-yet unmodded rkdata record across to the
     * rktranslate method
     */
    override fun init() {
        super.init()
        this.getData()
        val s: Short
        val rknum = ByteArray(4)
        // if this is a 'standalone' RK number, then the byte array
        // contains row, col and ixfe data as well as the number value.
        if (this.length > 6) {
            // get the row information

            super.initRowCol()
            s = ByteTools.readShort(getByteAt(4).toInt(), getByteAt(5).toInt())
            ixfe = s.toInt()
            val numdat = this.getBytesAt(6, 4)
            System.arraycopy(numdat!!, 0, rknum, 0, 4)
        } else {
            // get the ixfe information
            s = ByteTools.readShort(getByteAt(0).toInt(), getByteAt(1).toInt())
            ixfe = s.toInt()
            val numdat = this.getBytesAt(2, 4)
            System.arraycopy(numdat!!, 0, rknum, 0, 4)
        }
        this.translateRK(rknum)
    }


    /**
     * figures out the type of RK that we have, does the
     * little endian thing, then sends it over to getRealVal
     */
    internal fun translateRK(rkval: ByteArray) {

        var l = 0
        val num1 = ByteTools.readShort(rkval[0].toInt(), rkval[1].toInt())
        val num2 = ByteTools.readShort(rkval[2].toInt(), rkval[3].toInt())
        l = ByteTools.readInt(num2.toInt(), num1.toInt())

        val num = l.toLong()
        // num = num << 1;
        // num = num >>> 1;

        // check what the RK type bits are
        val bitset = l and (1 shl 0)
        val bitset2 = l and (1 shl 1)
        // add them to get the type
        type = bitset + bitset2
        val d = 1.0

        Rkdouble = Rk.getRealVal(type, num)
        if (DEBUG) Logger.logInfo(type.toString())
        if (DEBUG) Logger.logInfo(Rkdouble.toString())
        this.isValueForCell = true
        this.isDoubleNumber = false
        when (type) {
            Rk.RK_FP -> {
                // okay, am i dense or something or is
                // 1 NOT a Float?  see RK Type 0 on pg. 377
                // then tell me I'm not on crack... -jm 11/03
                var newnum = Rkdouble.toString()
                if (newnum.length > 12) this.isDoubleNumber = true
                var mantindex = newnum.indexOf(".")
                newnum = newnum.substring(mantindex + 1)
                try {
                    if (Integer.parseInt(newnum) > 0) { // there's FP digits
                        this.isFPNumber = true
                        this.isIntNumber = false
                    } else {
                        this.isFPNumber = false
                        this.isIntNumber = true
                        RKint = Rkdouble.toInt()
                    }
                } catch (e: NumberFormatException) {
                    this.isFPNumber = true
                    this.isIntNumber = false


                    //RKint = (int)Rkdouble;
                }

                if (Rkdouble > java.lang.Float.MAX_VALUE) {
                    this.isDoubleNumber = true
                }
            }

            Rk.RK_FP_100 -> {
                this.isFPNumber = true
                this.isIntNumber = false
            }
            Rk.RK_INT -> {
                this.isFPNumber = false
                this.isIntNumber = true
                RKint = Rkdouble.toInt()
            }
            Rk.RK_INT_100 -> {
                newnum = Rkdouble.toString()
                if (newnum.toUpperCase().indexOf("E") > -1) {
                    // do something intelligent
                    val nmf = NumberFormat.getInstance()
                    try {
                        val nm = nmf.parse(newnum)
                        val v = nm.toFloat()
                        if (this.DEBUGLEVEL > 5) Logger.logInfo("Rk number format: $v")
                    } catch (e: Exception) {
                    }

                } else {
                    mantindex = newnum.indexOf(".")
                    newnum = newnum.substring(mantindex + 1)
                }
                try {
                    if (java.lang.Long.parseLong(newnum) > 0) { // there's FP digits
                        this.isFPNumber = true
                        this.isIntNumber = false
                    } else {
                        this.isFPNumber = false
                        this.isIntNumber = true
                        RKint = Rkdouble.toInt()
                    }
                    //					happens with big numbers with exponents
                    //					they should be ints
                } catch (e: NumberFormatException) {
                    this.isFPNumber = false
                    this.isIntNumber = true
                    RKint = Rkdouble.toInt()
                }

            }
        }
    }

    /**
     * @see io.starter.formats.XLS.XLSRecord.setDoubleVal
     */
    override fun setDoubleVal(f: Double) {
        try {
            this.setRKVal(f)
        } catch (x: Exception) {
            Logger.logWarn("Rk.setDoubleVal() problem.  Fallback to floating point Number.")
            Rk.convertRkToNumber(this, f)
        }

    }

    /**
     * Change the row of this record, including parent Mulrk if it exists.
     */
    internal fun setMulrkRow(i: Int) {
        super.rowNumber = i
        if (mymul != null) {
            mymul!!.setRow(i)
        }
    }


    /**
     * Allows writing back to the RKRec.  In order to be written correctly
     * the RK must conform to one of the cases below, if it does not it will need
     * to be handled as a NUMBER.  Note the last 2 bits of each record need to be
     * set as an identifier for what type of RK record we are writing.  Types 0 and 1
     * are stored as a 32 bit modified float, with the mantissa of a double and the final
     * 2 bits changed to the RK type.  Types 2 and 3 are stored as a 30 bit integer with the
     * final 2 bits as RK type.  Also note types 1 and 2 are stored in the excel file as x*100.
     */
    protected fun setRKVal(d: Double) {
        val b = Rk.getRkBytes(d)    // returns a 5-byte structure
        this.type = b[4].toInt()    // the last byte is the Rk type

        when (this.type) {
            Rk.RK_FP -> this.isFPNumber = true
            Rk.RK_INT -> this.isIntNumber = true
            Rk.RK_FP_100 -> this.isFPNumber = true
            Rk.RK_INT_100 -> this.isIntNumber = true
            else -> {
                // we need to convert this into a number rec -- not an rk.
                Rk.convertRkToNumber(this, d)
                return         // it's not an Rk anymore so split
            }
        }

        System.arraycopy(b, 0, this.getData()!!, 6, 4)
        this.init(this.getData()!!)

        // failsafe... if for any reason it did not work
        if (this.Rkdouble != d) {
            Logger.logWarn("Rk.setRKVal() problem.  Fallback to floating point Number.")
            Rk.convertRkToNumber(this, d)
        }
    }


    /** return a prototype RK record
     *
     * offset  name        size    contents
     * ---
     * 4       rw          2       Row Number
     * 6       col         2       Column Number of the RK record
     * 8       ixfe        2       Index to XF cell format record
     * 10      rk          4       RK number
     *
     * public static XLSRecord getPrototype(){
     *
     * Rk newrec = new Rk();
     * // for larger records, we need to read in from a file...
     * byte[] protobytes = {(byte)0x2,
     * (byte)0x0,
     * (byte)0x0,
     * (byte)0x0,
     * (byte)0xf,
     * (byte)0x0,
     * (byte)0x1,
     * (byte)0xc0,
     * (byte)0x5e,
     * (byte) 0x40};
     * newrec.setOpcode(RK);
     * newrec.setLength((short) 0xa);
     * newrec.setData(protobytes);
     * newrec.init(newrec.getData());
     * return newrec;
     * }
     */

    /**
     * set the XF (format) record for this rec
     */
    override fun setXFRecord(i: Int) {
        if (mymul != null) {
            val b = this.getData()
            this.ixfe = i
            val newxfe = ByteTools.cLongToLEBytes(i)
            System.arraycopy(newxfe, 0, b!!, 4, 2)
            mymul!!.updateRks()
        } else {
            super.setIxfe(i)
        }
        super.setXFRecord()
    }

    override fun close() {
        super.close()
        if (mymul != null)
            mymul!!.close()
        mymul = null
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3027662614434608240L

        val RK_FP = 0
        val RK_FP_100 = 1
        val RK_INT = 2
        val RK_INT_100 = 3


        /**
         * static method which parses a 4-byte RK number
         * into a double value using specific MS Rules
         *
         * @param byte[] rkbytes - 4 byte Rk Number
         * @return double - translated from Rk bytes
         * @see Rk
         */
        fun parseRkNumber(rkbytes: ByteArray): Double {
            var num = 0
            val num1 = ByteTools.readShort(rkbytes[0].toInt(), rkbytes[1].toInt())
            val num2 = ByteTools.readShort(rkbytes[2].toInt(), rkbytes[3].toInt())
            num = ByteTools.readInt(num2.toInt(), num1.toInt())

            // check what the RK type bits are
            val bitset = num and (1 shl 0)
            val bitset2 = num and (1 shl 1)
            // add them to get the type
            val rkType = bitset + bitset2

            return Rk.getRealVal(rkType, num.toLong())
        }

        /**
         * Madness with Microsoft trying to figure out its own format for floating point.
         * to really understand what is going on RTFM, but basically there are four different
         * encodings used for RKs, a modified int, which uses its last two bits as an identifer,
         * this same int divided by 100, a modified float with the Exponent of a double, but with the
         * last 2 bits used as an identifier, and finally the same as the last divided by 100. Hmmm
         */
        private fun getRealVal(RKType: Int, waknum: Long): Double {
            var waknum = waknum
            // NOTE isFPNumber, isIntNumber was set in translateRk anyways so take out of here
            // we need to mask the number to avoid the final 2 bits messin stuff up
            waknum = waknum and -0x4
            // perform the change for each type
            when (RKType) {
                Rk.RK_FP -> {
                    // IEEE Number
                    waknum = waknum shl 32
                    val testq = java.lang.Long.toBinaryString(waknum)
                    return java.lang.Double.longBitsToDouble(waknum)
                }

                Rk.RK_FP_100 -> {
                    var res = java.lang.Double.longBitsToDouble(waknum shl 32)
                    res /= 100.0
                    return res
                }

                Rk.RK_INT -> {
                    // Integer
                    waknum = waknum shr 2
                    return waknum.toDouble()
                }

                Rk.RK_INT_100 -> {
                    if (waknum >= 4290773292L) {
                        Logger.logWarn("Erroneous Rk") // THIS IS THE CUTOFF NUMBER -- ANYTHING THIS SIZE IS < -10,485.01
                    }
                    // Integer x 100
                    waknum = waknum shr 2
                    val ddd = waknum.toDouble()
                    return ddd / 100
                }

                else -> Logger.logInfo("incorrect RK type for RK record: $RKType")
            }
            return 0.0
        }


        /**
         * static method returns the double value converted to Rk-type bytes (a 4-byte structure)
         * as well as the Rk Type (FP, INT, FP_100, INT_100) in the 5th byte position
         * @param  double d - double to convert to Rk
         * @return byte[] array of bytes representing the Rk number (4 bytes) and the Rk type (1 byte)
         */
        /**
         * Structure of RkNumber:
         * A - fX100 (1 bit): A bit that specifies whether num is the value of the RkNumber or 100 times the value of the RkNumber. MUST be a value from the following table:
         * 0 = The value of RkNumber is the value of num.
         * 1 = The value of RkNumber is the value of num divided by 100.
         * B - fInt (1 bit): A bit that specifies the type of num.
         * 0 =    num is the 30 most significant bits of a 64-bit binary floating-point number as defined in [IEEE754].
         * The remaining 34-bits of the floating-point number MUST be 0.
         * 1 =    num is a signed integer
         * num (30 bits): A variable type field whose type and meaning is specified by the value of fInt, as defined in the following table:
         */
        fun getRkBytes(d: Double): ByteArray {
            val bitlong = java.lang.Double.doubleToLongBits(d)
            val bigger = java.lang.Double.doubleToLongBits(d * 100)
            var l = d.toLong()
            l = java.lang.Math.abs(l)

            val rkbytes = ByteArray(5)
            rkbytes[4] = -1    // uninitialized type
            var doublebytes = ByteArray(8)

            // are low order 34 bits of d = 0?  RK type = 0 (RK_FP)
            // d is a 64 bit num, so move it over a bit...
            if (bitlong shl 30 == 0L) {
                doublebytes = ByteTools.doubleToByteArray(d)
                // add the RK type at the end of the last byte
                val mask = 0xfc.toByte()
                doublebytes[3] = (doublebytes[3] and mask).toByte()
                // bit flipping for the LE switch
                rkbytes[0] = doublebytes[3]
                rkbytes[1] = doublebytes[2]
                rkbytes[2] = doublebytes[1]
                rkbytes[3] = doublebytes[0]
                rkbytes[4] = RK_FP.toByte()
            } else if ((l shl 2).ushr(30) == 0L && d % 1 == 0.0) {
                var lo = d.toLong()
                lo = lo shl 2
                doublebytes = ByteTools.longToByteArray(lo)        // RK_INT
                val mask = 0x2.toByte()
                doublebytes[7] = (doublebytes[7] or mask).toByte()
                rkbytes[0] = doublebytes[7]
                rkbytes[1] = doublebytes[6]
                rkbytes[2] = doublebytes[5]
                rkbytes[3] = doublebytes[4]
                rkbytes[4] = RK_INT.toByte()
            }/**/
            else if (bigger shl 30 == 0L) {
                doublebytes = ByteTools.doubleToByteArray(d * 100)        // F100
                val mask = 0x1.toByte()
                doublebytes[3] = (doublebytes[3] or mask).toByte()
                rkbytes[0] = doublebytes[3]
                rkbytes[1] = doublebytes[2]
                rkbytes[2] = doublebytes[1]
                rkbytes[3] = doublebytes[0]
                rkbytes[4] = RK_FP_100.toByte()
            } else if (d * 100 % 1 == 0.0 && (l * 100).ushr(30) == 0L) {
                var lo = (d * 100).toLong()        // F100 + INT
                lo = lo shl 2
                doublebytes = ByteTools.longToByteArray(lo)
                val mask = 0x3.toByte()
                doublebytes[7] = (doublebytes[7] or mask).toByte()
                rkbytes[0] = doublebytes[7]
                rkbytes[1] = doublebytes[6]
                rkbytes[2] = doublebytes[5]
                rkbytes[3] = doublebytes[4]
                rkbytes[4] = RK_INT_100.toByte()
            }// Can d * 100 be represented by a 30 bit integer? RK type = 3 (RK_INT_100)
            // are low order 34 bits of d * 100 = 0?  RK type = 1 (RK_FP_100)
            // Can d be represented by a 30 bit integer? RK Type = 2 (RK_INT)
            // ORIGINAL -- else if ((l>>>30) ==0 && ((d%1)==0)){
            if (rkbytes[4].toInt() != -1 && d != Rk.parseRkNumber(rkbytes))
            // if it was processed as an RK, ensure results are accurate
                throw RuntimeException("$d could not be translated to Rk value")

            return rkbytes
        }

        /**
         * Converts an RK valrec to a number record.  This allows the number format to be changed within the cell
         */
        fun convertRkToNumber(reek: Rk, d: Double) {
            val fmt = reek.ixfe
            val addy = reek.cellAddress
            val bs = reek.sheet
            bs!!.removeCell(reek)
            val addedrec = bs.addValue(d, addy)
            addedrec.ixfe = reek.getIxfe()
        }

        // DEBUGGING - throws exception if

        /**
         * internal debugging method
         */
        fun testVALUES() {
            for (v in Integer.MAX_VALUE downTo Integer.MIN_VALUE) {
                Rk.getRkBytes(v.toDouble())    // throws Exception if converted bytes do not match original value
            }
            /* problem values:			for (long v=1073741824; v > 536870912; v--)
				Rk.getRkBytes(v);	// throws Exception if converted bytes do not match original value
*/
        }
    }

}