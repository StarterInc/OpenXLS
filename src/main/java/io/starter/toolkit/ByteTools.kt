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
package io.starter.toolkit

import java.io.*


/**
 * Helper methods for working with byte arrays and XLS files.
 */

class ByteTools : Serializable {

    /**
     * C Longs are only 32 bits, Java Longs are 64.
     * This method converts a 32-bit 'C' long to a
     * byte array.
     *
     *
     * Also performs 'little-endian' conversion.
     */
    fun cLongToLEBytesOLD(i: Int): ByteArray {
        //if(true)return Integer.
        val sbuf = cLongToLEShorts(i)
        val b1 = shortToLEBytes(sbuf[0])
        val b2 = shortToLEBytes(sbuf[1])

        val bbuf = ByteArray(4)
        bbuf[0] = b1[0]
        bbuf[1] = b1[1]
        bbuf[2] = b2[0]
        bbuf[3] = b2[1]

        // System.arraycopy(b1, 0, bbuf, 0, 2);
        // System.arraycopy(b2, 0, bbuf, 2, 2);
        return bbuf
    }

    companion object {


        /**
         * serialVersionUID
         */
        private const val serialVersionUID = 1220042103372057083L

        /**
         * Returns a string representation of the byte array.
         *
         * @param bt     the byte array
         * @param offset - offset into byte array
         * @return the string representation
         */
        fun getByteDump(bt: ByteArray, offset: Int): String {
            return ByteTools.getByteDump(bt, offset, bt.size)
        }

        /**
         * Returns a string representation of the byte array.
         *
         * @param bt     the byte array
         * @param offset - offset into byte array
         * @param len    - length of byte array segment to return
         * @return the string representation
         */
        fun getByteDump(bt: ByteArray?, offset: Int, len: Int): String {
            var offset = offset
            if (bt == null)
                return ""
            val buf = StringBuffer()
            var every4 = 0
            var every16 = 0
            buf.append("\r\n")// start on a new line
            var offst = Integer.toHexString(offset)
            // now calculate where the 4 byte words should be offset so display matches...
            //int remainder = offset%16;
            while (offst.length < 4) offst = "0$offst"
            buf.append(offst)
            buf.append(":    ")
            val origOffset = offset
            offset += 16
            for (i in origOffset until /*origOffset+*/len) {
                buf.append(hexits[bt[i].ushr(4) and 0xf])
                buf.append(hexits[bt[i] and 0xf])
                buf.append(" ")
                every4++
                if (every4 == 4) {
                    every4 = 0
                    buf.append("  ")
                    every16++
                    if (every16 == 4) {
                        buf.append("\r\n")
                        offst = Integer.toHexString(offset)
                        while (offst.length < 4) offst = "0$offst"
                        buf.append(offst)
                        buf.append(":    ")
                        offset += 16
                        every16 = 0
                    }
                }
            }
            return buf.toString() + ""
        }

        private val hexits = "0123456789ABCDEF"

        /**
         * Returns a string representation of the byte array.
         *
         * @param bt  the byte array
         * @param pad whether to pad the strings so they align
         * @return the string representation
         */
        fun getByteString(bt: ByteArray, pad: Boolean): String {
            if (bt.size == 0) return "null"
            val ret = StringBuffer()
            for (x in bt.indices) {
                if (x % 8 == 0 && x > 0) ret.append("\r\n")
                //String bstr = Integer.toOctalString(bt[x]); // toBinaryString(bt[x]); // Byte.toString(bt[x]);
                var bstr = java.lang.Byte.toString(bt[x]) // toBinaryString(bt[x]); // Byte.toString(bt[x]);
                if (pad) while (bstr.length < 4) bstr = " $bstr"
                ret.append(bstr)
                ret.append(",")
            }
            ret.setLength(ret.length - 1)
            //ret.append();
            return ret.toString()
        }

        /**
         * Appends one byte array to another.
         * If either input (but not both) is null, a clone of the other will be
         * returned. This method is guaranteed to always return an array different
         * from either of those passed in.
         *
         * @param src  the array which will be appended to `dest`
         * @param dest the array to which `src` will be appended
         * @throws NullPointerException if both inputs are null
         */
        fun append(src: ByteArray?, dest: ByteArray?): ByteArray {
            // Deal with null input correctly
            if (src == null) return dest!!.clone()
            if (dest == null) return src.clone()

            val srclen = src.size
            val destlen = dest.size

            val ret = ByteArray(srclen + destlen)
            System.arraycopy(dest, 0, ret, 0, destlen)
            System.arraycopy(src, 0, ret, destlen, srclen)

            return ret
        }


        /**
         * append one byte array to an empty array
         * of the proper size
         * usage:
         * newarray = bytetool.append(sourcearray, destinationarray, position to start copy at);
         */
        fun append(src: ByteArray, dest: ByteArray?, pos: Int): ByteArray {
            var dest = dest
            var srclen = src.size
            if (dest == null) dest = ByteArray(srclen)
            val destlen = dest.size
            if (destlen < srclen) {
                Logger.logInfo("Your destination byte array is too small to copy into: srclen=$srclen: destlen=$destlen")
                srclen = destlen
            }
            System.arraycopy(src, 0, dest, pos, srclen)
            return dest
        }


        fun cLongToLEBytes(i: Int): ByteArray {
            val ret = ByteArray(4)
            ret[0] = (i and 0xff).toByte()
            ret[1] = (i shr 8 and 0xff).toByte()
            ret[2] = (i shr 16 and 0xff).toByte()
            ret[3] = (i shr 24 and 0xff).toByte()
            return ret
        }

        /**
         * C Longs are only 32 bits, Java Longs are 64.
         * This method converts a 32-bit 'C' long to a
         * pair of java shorts.
         *
         *
         * Also performs 'little-endian' conversion.
         */
        fun cLongToLEShorts(x: Int): ShortArray {
            val buf = ShortArray(2)
            val high = x.ushr(16).toShort()
            val low = x.toShort()
            buf[0] = low
            buf[1] = high
            // if(DEBUG)Logger.logInfo(Info( ("x=" + x + " high=" + high + " low=" + low );
            return buf
        }

        fun doubleToLEByteArray(d: Double): ByteArray {
            val bite = ByteArray(8) // A long is 8 bytes
            val l = java.lang.Double.doubleToLongBits(d)
            var i: Int
            var t: Long
            t = l // variable t will be shifted right each time thru the loop.
            i = bite.size - 1
            while (i > -1) { //High order byte will be in b[0]
                val irr = t and 0xff
                bite[i] = Integer.valueOf(irr.toInt())!!.toByte() //get the last 8 bits into the byte array.
                t = t shr 8 //Shifts the long 1 byte. Same as divide by 256
                i--
            }
            val ret = ByteArray(bite.size)
            for (x in bite.indices)
                ret[x] = bite[bite.size - 1 - x]
            return ret
        }

        fun doubleToByteArray(d: Double): ByteArray {
            val bite = ByteArray(8) // A long is 8 bytes
            val l = java.lang.Double.doubleToLongBits(d)
            var i: Int
            var t: Long
            t = l // variable t will be shifted right each time thru the loop.
            i = bite.size - 1
            while (i > -1) { //High order byte will be in b[0]
                val irr = t and 0xff
                bite[i] = Integer.valueOf(irr.toInt())!!.toByte() //get the last 8 bits into the byte array.
                t = t shr 8 //Shifts the long 1 byte. Same as divide by 256
                i--
            }
            return bite
        }

        /**
         * converts and bitswaps an eight bite byte array into an IEEE double.
         */
        fun eightBytetoLEDouble(bite: ByteArray): Double {
            val b = ByteArray(8)
            b[0] = bite[7]
            b[1] = bite[6]
            b[2] = bite[5]
            b[3] = bite[4]
            b[4] = bite[3]
            b[5] = bite[2]
            b[6] = bite[1]
            b[7] = bite[0]
            var d = 0.0
            val bais = java.io.ByteArrayInputStream(b)
            val dis = java.io.DataInputStream(bais)
            try {
                val dbl = dis.readDouble()
                d = dbl
            } catch (e: java.io.IOException) {
                Logger.logInfo("io exception in byte to Double conversion$e")
            }

            return d
        }


        /**
         * converts and bitswaps an eight bite byte array into an IEEE double.
         */
        fun eightBytetoLELong(bite: ByteArray): Long {
            val b = ByteArray(8)
            b[0] = bite[7]
            b[1] = bite[6]
            b[2] = bite[5]
            b[3] = bite[4]
            b[4] = bite[3]
            b[5] = bite[2]
            b[6] = bite[1]
            b[7] = bite[0]
            var l: Long = 0
            val bais = java.io.ByteArrayInputStream(b)
            val dis = java.io.DataInputStream(bais)
            try {
                val lg = dis.readLong()
                l = lg
            } catch (e: java.io.IOException) {
                Logger.logInfo("io exception in byte to Double conversion$e")
            }

            return l
        }

        /**
         * Get an array of bytes from a collection of byte arrays
         *
         *
         * Seems slow, why 2 iterations, should be faster way?
         */
        fun getBytes(records: List<*>): ByteArray {
            val e = records.iterator()
            var buflen = 0
            while (e.hasNext()) {
                val barr = e.next() as ByteArray
                buflen += barr.size
            }

            var outbytes = ByteArray(buflen)
            var pos = 0
            for (i in records.indices) {
                val stream = records[i] as ByteArray
                outbytes = append(stream, outbytes, pos)
                pos += stream.size
            }
            return outbytes
        }

        /*
       Makes sure unicode strings are in the correct format to match Excel's strings.
       If the string has all low order bytes as 0x0 then return original string, as we do not
       want that extra space.

    */
        fun getExcelEncoding(s: String): ByteArray {
            var strbytes: ByteArray? = null
            try {
                strbytes = s.toByteArray(charset("UnicodeLittleUnmarked"))
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("Error creating encoded string: $e")
            }

            var unicode = false
            var i = 0
            while (i < strbytes!!.size) {
                i = i + 1
                if (strbytes[i].toInt() != 0x0) {
                    unicode = true
                    i = strbytes.size
                }
                i++

            }
            return if (unicode) {
                //try{
                strbytes
                //}catch(UnsupportedEncodingException e){Logger.logInfo("Error creating encoded string: " + e);}
            } else s.toByteArray()
        }

        /**
         * This is a working longToLEByteArray.  I'm not sure whats up with the other ones
         *
         * @param l
         * @return
         */
        fun longToLEByteArray(l: Long): ByteArray {
            val bos = ByteArrayOutputStream()
            val dos = DataOutputStream(bos)
            try {
                dos.writeLong(l)
                dos.flush()
            } catch (e: IOException) {

            }

            val bite = bos.toByteArray()
            val b = ByteArray(8)
            b[0] = bite[7]
            b[1] = bite[6]
            b[2] = bite[5]
            b[3] = bite[4]
            b[4] = bite[3]
            b[5] = bite[2]
            b[6] = bite[1]
            b[7] = bite[0]
            return b
        }

        fun isUnicode(s: String): Boolean {
            var strbytes: ByteArray? = null
            try {
                strbytes = s.toByteArray(charset("UnicodeLittleUnmarked"))
            } catch (e: UnsupportedEncodingException) {
            }

            var i = 0
            while (i < strbytes!!.size) {
                if (strbytes[i] >= 0x7f)
                    return true  // deal with non-compressible Eastern Strings
                i = i + 1
                if (strbytes[i].toInt() != 0x0) {
                    return true // there is a non-zero high-byte
                }
                i++
            }
            return false
        }

        fun longToByteArray(l: Long): ByteArray {
            val bite = ByteArray(8) // A long is 8 bytes
            var i: Int
            var t: Long
            t = l // variable t will be shifted right each time thru the loop.
            i = bite.size - 1
            while (i > -1) { //High order byte will be in b[0]
                val irr = t and 0xff
                bite[i] = Integer.valueOf(irr.toInt())!!.toByte() //get the last 8 bits into the byte array.
                t = t shr 8 //Shifts the long 1 byte. Same as divide by 256
                i--
            }
            return bite
        }


        /**
         * same as readInt, but takes 4 raw bytes instead of 2 shorts
         */
        fun readInt(bs: ByteArray): Int {
            return readInt(readShort(bs[2].toInt(), bs[3].toInt()).toInt(), readShort(bs[0].toInt(), bs[1].toInt()).toInt())
        }

        /**
         * same as readInt, but takes 4 raw bytes instead of 2 shorts
         */
        fun readInt(b1: Byte, b2: Byte, b3: Byte, b4: Byte): Int {
            return readInt(readShort(b3.toInt(), b4.toInt()).toInt(), readShort(b1.toInt(), b2.toInt()).toInt())
        }

        /**
         * Reads a 4 byte int from a byte array at the specified position
         * and handles a little endian conversion
         */
        fun readInt(b: ByteArray, offset: Int): Int {
            var offset = offset
            return readInt(b[offset++], b[offset++], b[offset++], b[offset++])
        }


        /**
         * bit-flipping action converting a 'little-endian'
         * pair of shorts to a 'big-endian' long.
         * This is really a java int as it represents a C-language
         * long value which is only 32 bits, like the java int.
         */
        fun readInt(low: Int, high: Int): Int {
            var low = low
            var high = high
            if (low == 0x0 && high == 0x0) return 0
            low = low and 0xffff
            high = high and 0xffff
            return low shl 16 or high
        }


        /**
         * bit-flipping action converting a 'little-endian'
         * pair of bytes to a 'big-endian' short.  Returns an int as
         * excel uses unsigned shorts which can exceed the boundary
         * of a java signed short
         */
        fun readUnsignedShort(low: Byte, high: Byte): Int {
            return readInt(low, high, 0x0.toByte(), 0x0.toByte())
        }

        /**
         * bit-flipping action converting a 'little-endian'
         * pair of bytes to a 'big-endian' short.
         *
         *
         * This will break if you pass it any values larger than a byte.  Will
         * probably return a value, but I wouldn't trust it.  Fix in R2  -Rab
         */
        fun readShort(low: Int, high: Int): Short {
            var low = low
            var high = high
            // 2 bytes
            low = low and 0xff
            high = high and 0xff
            return (high shl 8 or low).toShort()
        }


        /** bit-flipping action converting a 'little-endian'
         * pair of bytes to a 'big-endian' short.
         *
         * public static short readShort(byte low, byte high)
         * {
         * return (short)(high << 8 | low);
         * }
         */

        /**
         * take 16-bit short apart into two 8-bit bytes.
         */
        fun shortToLEBytes(x: Short): ByteArray {
            val buf = ByteArray(2)
            buf[1] = x.ushr(8).toByte()
            buf[0] = x.toByte()/* cast implies & 0xff */
            return buf
        }

        fun toBEByteArray(d: Double): ByteArray {
            val bite = ByteArray(8) // A long is 8 bytes
            val l = java.lang.Double.doubleToLongBits(d)
            var i: Int
            var t: Long
            t = l // variable t will be shifted right each time thru the loop.
            i = bite.size - 1
            while (i > -1) { //High order byte will be in b[0]
                val irr = t and 0xff
                bite[i] = Integer.valueOf(irr.toInt())!!.toByte() //get the last 8 bits into the byte array.
                t = t shr 8 //Shifts the long 1 byte. Same as divide by 256
                i--
            }
            val b = ByteArray(8)
            b[0] = bite[7]
            b[1] = bite[6]
            b[2] = bite[5]
            b[3] = bite[4]
            b[4] = bite[3]
            b[5] = bite[2]
            b[6] = bite[1]
            b[7] = bite[0]
            return b
        }
        //  private boolean DEBUG = false;

        /**
         * write bytes to a file
         */
        fun writeToFile(b: ByteArray, fname: String) {
            try {
                val fos = FileOutputStream(fname)
                fos.write(b)
                fos.flush()
                fos.close()
            } catch (e: IOException) {
                Logger.logInfo("Error writing bytes to file in ByteTools: $e")
            }

        }


        // NICK -- let's put the following in a JUnit Test.  -jm

        /** this is a good test for some of the above methods..
         * byte[] bytes;
         * java.io.ByteArrayInputStream bis = new java.io.ByteArrayInputStream(bite);
         * java.io.DataInputStream dis = new java.io.DataInputStream(bis);
         * try {
         * long lon = dis.readLong();
         * Logger.logInfo("pleeze stop here");
         * } catch (java.io.IOException e){ Logger.logInfo("io exception in byte to long conversion" + e);}
         */
        /**
         * generic method to update (set or clear) a short at bitNum
         *
         * @param set
         * @param bitNum
         */
        fun updateGrBit(grbit: Short, set: Boolean, bitNum: Int): Short {
            var grbit = grbit
            when (bitNum) {
                0 -> if (set)
                    grbit = grbit or 0x1.toShort()
                else
                    grbit = grbit and 0xFFFE.toShort()
                1 -> if (set)
                    grbit = grbit or 0x2.toShort()
                else
                    grbit = grbit and 0xFFFD.toShort()
                2 -> if (set)
                    grbit = grbit or 0x4.toShort()
                else
                    grbit = grbit and 0xFFFB.toShort()
                3 -> if (set)
                    grbit = grbit or 0x8.toShort()
                else
                    grbit = grbit and 0xFFF7.toShort()
                4 -> if (set)
                    grbit = grbit or 0x10.toShort()
                else
                    grbit = grbit and 0xFFEF.toShort()
                5 -> if (set)
                    grbit = grbit or 0x20.toShort()
                else
                    grbit = grbit and 0xFFDF.toShort()
                6 -> if (set)
                    grbit = grbit or 0x40.toShort()
                else
                    grbit = grbit and 0xFFBF.toShort()
                7 -> if (set)
                    grbit = grbit or 0x80.toShort()
                else
                    grbit = grbit and 0xFF7F.toShort()
                8 -> if (set)
                    grbit = grbit or 0x100.toShort()
                else
                    grbit = grbit and 0xFEFF.toShort()
                9 -> if (set)
                    grbit = grbit or 0x200.toShort()
                else
                    grbit = grbit and 0xFDFF.toShort()
                10 -> if (set)
                    grbit = grbit or 0x400.toShort()
                else
                    grbit = grbit and 0xFBFF.toShort()
                11 -> if (set)
                    grbit = grbit or 0x800.toShort()
                else
                    grbit = grbit and 0xF7FF.toShort()
                12 -> if (set)
                    grbit = grbit or 0x1000.toShort()
                else
                    grbit = grbit and 0xEFFF.toShort()
                13 -> if (set)
                    grbit = grbit or 0x2000.toShort()
                else
                    grbit = grbit and 0xDFFF.toShort()
                14 -> if (set)
                    grbit = grbit or 0x4000.toShort()
                else
                    grbit = grbit and 0xBFFF.toShort()
                15 -> if (set)
                    grbit = grbit or 0x8000.toShort()
                else
                    grbit = grbit and 0x7FFF.toShort()
            }
            return grbit
        }
    }
}

     