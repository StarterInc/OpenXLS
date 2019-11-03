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

/**
 * implementation of MD4 as RFC 1320 by R. Rivest, MIT Laboratory for
 * Computer Science and RSA Data Security, Inc.
 *
 *
 * **NOTE**: This algorithm is only included for backwards compatability
 * with legacy applications, it's not secure, don't use it for anything new!
 */

class MD4Digest {

    private var H1: Int = 0
    private var H2: Int = 0
    private var H3: Int = 0
    private var H4: Int = 0         // IV's
    private val X = IntArray(16)
    private var xOff: Int = 0
    private val xBuf: ByteArray
    private var xBufOff: Int = 0
    private var byteCount: Long = 0


    val byteLength: Int
        get() = BYTE_LENGTH

    /**
     * Standard constructor
     */
    init {
        xBuf = ByteArray(4)
        xBufOff = 0
        reset()
    }

    fun getDigest(data: ByteArray): ByteArray {

        this.update(data, 0, data.size)
        val digest = ByteArray(16)
        this.doFinal(digest, 0)

        return digest

    }


    protected fun processWord(
            `in`: ByteArray,
            inOff: Int) {
        X[xOff++] = (`in`[inOff] and 0xff or (`in`[inOff + 1] and 0xff shl 8)
                or (`in`[inOff + 2] and 0xff shl 16) or (`in`[inOff + 3] and 0xff shl 24))

        if (xOff == 16) {
            processBlock()
        }
    }

    protected fun processLength(
            bitLength: Long) {
        if (xOff > 14) {
            processBlock()
        }

        X[14] = (bitLength and -0x1).toInt()
        X[15] = bitLength.ushr(32).toInt()
    }

    private fun unpackWord(
            word: Int,
            out: ByteArray,
            outOff: Int) {
        out[outOff] = word.toByte()
        out[outOff + 1] = word.ushr(8).toByte()
        out[outOff + 2] = word.ushr(16).toByte()
        out[outOff + 3] = word.ushr(24).toByte()
    }

    fun doFinal(
            out: ByteArray,
            outOff: Int): Int {
        finish()

        unpackWord(H1, out, outOff)
        unpackWord(H2, out, outOff + 4)
        unpackWord(H3, out, outOff + 8)
        unpackWord(H4, out, outOff + 12)

        reset()

        return 16
    }

    /**
     * reset the chaining variables to the IV values.
     */
    fun reset() {
        byteCount = 0

        xBufOff = 0
        for (i in xBuf.indices) {
            xBuf[i] = 0
        }

        H1 = 0x67452301
        H2 = -0x10325477
        H3 = -0x67452302
        H4 = 0x10325476

        xOff = 0

        for (i in X.indices) {
            X[i] = 0
        }
    }

    /*
     * rotate int x left n bits.
     */
    private fun rotateLeft(
            x: Int,
            n: Int): Int {
        return x shl n or x.ushr(32 - n)
    }

    /*
     * F, G, H and I are the basic MD4 functions.
     */
    private fun F(
            u: Int,
            v: Int,
            w: Int): Int {
        return u and v or (u.inv() and w)
    }

    private fun G(
            u: Int,
            v: Int,
            w: Int): Int {
        return u and v or (u and w) or (v and w)
    }

    private fun H(
            u: Int,
            v: Int,
            w: Int): Int {
        return u xor v xor w
    }

    protected fun processBlock() {
        var a = H1
        var b = H2
        var c = H3
        var d = H4

        //
        // Round 1 - F cycle, 16 times.
        //
        a = rotateLeft(a + F(b, c, d) + X[0], S11)
        d = rotateLeft(d + F(a, b, c) + X[1], S12)
        c = rotateLeft(c + F(d, a, b) + X[2], S13)
        b = rotateLeft(b + F(c, d, a) + X[3], S14)
        a = rotateLeft(a + F(b, c, d) + X[4], S11)
        d = rotateLeft(d + F(a, b, c) + X[5], S12)
        c = rotateLeft(c + F(d, a, b) + X[6], S13)
        b = rotateLeft(b + F(c, d, a) + X[7], S14)
        a = rotateLeft(a + F(b, c, d) + X[8], S11)
        d = rotateLeft(d + F(a, b, c) + X[9], S12)
        c = rotateLeft(c + F(d, a, b) + X[10], S13)
        b = rotateLeft(b + F(c, d, a) + X[11], S14)
        a = rotateLeft(a + F(b, c, d) + X[12], S11)
        d = rotateLeft(d + F(a, b, c) + X[13], S12)
        c = rotateLeft(c + F(d, a, b) + X[14], S13)
        b = rotateLeft(b + F(c, d, a) + X[15], S14)

        //
        // Round 2 - G cycle, 16 times.
        //
        a = rotateLeft(a + G(b, c, d) + X[0] + 0x5a827999, S21)
        d = rotateLeft(d + G(a, b, c) + X[4] + 0x5a827999, S22)
        c = rotateLeft(c + G(d, a, b) + X[8] + 0x5a827999, S23)
        b = rotateLeft(b + G(c, d, a) + X[12] + 0x5a827999, S24)
        a = rotateLeft(a + G(b, c, d) + X[1] + 0x5a827999, S21)
        d = rotateLeft(d + G(a, b, c) + X[5] + 0x5a827999, S22)
        c = rotateLeft(c + G(d, a, b) + X[9] + 0x5a827999, S23)
        b = rotateLeft(b + G(c, d, a) + X[13] + 0x5a827999, S24)
        a = rotateLeft(a + G(b, c, d) + X[2] + 0x5a827999, S21)
        d = rotateLeft(d + G(a, b, c) + X[6] + 0x5a827999, S22)
        c = rotateLeft(c + G(d, a, b) + X[10] + 0x5a827999, S23)
        b = rotateLeft(b + G(c, d, a) + X[14] + 0x5a827999, S24)
        a = rotateLeft(a + G(b, c, d) + X[3] + 0x5a827999, S21)
        d = rotateLeft(d + G(a, b, c) + X[7] + 0x5a827999, S22)
        c = rotateLeft(c + G(d, a, b) + X[11] + 0x5a827999, S23)
        b = rotateLeft(b + G(c, d, a) + X[15] + 0x5a827999, S24)

        //
        // Round 3 - H cycle, 16 times.
        //
        a = rotateLeft(a + H(b, c, d) + X[0] + 0x6ed9eba1, S31)
        d = rotateLeft(d + H(a, b, c) + X[8] + 0x6ed9eba1, S32)
        c = rotateLeft(c + H(d, a, b) + X[4] + 0x6ed9eba1, S33)
        b = rotateLeft(b + H(c, d, a) + X[12] + 0x6ed9eba1, S34)
        a = rotateLeft(a + H(b, c, d) + X[2] + 0x6ed9eba1, S31)
        d = rotateLeft(d + H(a, b, c) + X[10] + 0x6ed9eba1, S32)
        c = rotateLeft(c + H(d, a, b) + X[6] + 0x6ed9eba1, S33)
        b = rotateLeft(b + H(c, d, a) + X[14] + 0x6ed9eba1, S34)
        a = rotateLeft(a + H(b, c, d) + X[1] + 0x6ed9eba1, S31)
        d = rotateLeft(d + H(a, b, c) + X[9] + 0x6ed9eba1, S32)
        c = rotateLeft(c + H(d, a, b) + X[5] + 0x6ed9eba1, S33)
        b = rotateLeft(b + H(c, d, a) + X[13] + 0x6ed9eba1, S34)
        a = rotateLeft(a + H(b, c, d) + X[3] + 0x6ed9eba1, S31)
        d = rotateLeft(d + H(a, b, c) + X[11] + 0x6ed9eba1, S32)
        c = rotateLeft(c + H(d, a, b) + X[7] + 0x6ed9eba1, S33)
        b = rotateLeft(b + H(c, d, a) + X[15] + 0x6ed9eba1, S34)

        H1 += a
        H2 += b
        H3 += c
        H4 += d

        //
        // reset the offset and clean out the word buffer.
        //
        xOff = 0
        for (i in X.indices) {
            X[i] = 0
        }
    }


    fun update(
            `in`: Byte) {
        xBuf[xBufOff++] = `in`

        if (xBufOff == xBuf.size) {
            processWord(xBuf, 0)
            xBufOff = 0
        }

        byteCount++
    }

    fun update(
            `in`: ByteArray,
            inOff: Int,
            len: Int) {
        var inOff = inOff
        var len = len
        //
        // fill the current word
        //
        while (xBufOff != 0 && len > 0) {
            update(`in`[inOff])

            inOff++
            len--
        }

        //
        // process whole words.
        //
        while (len > xBuf.size) {
            processWord(`in`, inOff)

            inOff += xBuf.size
            len -= xBuf.size
            byteCount += xBuf.size.toLong()
        }

        //
        // load in the remainder.
        //
        while (len > 0) {
            update(`in`[inOff])

            inOff++
            len--
        }
    }

    fun finish() {
        val bitLength = byteCount shl 3

        //
        // add the pad bytes.
        //
        update(128.toByte())

        while (xBufOff != 0) {
            update(0.toByte())
        }

        processLength(bitLength)

        processBlock()
    }

    companion object {
        private val BYTE_LENGTH = 64

        //
        // round 1 left rotates
        //
        private val S11 = 3
        private val S12 = 7
        private val S13 = 11
        private val S14 = 19

        //
        // round 2 left rotates
        //
        private val S21 = 3
        private val S22 = 5
        private val S23 = 9
        private val S24 = 13

        //
        // round 3 left rotates
        //
        private val S31 = 3
        private val S32 = 9
        private val S33 = 11
        private val S34 = 15
    }


}






