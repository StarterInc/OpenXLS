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

import java.io.UnsupportedEncodingException
import java.nio.charset.StandardCharsets

/**
 * Password - specifies the password verifier for the sheet or workbook.
 * This record is required in the Globals sub-stream. If the workbook has no
 * password the value written is zero. If a sheet has a password this record
 * appears in its sub-stream. In a sheet sub-stream the value may not be zero.
 *
 * @see MS-XLS §2.4.191
 */
class Password : XLSRecord() {

    /**
     * Returns the password verifier value.
     * This is the password after being passed through a rather strange hash
     * function that reduces it to a two byte integer.
     */
    var passwordHash: Short = 0
        private set

    /**
     * Gets the password verifier as a hexadecimal string.
     *
     * @return the verifier as four upper-case hexits (0123456789ABCDEF)
     * @see .getPasswordHash
     */
    val passwordHashString: String
        get() {
            val raw = Integer.toHexString(passwordHash and 0xFFFF).toUpperCase()
            return "0000".substring(0, 4 - raw.length) + raw
        }

    init {
        opcode = XLSConstants.PASSWORD
        length = 2.toShort()
        this.originalsize = 2
    }

    /**
     * Checks whether the given password matches the stored one.
     * This hashes the given guess and then compares it against the stored
     * value. Note that the hash is only two bytes so collisions are likely.
     *
     * @param guess the clear text password guess to be verified
     * @return whether the hash of the guess matches the stored value
     */
    fun validatePassword(guess: String): Boolean {
        return hashPassword(guess) == passwordHash
    }

    /**
     * Sets the stored password.
     *
     * @param password the clear text of the password to be applied
     * or null to remove the existing password
     */
    fun setPassword(password: String) {
        passwordHash = hashPassword(password)
        updateRecord()
    }

    /**
     * Sets the stored password.
     *
     * @param hash the hash of the password to be applied
     * or zero to remove the existing password
     */
    fun setHashedPassword(hash: Short) {
        this.passwordHash = hash
        updateRecord()
    }

    /**
     * Sets the stored password.
     *
     * @param hash the four-digit hex hash of the password to be applied
     * or "0000" to remove the existing password
     */
    fun setHashedPassword(hash: String) {
        this.passwordHash = Integer.parseInt(hash, 16).toShort()
        updateRecord()
    }

    private fun updateRecord() {
        setData(ByteTools.shortToLEBytes(passwordHash))
    }

    override fun init() {
        super.init()
        passwordHash = ByteTools.readShort(getByteAt(0).toInt(), getByteAt(1).toInt())

        val sheet = this.sheet
        sheet?.protectionManager?.addRecord(this)
    }

    companion object {
        private val serialVersionUID = 380140909635538745L

        /**
         * Hashes the given password with the Excel password verifier algorithm.
         *
         * The algorithm is specified by MS-XLS §2.2.9. It defers to ECMA-376-1
         * part 4 §3.2.29 for character encoding and MS-OFFCRYPTO §2.3.7.1 for the
         * hash function itself. That section of ECMA-376-1 also specifies the hash
         * function and is currently technically compatible with MS-OFFCRYPTO.
         * MS-XLS appears to indicate that MS-OFFCRYPTO is the normative spec.
         */
        protected fun hashPassword(password: String): Short {
            var strBytes: ByteArray
            var hash: Int

            /* Encode the password string as CP1252.
         * According to ECMA-376-1 16-bit Unicode characters should be encoded
         * in the best fit Windows "ANSI" code page. Java does not have the
         * best fit code page selection algorithm built in. CP1252 is the
         * correct choice for most users; other code pages are used for
         * non-Latin scripts, primarily Asian languages. We'll just use it for
         * everything until someone complains.
         */
            try {
                strBytes = password.toByteArray(charset("windows-1252"))
            } catch (e: UnsupportedEncodingException) {
                // The JVM doesn't support CP1252, ISO-8859-1 is almost identical
                strBytes = password.toByteArray(StandardCharsets.ISO_8859_1)
            }

            // start with a hash value of zero
            hash = 0

            // iterate backwards over the password bytes starting with the last one
            for (idx in strBytes.indices.reversed()) {
                // bitwise XOR the hash with the current password byte
                hash = hash xor strBytes[idx].toInt()

                // rotate the 15 lowest-order bits (mask 0x7FFF) left by one place
                hash = hash shl 1 and 0x7FFF or (hash.ushr(14) and 0x0001)
            }

            // bitwise XOR the hash with the length of the password
            hash = hash xor (strBytes.size and 0xFFFF)

            // ECMA-376 specifies this as (0x8000 | ('N' << 8) | 'K'), which
            // always evaluates to 0xCE4B. MS-OFFCRYPTO uses 0xCE4B directly.
            hash = hash xor 0xCE4B

            return hash.toShort()
        }
    }
}