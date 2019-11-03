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

/**
 * Coordinates the various records involved in protection.
 */
abstract class ProtectionManager {
    protected var protect: Protect? = null
    protected var enhancedProtection: FeatHeadr? = null
    protected var password: Password? = null

    /**
     * Returns whether the entity is protected.
     */
    /**
     * Sets whether the sheet is protected.
     */
    var protected: Boolean
        get() = protect != null && protect!!.isLocked
        set(value) {
            protect!!.isLocked = value
        }

    /**
     * Adds a protection-related record to be managed.
     * This method will be called automatically by the `init()`
     * method of the record where appropriate. It should probably never be
     * called anywhere else.
     *
     * @param record the record to be managed
     */
    open fun addRecord(record: BiffRec) {
        if (record is Protect)
            protect = record
        else if (record is Password)
            password = record
    }

    /**
     * Returns the entity's password verifier.
     *
     * @return the password verifier as four upper-case hexadecimal digits
     * or "0000" if no password is set on the sheet
     */
    fun getPassword(): String {
        return if (password == null) "0000" else password!!.passwordHashString
    }

    /**
     * Sets or removes the entity's protection password.
     *
     * @param pass the string password to set or null to remove the password
     */
    open fun setPassword(pass: String) {
        password!!.setPassword(pass)
    }

    /**
     * Sets or removes the entity's protection password.
     *
     * @param pass the pre-hashed string password to set
     * or null to remove the password
     */
    open fun setPasswordHashed(pass: String) {
        password!!.setHashedPassword(pass)
    }

    /**
     * Checks whether the given password matches the protection password.
     *
     * @param guess the password to be checked against the stored hash
     * @return whether the given password matches the stored hash
     */
    fun checkPassword(guess: String?): Boolean {
        return if (password == null) guess == null || guess == "" else password!!.validatePassword(guess)
    }

    /**
     * clear out object references in prep for closing workbook
     */
    open fun close() {
        if (enhancedProtection != null)
            enhancedProtection!!.close()
        if (password != null)
            password!!.close()
        if (protect != null)
            protect!!.close()
    }
}
