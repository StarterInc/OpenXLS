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

import java.io.Serializable

/**
 * Coordinates the various records involved in sheet-level protection.
 */
class SheetProtectionManager/*
    The PROTECT record in the Worksheet Protection Block indicates that the sheet is protected.
    There may follow a SCENPROTECT record or/and an OBJECTPROTECT record.
    The optional PASSWORD record contains the hash value of the password used to protect the sheet
    In BIFF8, there may occur additional records following the cell records in the Sheet Substream
    Sheet protection with password does not cause to switch on read/write file protection.
    Therefore the file will not be encrypted.
    Structure of the Worksheet Protection Block, BIFF5-BIFF8:
    ○ PROTECT Worksheet contents: 1 = protected
    ○ OBJECTPROTECT Embedded objects: 1 = protected (if not present, objects are not protected)
    ○ SCENPROTECT Scenarios: 1 = protected (if not present, not protected)
    ○ PASSWORD Hash value of the password; 0 = no password
     */
(
        /**
         * The worksheet whose protection state this instance manages.
         */
        private var sheet: Boundsheet?) : ProtectionManager(), Serializable {
    private var objprotect: ObjProtect? = null
    private var scenprotect: ScenProtect? = null

    /**
     * Sets whether the sheet is protected.
     */
    override// copy legacy values from the EnhancedProtection record
    // the Protect record's presence implies protection, so remove it
    // the ObjProtect record cannot exist if protection is disabled
    // Note that ScenProtect and EnhancedProtection can and should be
    // retained when disabling protection so that the same settings
    // will be restored if it is re-enabled.
    var protected: Boolean
        get() = super.protected
        set(value) {
            if (value) {
                if (protect == null)
                    addProtectionRecord()
                if (enhancedProtection != null) {
                    if (enhancedProtection!!.getProtectionOption(
                                    FeatHeadr.ALLOWOBJECTS.toInt())) {
                        if (objprotect == null) insertObjProtect()
                    }

                    setScenProtect(enhancedProtection!!.getProtectionOption(
                            FeatHeadr.ALLOWSCENARIOS.toInt()))
                }
            } else {
                if (protect != null) {
                    sheet!!.removeRecFromVec(protect!!)
                    protect = null
                }
                if (objprotect != null) {
                    sheet!!.removeRecFromVec(objprotect!!)
                    objprotect = null
                }
            }
        }

    /**
     * Adds a protection-related record to be managed.
     * This method will be called automatically by the `init()`
     * method of the record where appropriate. It should probably never be
     * called anywhere else.
     *
     * @param record the record to be managed
     */
    override fun addRecord(record: BiffRec) {
        if (record is ObjProtect)
            objprotect = record
        else if (record is ScenProtect)
            scenprotect = record
        else if (record is FeatHeadr)
            enhancedProtection = record
        else
            super.addRecord(record)
    }

    /**
     *
     */
    fun getProtected(option: Int): Boolean {
        if (enhancedProtection != null)
            return enhancedProtection!!.getProtectionOption(option)

        // fallback if we don't have an EnhancedProtection record
        when (option) {
            FeatHeadr.ALLOWOBJECTS -> {
                // if the sheet is protected, check for an ObjProtect
                return if (protect != null) objprotect == null else true
// defaults to allowed
            }

            FeatHeadr.ALLOWSCENARIOS -> {
                return if (scenprotect != null) scenprotect!!.isLocked else true
// defaults to allowed
            }

            // these extended settings default to allowed
            FeatHeadr.ALLOWSELLOCKEDCELLS, FeatHeadr.ALLOWSELUNLOCKEDCELLS -> return true

            // all other extended settings default to prohibited
            else -> return false
        }
    }

    /**
     *
     */
    fun setProtected(option: Int, value: Boolean) {
        // special cases for legacy records
        when (option) {
            FeatHeadr.ALLOWOBJECTS -> if (value) {
                if (protect != null)
                    insertObjProtect()
            } else {
                if (objprotect != null)
                    sheet!!.removeRecFromVec(objprotect!!)
            }

            FeatHeadr.ALLOWSCENARIOS -> if (protect != null) setScenProtect(value)
        }

        // add an enhanced protection record if one does not exist
        if (enhancedProtection == null) {
            enhancedProtection = FeatHeadr.prototype as FeatHeadr?
            val i = sheet!!.sheetRecs.size - 1    // just before EOF;
            sheet!!.insertSheetRecordAt(enhancedProtection!!, i)
        }

        // set the value in the enhanced protection record
        enhancedProtection!!.setProtectionOption(option, value)
    }

    private fun insertPassword() {
        if (password != null) return

        password = Password()
        if (protect == null)
            addProtectionRecord()
        var insertIdx = protect!!.recordIndex

        if (protect != null) insertIdx++
        if (objprotect != null) insertIdx++
        if (scenprotect != null) insertIdx++

        sheet!!.insertSheetRecordAt(password!!, insertIdx)
    }

    private fun removePassword() {
        if (password == null) return

        sheet!!.removeRecFromVec(password!!)
        password = null
    }

    /**
     * Sets or removes the sheet's protection password.
     *
     * @param pass the string password to set or null to remove the password
     */
    override fun setPassword(pass: String?) {
        if (pass != null && pass != "") {
            insertPassword()
            super.setPassword(pass)
        } else
            removePassword()
    }

    /**
     * Sets or removes the sheet's protection password.
     *
     * @param pass the pre-hashed string password to set
     * or null to remove the password
     */
    override fun setPasswordHashed(pass: String?) {
        if (pass != null) {
            insertPassword()
            super.setPasswordHashed(pass)
        } else
            removePassword()
    }

    private fun setScenProtect(value: Boolean) {
        if (scenprotect == null) {
            scenprotect = ScenProtect()
            sheet!!.insertSheetRecordAt(scenprotect!!, protect!!.recordIndex + 1)
        }
        scenprotect!!.isLocked = value
    }

    private fun insertObjProtect() {
        if (objprotect == null) {
            objprotect = ObjProtect()
            objprotect!!.isLocked = true
            sheet!!.insertSheetRecordAt(objprotect!!, (if (scenprotect != null)
                scenprotect
            else
                protect)
                    .recordIndex + 1)
        }
    }

    /**
     * create and add new protection record to sheet records
     */
    private fun addProtectionRecord() {
        protect = Protect()
        // baseRecordIndex= 18
        // no! is: PROTECTIONBLOCK, [DEFCOLW], [COLINFO], [SORT], DIMENSIONS
        var i = 0
        while (++i < sheet!!.sheetRecs.size) {
            val opc = (sheet!!.sheetRecs[i] as BiffRec).opcode.toInt()
            if (opc == XLSConstants.PROTECT.toInt()) {
                protect = sheet!!.sheetRecs[i] as Protect
                i++
                if ((sheet!!.sheetRecs[i] as BiffRec).opcode == XLSConstants.SCENPROTECT) {
                    scenprotect = sheet!!.sheetRecs[i] as ScenProtect
                    i++
                }
                if ((sheet!!.sheetRecs[i] as BiffRec).opcode == XLSConstants.OBJPROTECT) {
                    objprotect = sheet!!.sheetRecs[i] as ObjProtect
                }
                break
            }
            if (opc == XLSConstants.DEFCOLWIDTH.toInt() || opc == XLSConstants.COLINFO.toInt() || opc == XLSConstants.DIMENSIONS.toInt()) {
                break
            }
        }
        sheet!!.insertSheetRecordAt(protect!!, i)
        protect!!.isLocked = true
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        super.close()
        sheet = null
        if (objprotect != null)
            objprotect!!.close()
        if (scenprotect != null)
            scenprotect!!.close()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -7450088022236591508L
    }
}
