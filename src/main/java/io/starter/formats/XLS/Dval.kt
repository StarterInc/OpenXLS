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
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList

/**
 * **Dval: Data Validity Settings (01B2h)**<br></br>
 *
 *
 * This record is the list header of the Data Validity Table in the current sheet.
 *
 *
 * Offset          Name            Size                 Contents
 * -------------------------------------------------------
 * 0                 wDviFlags     2                       Option flags:
 * 2                 xLeft             4                       Horizontal position of the prompt box, if it has fixed position, in pixel
 * 6                 yTop              4                       Vertical position of the prompt box, if it has fixed position, in pixel
 * 10              inObj               4                       Object identifier of the drop down arrow object for a list box , if a list box is visible at
 * the current cursor position, FFFFFFFFH otherwise
 * 14                idvMac            4                      Number of following DV records
 *
 *
 *
 *
 * wDviFlags
 * Bit         Mask              Name            Contents
 * ------------------------------------------------------
 * 0           0001H          fWnClosed            0 = Prompt box not visible 1 = Prompt box currently visible
 * 1           0002H          fWnPinned            0 = Prompt box has fixed position 1 = Prompt box appears at cell
 * 2           0004H          fCached              1 = Cell validity data cached in following DV records
 */

class Dval : io.starter.formats.XLS.XLSRecord() {
    // primary fields
    private var grbit: Short = 0
    private var xLeft: Int = 0
    private var yTop: Int = 0
    private var inObj: Int = 0
    private var idvMac: Int = 0

    // 200906060 KSC: Mod bytes to have no x, y or dv recs following:    private static final byte[] PROTOTYPE_BYTES = {04, 00, 124, 00, 00, 00, 00, 00, 00, 00, -1, -1, -1, -1, 01, 00, 00, 00};
    //    private static final byte[] PROTOTYPE_BYTES = {04, 00, 00, 00, 00, 00, 00, 00, 00, 00, -1, -1, -1, -1, 00, 00, 00, 00};
    private val PROTOTYPE_BYTES = byteArrayOf(4, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, -1, -1, -1, 0, 0, 0, 0)

    private var dvRecs: ArrayList<*>? = null

    /**
     * Is Cell validity data cached in following DV records?
     *
     * @return
     */
    /**
     * Set cell validity data cached in following DV records
     */
    var isValidityCached: Boolean
        get() = grbit and BITMASK_F_CACHED == BITMASK_F_CACHED.toInt()
        set(cached) {
            if (cached)
                grbit = (grbit or BITMASK_F_CACHED).toShort()
            else
                grbit = (grbit xor BITMASK_F_CACHED).toShort()
            this.setGrbit()
        }

    /**
     * Get where the prompt box is located.
     * true = Prompt box is at cell; false= Prompt box in fixed position
     *
     * @return
     */
    /**
     * Set where the prompt box is located.
     * true = Prompt box is at cell; false= Prompt box in fixed position
     *
     * @return
     */
    var isPromptBoxAtCell: Boolean
        get() = grbit and BITMASK_F_WN_PINNED == BITMASK_F_WN_PINNED.toInt()
        set(location) {
            if (location)
                grbit = (grbit or BITMASK_F_WN_PINNED).toShort()
            else
                grbit = (grbit xor BITMASK_F_WN_PINNED).toShort()
            this.setGrbit()
        }


    /**
     * Get visibility of prompt box
     * true = Prompt box currently visible; false = Prompt box not visible
     *
     * @return
     */
    /**
     * Set visibility of prompt box
     * true = Prompt box currently visible; false = Prompt box not visible
     *
     * @return
     */
    var isPromptBoxVisible: Boolean
        get() = grbit and BITMASK_F_WN_CLOSED == BITMASK_F_WN_CLOSED.toInt()
        set(location) {
            if (location)
                grbit = (grbit or BITMASK_F_WN_CLOSED).toShort()
            else
                grbit = (grbit xor BITMASK_F_WN_CLOSED).toShort()
            this.setGrbit()
        }


    /**
     * Get the count of dv records following this Dval
     *
     * @return
     */
    /**
     * Set the count of following Dv records
     *
     * @param cnt = count
     */
    var followingDvCount: Int
        get() = idvMac
        set(cnt) {
            this.idvMac = cnt
            val data = this.getData()
            val b = ByteTools.cLongToLEBytes(idvMac)
            System.arraycopy(b, 0, data!!, 14, 4)
            this.setData(data)
        }

    /**
     * Object identifier of the drop down arrow object for a list box ,
     * if a list box is visible at the current cursor position, FFFFFFFFH otherwise
     *
     * @return
     */
    /**
     * Object identifier of the drop down arrow object for a list box ,
     * if a list box is visible at the current cursor position, FFFFFFFFH otherwise
     *
     * @param cnt = identifier
     */
    var objectIdentifier: Int
        get() = inObj
        set(cnt) {
            this.inObj = cnt
            val data = this.getData()
            val b = ByteTools.cLongToLEBytes(inObj)
            System.arraycopy(b, 0, data!!, 10, 4)
            this.setData(data)
        }

    /**
     * Horizontal position of the prompt box, if it has fixed position, in pixel
     *
     * @return
     */
    /**
     * Horizontal position of the prompt box, if it has fixed position, in pixel
     *
     * @param cnt = position
     */
    var horizontalPosition: Int
        get() = xLeft
        set(cnt) {
            this.xLeft = cnt
            val data = this.getData()
            val b = ByteTools.cLongToLEBytes(xLeft)
            System.arraycopy(b, 0, data!!, 2, 4)
            this.setData(data)
        }

    /**
     * Vertical position of the prompt box, if it has fixed position, in pixel
     *
     * @return
     */
    /**
     * Vertical position of the prompt box, if it has fixed position, in pixel
     *
     * @param cnt = position
     */
    var verticalPosition: Int
        get() = yTop
        set(cnt) {
            this.yTop = cnt
            val data = this.getData()
            val b = ByteTools.cLongToLEBytes(yTop)
            System.arraycopy(b, 0, data!!, 2, 4)
            this.setData(data)
        }

    /**
     * Return all dvs for this Dval
     *
     * @return
     */
    val dvs: List<*>?
        get() = dvRecs

    /**
     * generate the proper OOXML to define this Dval
     *
     * @return
     */
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (this.dvRecs!!.size > 0) {
                ooxml.append("<dataValidations count=\"" + this.dvRecs!!.size + "\"")
                if (!this.isPromptBoxVisible) ooxml.append(" disablePrompts=\"1\"")
                if (this.horizontalPosition != 0) ooxml.append(" xWindow=\"" + this.horizontalPosition + "\"")
                if (this.verticalPosition != 0) ooxml.append(" yWindow=\"" + this.verticalPosition + "\"")
                ooxml.append(">")
                for (i in this.dvRecs!!.indices) {
                    ooxml.append((dvRecs!![i] as Dv).ooxml)
                }
                ooxml.append("</dataValidations>")
            }
            return ooxml.toString()
        }


    /**
     * Standard init method, nothing new
     *
     * @see io.starter.formats.XLS.XLSRecord.init
     */
    override fun init() {
        super.init()
        var offset = 0
        dvRecs = ArrayList()
        grbit = ByteTools.readShort(this.getByteAt(offset++).toInt(), this.getByteAt(offset++).toInt())
        xLeft = ByteTools.readInt(this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++))
        yTop = ByteTools.readInt(this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++))
        inObj = ByteTools.readInt(this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++))
        idvMac = ByteTools.readInt(this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++), this.getByteAt(offset++))
    }

    /**
     * Apply the grbit to the record in the streamer
     */
    fun setGrbit() {
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(grbit)
        System.arraycopy(b, 0, data!!, 0, 2)
        this.setData(data)
    }

    /**
     * Add a new Dv record to this Dval;
     *
     * @param dv
     */
    fun addDvRec(dv: Dv) {
        dvRecs!!.add(dv)
    }

    /**
     * Add a new (ie not on parse) dv record,
     * updates the parent record with the count
     *
     * @param location = the cell/range that the dv attaches to.  Sheet name not required
     * as Dval is a sheet not book level record
     */
    fun createDvRec(location: String): Dv {
        val d = Dv.getPrototype(this.workBook) as Dv
        d.setSheet(this.sheet)
        d.setRange(location)
        this.addDvRec(d)
        this.followingDvCount = dvRecs!!.size
        return d
    }

    /**
     * Remove a dv rec from this dval.
     *
     * @param dv
     */
    fun removeDvRec(dv: Dv) {
        dvRecs!!.remove(dv)
    }

    /**
     * @param cellAddress
     * @return
     */
    fun getDv(cellAddress: String): Dv? {
        var cellAddress = cellAddress
        if (cellAddress.indexOf("!") != -1) {
            cellAddress = cellAddress.substring(cellAddress.indexOf("!"))
        }
        for (i in dvRecs!!.indices) {
            val d = dvRecs!![i] as Dv
            if (d.isInRange(cellAddress)) return d
        }
        return null
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3954586766300169606L

        // grbitlookups
        private val BITMASK_F_WN_CLOSED: Short = 0x0001
        private val BITMASK_F_WN_PINNED: Short = 0x0002
        private val BITMASK_F_CACHED: Short = 0x0004

        /**
         * Create a dval record & populate with prototype bytes
         *
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val dval = Dval()
                dval.opcode = XLSConstants.DVAL
                dval.setData(dval.PROTOTYPE_BYTES)
                dval.init()
                return dval
            }

        /**
         * OOXML Element:
         * dataValidations (Data Validations)
         * This collection expresses all data validation information for cells in a sheet which have data validation features
         * applied.
         * Data validation is used to specify constaints on the type of data that can be entered into a cell. Additional UI can
         * be provided to help the user select valid values (e.g., a dropdown control on the cell or hover text when the cell
         * is active), and to help the user understand why a particular entry was considered invalid (e.g., alerts and
         * messages).
         * Various data types can be selected, and logical operators (e.g., greater than, less than, equal to, etc) can be
         * used. Additionally, instead of specifying an explicit set of values that are valid, a cell or range reference may be
         * used.
         * An input message can be specified to help the user know what kind of value is expected, and a warning message
         * (and warning type) can be specified to alert the user when they've entered invalid data.
         *
         * parent:  worksheet
         * children:  datraValidation
         * attributes:  count (uint), xWindow (uint)-per-sheet x, yWindow (uint)- per-sheet y, disablePrompts (bool)
         *
         * TODO:  shouldnt this be in the OOXML package? why generate/treat this differently?
         */

        /**
         * create one or more Data Validation records based on OOXML input
         */
        fun parseOOXML(xpp: XmlPullParser, bs: Boundsheet): Dval? {
            val dval = bs.insertDvalRec()    // creates or retrieves Dval rec
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "dataValidations") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (n == "count") {
                                } else if (n == "disablePrompts") {
                                    dval.isPromptBoxVisible = false
                                } else if (n == "xWindow") {
                                    dval.horizontalPosition = Integer.valueOf(v).toInt()
                                } else if (n == "yWindow") {
                                    dval.verticalPosition = Integer.valueOf(v).toInt()
                                }
                            }
                        } else if (tnm == "dataValidation") {    // one or more
                            Dv.parseOOXML(xpp, bs)    // creates and adds a new Dv record to the Dval recs list
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "dataValidations") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("OOXMLELEMENT.parseOOXML: $e")
            }

            return dval
        }
    }
}
