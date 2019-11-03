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


/**
 * **USERSVIEWBEGIN: Custom View Settings (1AAh)**<br></br>
 *
 *
 * USERSVIEWBEGIN describes the settings for a custom view for the sheet
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       guid            16      GID for custom view
 * 20      iTabid          4       Tab index for the sheet (1-based)
 * 24      wScale          4       Window Zoom
 * 28      icv             4       Index to color val
 * 32      pnnSel          4       Pane number of active pane
 * 36      grbit           4       Option flags
 * 40      refTopLeft      8       Ref struct describing the visible area of top left pane
 * 48      operNum         16      array of 2 floats specifying vert/horiz pane split
 * 64      colRPane        2       first visible right pane col
 * 66      rwBPane         2       first visible bottom pane col
 *
</pre> *
 */
class Usersviewbegin @JvmOverloads internal constructor(b: ByteArray = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)) : io.starter.formats.XLS.XLSRecord() {
    // record fields
    private var tabid = -1
    private var wScale = -1
    private var icv = -1
    private var pnnSel = -1
    private var grbit = -1

    // grbit fields
    private var fDspGutsSv = false // true if outline symbols are displayed

    /**
     * Checks to see if outlines are being displayed on the worksheet
     *
     * @return fDspGutsSv
     */
    /**
     * Sets whether outlines are displayed on the worksheet.  Should be set to true
     * when using any of the 'grouping' methods.
     *
     * @param disp
     */
    var displayOutlines: Boolean
        get() = fDspGutsSv
        set(disp) {
            fDspGutsSv = disp
            updateGrbit()
        }

    init {
        setData(b)
        opcode = XLSConstants.USERSVIEWBEGIN
        length = 6.toShort()
        this.init()
    }

    // TODO: implement this class
    override fun init() {
        super.init()
        var num1 = ByteTools.readShort(this.getByteAt(16).toInt(), this.getByteAt(17).toInt())
        var num2 = ByteTools.readShort(this.getByteAt(18).toInt(), this.getByteAt(19).toInt())
        tabid = ByteTools.readInt(num2.toInt(), num1.toInt())
        num1 = ByteTools.readShort(this.getByteAt(20).toInt(), this.getByteAt(21).toInt())
        num2 = ByteTools.readShort(this.getByteAt(22).toInt(), this.getByteAt(23).toInt())
        wScale = ByteTools.readInt(num1.toInt(), num2.toInt())
        num1 = ByteTools.readShort(this.getByteAt(24).toInt(), this.getByteAt(25).toInt())
        num2 = ByteTools.readShort(this.getByteAt(26).toInt(), this.getByteAt(27).toInt())
        icv = ByteTools.readInt(num1.toInt(), num2.toInt())
        num1 = ByteTools.readShort(this.getByteAt(24).toInt(), this.getByteAt(25).toInt())
        num2 = ByteTools.readShort(this.getByteAt(26).toInt(), this.getByteAt(27).toInt())
        pnnSel = ByteTools.readInt(num1.toInt(), num2.toInt())
        num1 = ByteTools.readShort(this.getByteAt(28).toInt(), this.getByteAt(29).toInt())
        num2 = ByteTools.readShort(this.getByteAt(30).toInt(), this.getByteAt(31).toInt())
        grbit = ByteTools.readInt(num1.toInt(), num2.toInt())
        this.decodeGrbit()
        if (DEBUGLEVEL > 3) Logger.logInfo("Usersviewbegin Tab Index: $tabid")
    }

    /**
     * decodeGrbit does masking to determine the grbit settings
     */
    private fun decodeGrbit() {
        fDspGutsSv = grbit and 0x00000010 == 0x00000010


    }

    /**
     * updateGrbit looks at all the grbit variables and rebuilds the grbit
     * field based off those.
     */
    private fun updateGrbit() {
        if (fDspGutsSv) {
            grbit = 0x00000010 or grbit
        } else {
            grbit = -0x11 and grbit
        }
        // update the record bytes
        val b = this.getData()
        val grbytes = ByteTools.cLongToLEBytes(grbit)
        System.arraycopy(grbytes, 0, b!!, 28, 4)
        this.setData(b)
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1877650235927064991L
    }
    /*
      20      iTabid          4       Tab index for the sheet (1-based)
    */


}//private float operNum1 = 0;
//private float operNum2 = 0;
//private short colRPane = 0;
//private short rwBPane = 0;
//66 Long!
//byte[] RECBYTES =  {0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0};
/**
 * Constructor for a Usersviewbegin to be made on the fly.
 */// TODO: init Usersviewbegin