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
 * **Scl: Sheet Zoom (A0h)**<br></br>
 *
 *
 * Scl stores the zoom magnification for the sheet
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       num             2       = Numerator of the view magnification fraction (num)
 * 6		denum			2		= Denumerator of the view magnification fraction (den)
</pre> *
 */

class Scl
/**
 * default constructor
 */
internal constructor() : io.starter.formats.XLS.XLSRecord() {
    //	int num = 100; 20081231 KSC: default val is 1, making the calc (num/denum)*100
    internal var num = 1
    internal var denum = 1

    /**
     * gets the zoom as a percentage for this sheet
     *
     * @return
     */
    /**
     * sets the zoom as a percentage for this sheet
     *
     * @param b
     */
    /* 20081231 KSC:  appears that zooming is such that 1/1=100%
        // set our scale to 1000
        denum = 1000;
        byte[] denmbd= ByteTools.shortToLEBytes((short) denum);
        System.arraycopy(denmbd, 0, data, 2, 2);

        // take something like .2345 and come up with 24 & 100
        float nx = b * denum; // get denum
        // get the num
        num = (int)nx;

        if((denum % b)>0){
        	if(b>999) // only 2 precision places for zoom... out a warn
        		Logger.logWarn("Cannot set zoom to : " +b + " rounding to nearest valid zoom setting.");
        }
*/// 20081231 KSC: Convert double to fraction and set num/denum to results
    var zoom: Float
        get() = num.toFloat() / denum.toFloat()
        set(b) {
            val data = this.getData()
            val n = gcd((b * 100).toInt(), 100)
            num = n[0]
            denum = n[1]
            var nmbd = ByteTools.shortToLEBytes(num.toShort())
            System.arraycopy(nmbd, 0, data!!, 0, 2)
            nmbd = ByteTools.shortToLEBytes(denum.toShort())
            System.arraycopy(nmbd, 0, data, 2, 2)

            this.setData(data)
        }

    init {
        val bs = ByteArray(4)
        bs[0] = 1
        bs[1] = 0
        bs[2] = 1
        bs[3] = 0
        opcode = XLSConstants.SCL
        length = 4.toShort()
        if (DEBUGLEVEL > XLSConstants.DEBUG_LOW)
            Logger.logInfo("Scl.init()" + this.offset)
        this.setData(bs)
        this.originalsize = 4
    }

    override fun init() {
        super.init()
        num = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        denum = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
        if (DEBUGLEVEL > XLSConstants.DEBUG_LOW)
            Logger.logInfo("Scl.init() sheet zoom:$zoom")
    }


    private fun gcd(numerator: Int, denominator: Int): IntArray {
        val highest: Int
        var n = 1
        var d = 1

        if (denominator > numerator)
            highest = denominator
        else
            highest = numerator

        for (x in highest downTo 1) {
            if (denominator % x == 0 && numerator % x == 0) {
                n = numerator / x
                d = denominator / x
                break
            }
        }
        return intArrayOf(n, d)
    }

    companion object {
        /**
         *
         */
        private val serialVersionUID = -4595833226859365049L
    }
}