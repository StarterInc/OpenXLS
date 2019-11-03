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

package io.starter.formats.XLS.charts

import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **Tick: Tick Marks and Labels Format (0x101e)**
 *
 *
 * Offset 	Name		Size	Contents
 * 4		tktMajor	1		Type of major tick mark 0= invisible (none) I = inside of axis line 2 = outside of axis line 3 = cross axis line
 * 5		tktMinor	1		Type of minor tick mark 0= invisible (none)	I = inside of axis line 2 = outside of axis line 3 = cross axis line
 * 6		tit			1		Tick label position relative to axis line 0= invisible (none) 1 = low end of plot area 2 = high end of plot area 3 = next to axis
 * 7		wBkgMode	2		Background mode: I = transparent 2 = opaque
 * 8		rgb			4		Tick-label text color; ROB value, high byte = 0
 * 12		(reserved)	16		Reserved; must be zero
 * 28		grbit		2		Display flags
 * 30		icv			2		Index to color of tick label
 * 32		(reserved)	2		Reserved; must be zero
 *
 *
 * The grbit field contains the following option flags.
 *
 *
 * Bits	Mask		Name
 * 0		0xOl		fAutoColor		Automatic text color
 * 1		0x02		fAutoMode 		Automatic text back~
 * 4-2		0xlC		rot				0= no rotation (text appears left-to-right), 1=  text appears top-~~ are upright,
 * 2= text is rotated 90 degrees counterclockwise,  3= text is rotated
 * 5		0x20		fAutoRot		Automatic rotation
 * 7-6		0xCO		(reserved)		Reserved; must be zero
 * 7-0		0xFF		(reserved)		Reserved; must be zero
 */
class Tick : GenericChartObject(), ChartObject {
    internal var tktMajor: Byte = 0
    internal var tktMinor: Byte = 0
    internal var tit: Byte = 0
    internal var grbit: Short = 0
    /**
     * 0= no rotation (text appears left-to-right),
     * 1= Text is drawn stacked, top-to-bottom, with the letters upright.
     * 2= text is rotated 90 degrees counterclockwise,
     * 3= text is rotated at 90 degrees clockwise.
     *
     * @return
     */
    var rotation: Short = 0
        internal set

    private val PROTOTYPE_BYTES = byteArrayOf(2, 0, 3, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 35, 0, 77, 0, 0, 0)

    override fun init() {
        super.init()
        tktMajor = this.getByteAt(0)
        tktMinor = this.getByteAt(1)
        tit = this.getByteAt(2)
        grbit = ByteTools.readShort(this.getByteAt(24).toInt(), this.getByteAt(25).toInt())
        // TODO: Finish ops
        rotation = (grbit and 0x1C shr 2).toShort()
    }

    private fun updateRecord() {
        this.data[0] = tktMajor
        this.data[1] = tktMinor
        this.data[2] = tit
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[24] = b[0]
        this.data[25] = b[1]
    }

    /**
     * set generic Tick option
     * <br></br>op/val can be one of:
     * <br></br>tickLblPos			none, low, high or nextTo
     * <br></br>majorTickMark		none, in, out, cross
     * <br></br>minorTickMark		none, in, out, cross
     *
     * @param op
     * @param val
     */
    fun setOption(op: String, `val`: String) {
        if (op == "tickLblPos") {        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick
            if (`val` == "high")
                tit = 2
            else if (`val` == "low")
                tit = 1
            else if (`val` == "none")
                tit = 0
            else if (`val` == "nextTo")
                tit = 3
        } else if (op == "majorTickMark") {    // major tick marks (cross, in, none, out)
            if (`val` == "cross")
                tktMajor = 3
            else if (`val` == "in")
                tktMajor = 1
            else if (`val` == "out")
                tktMajor = 2
            else if (`val` == "none")
                tktMajor = 0
        } else if (op == "minorTickMark") {    // minor tick marks (cross, in, none, out)
            if (`val` == "cross")
                tktMinor = 3
            else if (`val` == "in")
                tktMinor = 1
            else if (`val` == "out")
                tktMinor = 2
            else if (`val` == "none")
                tktMinor = 0
        }
        updateRecord()
    }
    /*  4		tktMajor	1		Type of major tick mark 0= invisible (none) I = inside of axis line 2 = outside of axis line 3 = cross axis line
     *  5		tktMinor	1		Type of minor tick mark 0= invisible (none)	I = inside of axis line 2 = outside of axis line 3 = cross axis line
     *  6		tit			1		Tick label position relative to axis line 0= invisible (none) 1 = low end of plot area 2 = high end of plot area 3 = next to axis
     */

    /**
     * retrieve generic Value axis option as OOXML string
     * <br></br>can be one of:
     * <br></br>tickLblPos			none, low, high or nextTo
     * <br></br>majorTickMark		none, in, out, cross
     * <br></br>minorTickMark		none, in, out, cross
     *
     * @param op
     * @return
     */
    fun getOption(op: String): String? {
        if (op == "tickLblPos") {        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick
            when (tit) {
                0 -> return "none"
                1 -> return "low"
                2 -> return "high"
                3 -> return "nextTo"
            }
        } else if (op == "majorTickMark") {// major tick marks (cross, in, none, out)
            when (tktMajor) {
                0 -> return "none"
                1 -> return "in"
                2 -> return "out"
                3 -> return "cross"
            }
        } else if (op == "minorTickMark") {    // minor tick marks (cross, in, none, out)
            when (tktMinor) {
                0 -> return "none"
                1 -> return "in"
                2 -> return "out"
                3 -> return "cross"
            }
        }
        return null
    }

    /**
     * returns true if should show minor tick marks
     *
     * @return
     */
    fun showMinorTicks(): Boolean {
        return tktMinor.toInt() != 0
    }

    /**
     * returns true if should show major tick marks
     *
     * @return
     */
    fun showMajorTicks(): Boolean {
        return tktMajor.toInt() != 0
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3363212452589555220L

        // 20070723 KSC: Need to create new records
        // important when we parse options ...
        val prototype: XLSRecord?
            get() {
                val t = Tick()
                t.opcode = XLSConstants.TICK
                t.data = t.PROTOTYPE_BYTES
                t.init()
                return t
            }
    }

}
