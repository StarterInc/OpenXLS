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
 * **YMULT: Y Multiplier (857h)**
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 * This record describes the axis multiplier feature which scales the axis values
 * displayed by the axis tick labels. For instance, an axis multiplier value of
 * "millions" would cause the axis tick labels to show the axis value divided by one
 * million (e.g., the tick label for an axis value of 20,000,000 would show "20".)
 * This record is a "parent" record and is immediately followed by a set of records
 * surrounded by rtStartObject and rtEndObject which describes the axis multiplier label.
 *
 *
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0857h
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		axmid		2		Axis multiplier ID, one of the following values:
 * -1 = multiplier value is stored in numLabelMultiplier
 * 0 = no multiplier (same as 1.0)
 * 1 = Hundreds, 10 2nd
 * 2 = Thousands, 10 3rd
 * 3 = Ten Thousands, 10 4th
 * 4 = Hundred Thousands, 10 5th
 * 5 = Millions, 10 6th
 * 6 = Ten Millions, 10 7th
 * 7 = Hundred Millions, 10 8th
 * 8 = billion
 * 9 = trillion
 * 16		numLabelMultiplier	4	Numeric value
 * 18		grbit		2		Option flags for y axis multiplier (see description below)*
 *
 *
 * The grbit field contains the following category axis label option flags:
 * Bits	Mask	Flag Name	Contents
 * 0		0001h	fEnabled	=1 if the multiplier is enabled =0 otherwise
 * 1		0002h	fAutoShowMultiplier	=1 if the multiplier label is shown =0 otherwise
 * 15-2	FFFCh	(unused)	Reserved; must be zero
 */
class YMult : GenericChartObject(), ChartObject {
    /**
     * returns the Axis multiplier ID, one of the following values:
     *  * -1 = multiplier value is stored in numLabelMultiplier
     *  * 0 = no multiplier (same as 1.0)
     *  * 1 = Hundreds, 10 2nd
     *  * 2 = Thousands, 10 3rd
     *  * 3 = Ten Thousands, 10 4th
     *  * 4 = Hundred Thousands, 10 5th
     *  * 5 = Millions, 10 6th
     *  * 6 = Ten Millions, 10 7th
     *  * 7 = Hundred Millions, 10 8th
     *  * 8 = Thousand Millions, 10 9th
     *  * 9 = Billions, 10 12th
     *
     * @return
     */
    var axMultiplierId: Short = 0
        internal set
    internal var grbit: Short = 0
    internal var numLabelMultiplier: Double = 0.toDouble()

    // TODO: Prototype Bytes
    private val PROTOTYPE_BYTES = byteArrayOf()

    //?
    val axMultiplierIdAsString: String?
        get() {
            when (axMultiplierId) {
                -1 -> return null
                0 -> return null
                1 -> return "hundreds"
                2 -> return "thousands"
                3 -> return "tenThousands"
                4 -> return "hundredThousands"
                5 -> return "millions"
                6 -> return "tenMillions"
                7 -> return "hundredMillions"
                8 -> return "billions"
                9 -> return "trillions"
            }
            return null
        }

    // dispUnits -> builtInUnit (val:  billions,
    // custom
    var customMultiplier: Double
        get() = numLabelMultiplier
        set(m) {
            numLabelMultiplier = m
            axMultiplierId = -1
            var b = ByteTools.shortToLEBytes(axMultiplierId)
            this.data[4] = b[0]
            this.data[5] = b[1]
            b = ByteTools.doubleToLEByteArray(numLabelMultiplier)
            System.arraycopy(b, 0, this.data!!, 6, 8)
        }

    override fun init() {
        super.init()
        axMultiplierId = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        numLabelMultiplier = ByteTools.eightBytetoLEDouble(this.getBytesAt(6, 8)!!)
        grbit = ByteTools.readShort(this.getByteAt(14).toInt(), this.getByteAt(15).toInt())
    }

    /**
     * Sets Axis multiplier ID, one of the following values:
     *  * -1 = multiplier value is stored in numLabelMultiplier
     *  * 0 = no multiplier (same as 1.0)
     *  * 1 = Hundreds, 10 2nd
     *  * 2 = Thousands, 10 3rd
     *  * 3 = Ten Thousands, 10 4th
     *  * 4 = Hundred Thousands, 10 5th
     *  * 5 = Millions, 10 6th
     *  * 6 = Ten Millions, 10 7th
     *  * 7 = Hundred Millions, 10 8th
     *  * 8 = Thousand Millions, 10 9th
     *  * 9 = Billions, 10 12th
     *
     * @param m
     */
    fun setAxMultiplierId(m: Int) {
        if (!(m > -2 && m < 10))
            return     // report error?
        axMultiplierId = m.toShort()
        val b = ByteTools.shortToLEBytes(axMultiplierId)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    /**
     * Sets Axis multiplier ID via OOXML String value:
     *  * hundreds			Hundreds
     *  * thousands			Thousands
     *  * tenThousands		Ten Thousands
     *  * hundredThousands	Hundred Thousands
     *  * millions			Millions
     *  * tenMillions			Ten Millions
     *  * hundredMillions		Hundred Millions
     *  * billions			Billions
     *  * trillions			Trillions
     */
    fun setAxMultiplierId(m: String) {
        if (m.equals("hundreds", ignoreCase = true))
            axMultiplierId = 1
        else if (m.equals("thousands", ignoreCase = true))
            axMultiplierId = 2
        else if (m.equals("tenThousands", ignoreCase = true))
            axMultiplierId = 3
        else if (m.equals("hundredThousands", ignoreCase = true))
            axMultiplierId = 4
        else if (m.equals("millions", ignoreCase = true))
            axMultiplierId = 5
        else if (m.equals("tenMillions", ignoreCase = true))
            axMultiplierId = 6
        else if (m.equals("hundredMillions", ignoreCase = true))
            axMultiplierId = 7
        else if (m.equals("billions", ignoreCase = true))
            axMultiplierId = 8
        else if (m.equals("trillions", ignoreCase = true))
            axMultiplierId = 9
        else
        // default
            axMultiplierId = 0
        val b = ByteTools.shortToLEBytes(axMultiplierId)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    companion object {
        /**
         * serialVersionUID
         */

        private val serialVersionUID = -6166267220292885486L

        val prototype: XLSRecord?
            get() {
                val ym = YMult()
                ym.opcode = XLSConstants.YMULT
                ym.data = ym.PROTOTYPE_BYTES
                ym.init()
                return ym
            }
    }
}
