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
 * The DataLabExtContents record specifies the contents of an extended data label.
 * 12 bytes- FRTHEADER
 *
 *
 * A - fSerName (1 bit): A bit that specifies whether the name of the series is displayed in the extended data label.
 * B - fCatName (1 bit): A bit that specifies whether the category (3) name, or the horizontal value on bubble or scatter chart groups, is displayed in the extended data label. MUST be a value from the following table:
 * 0		Neither of the data values are displayed in the extended data label.
 * 1		If bubble or scatter chart group, the horizontal value is displayed in the extended data label. Otherwise, the category (3) name is displayed in the extended data label.
 * C - fValue (1 bit): A bit that specifies whether the data value, or the vertical value on bubble or scatter chart groups, is displayed in the extended data label. MUST be a value from the following table:
 * 0		Neither of the data values are displayed in the data label.
 * 1		If bubble or scatter chart group, the vertical value is displayed in the extended data label. Otherwise, the data value is displayed in the extended data label.
 * D - fPercent (1 bit): A bit that specifies whether the value of the corresponding data point, represented as a percentage of the sum of the values of the series the data label is associated with, is displayed in the extended data label.
 * MUST equal 0 if the chart group type of the corresponding chart group, series, or data point is not a bar of pie, doughnut, pie, or pie of pie chart group.
 * E - fBubSizes (1 bit): A bit that specifies whether the bubble size is displayed in the data label.
 * MUST equal 0 if the chart group type of the corresponding chart group, series, or data point is not a bubble chart group.
 * reserved (11 bits): MUST be zero, and MUST be ignored.
 * rgchSep (variable): A case-sensitive XLUnicodeStringMin2 structure that specifies the string that is inserted between every data value to form the extended data label.
 * For example, if fCatName and fValue are set to 1, the labels will look like "Category Name<value ofrgchSep>Data Value".
 * The length of the string is contained in the cch field of the XLUnicodeStringMin2 structure.
</value> */
class DataLabExtContents : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fSerName: Boolean = false
    private var fCatName: Boolean = false
    private var fValue: Boolean = false
    private var fPercent: Boolean = false
    private var fBubSizes: Boolean = false

    private val PROTOTYPE_BYTES = byteArrayOf(107, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)

    /**
     * return extended data label options as an int
     * <br></br>SHOWVALUE= 0x1;
     * <br></br>SHOWVALUEPERCENT= 0x2;
     * <br></br>SHOWCATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>SHOWCATEGORYLABEL= 0x10;
     * <br></br>SHOWBUBBLELABEL= 0x20;
     * <br></br>SHOWSERIESLABEL= 0x40;
     *
     * @return a combination of data label options above or 0 if none
     */
    val typeInt: Int
        get() {
            var grbit: Short = 0
            if (fValue)
                grbit = ByteTools.updateGrBit(grbit, true, 0)
            if (fPercent)
                grbit = ByteTools.updateGrBit(grbit, true, 1)
            if (fCatName)
                grbit = ByteTools.updateGrBit(grbit, true, 4)
            if (fBubSizes)
                grbit = ByteTools.updateGrBit(grbit, true, 5)
            if (fSerName)
                grbit = ByteTools.updateGrBit(grbit, true, 6)
            return grbit.toInt()
        }

    override fun init() {
        super.init()
        val data = this.data
        grbit = ByteTools.readShort(data!![12].toInt(), data[13].toInt())
        fSerName = grbit and 0x1 == 0x1
        fCatName = grbit and 0x2 == 0x2
        fValue = grbit and 0x4 == 0x4
        fPercent = grbit and 0x8 == 0x8
        fBubSizes = grbit and 0x10 == 0x10
        if (data.size > 14) { // seperator value
            /* not following documentation:  TODO: figure out
 			int cch= data[14];
			byte[] sepbytes= this.getBytesAt(16, data.length-16);
			try {
				String sep= new String(sepbytes, UNICODEENCODING);
			} catch(Exception e) {

			}
*/
        }


    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1228364285066204304L

        val prototype: XLSRecord?
            get() {
                val dlc = DataLabExtContents()
                dlc.opcode = XLSConstants.DATALABEXTCONTENTS
                dlc.data = dlc.PROTOTYPE_BYTES
                dlc.init()
                return dlc
            }
    }
}
