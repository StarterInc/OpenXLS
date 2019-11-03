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

import java.io.UnsupportedEncodingException

/**
 * SXDI 0xC5
 *
 *
 * The SXDI record specifies a data item for a PivotTable view.
 *
 *
 * isxvdData (2 bytes): A signed integer that specifies a pivot field index as specified in Pivot Fields.
 *
 *
 * If the PivotTable view is a non-OLAP PivotTable view, the values in the source data associated with the associated cache field of the referenced pivot field are aggregated as specified in this record.
 *
 *
 * If the PivotTable view is an OLAP PivotTable view, the associated pivot hierarchy of the referenced pivot field specifies the OLAPmeasure for this data item and the iiftab field is ignored. See Association of Pivot Hierarchies and Pivot Fields and Cache Fields to determine the associated pivot hierarchy.
 *
 *
 * MUST be greater than or equal to zero and less than the value of the cDim field of the preceding SxView record.
 *
 *
 * The value of the sxaxis.sxaxisData field of the Sxvd record of the referenced pivot field MUST be 1.
 *
 *
 * iiftab (2 bytes): A signed integer that specifies the aggregation function.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0x0000		    Sum of values
 * 0x0001		    Count of values
 * 0x0002		    Average of values
 * 0x0003		    Max of values
 * 0x0004		    Min of values
 * 0x0005		    Product of values
 * 0x0006		    Count of numbers
 * 0x0007		    Statistical standard deviation (sample)
 * 0x0008		    Statistical standard deviation (population)
 * 0x0009		    Statistical variance (sample)
 * 0x000A			Statistical variance (population)
 *
 *
 * df (2 bytes): A signed integer that specifies the calculation used to display the value of this data item.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0x0000		    The data item value is displayed.
 * 0x0001		    Display as the difference between this data item value and the value of the pivot item specified by isxvi.
 * 0x0002		    Display as a percentage of the value of the pivot item specified by isxvi.
 * 0x0003		    Display as a percentage difference from the value of the pivot item specified by isxvi.
 * 0x0004		    Display as the running total for successive pivot items in the pivot field specified by isxvd.
 * 0x0005		    Display as a percentage of the total for the row containing this data item.
 * 0x0006		    Display as a percentage of the total for the column containing this data item.
 * 0x0007		    Display as a percentage of the grand total of the data item.
 * 0x0008		    Calculate the value to display using the following formula:
 * ((this data item value) * (grand total of grand totals)) / ((row grand total) * (column grand total))
 *
 *
 * isxvd (2 bytes): A signed integer that specifies a pivot field index as specified in Pivot Fields.
 * The referenced pivot field is used in calculations as specified by the df field.
 * If df is 0x0001, 0x0002, 0x0003, or 0x0004 then the value of isxvd MUST be greater than or equal to zero and
 * less than the value of the cDim field in the preceding SxView record.
 * Otherwise, the value of isxvd is undefined and MUST be ignored.
 *
 *
 * isxvi (2 bytes): A signed integer that specifies the pivot item used by df.
 *
 *
 * If df is 0x0001, 0x0002, or 0x0003 then the value of this field MUST be a value from the following table:
 * Value		    Meaning
 * 0 to 0x7EFE	    A pivot item index, as specified by Pivot Items, that specifies a pivot item in the pivot field specified by isxvd.
 * MUST be less than the cItm field of the Sxvd record of the pivot field specified by isxvd.
 * 0x7FFB		    The previous pivot item in the pivot field specified by isxvd.
 * 0x7FFC		    The next pivot item in the pivot field specified by isxvd.
 * Otherwise, the value is undefined and MUST be ignored.
 *
 *
 * ifmt (2 bytes): An IFmt structure that specifies the number format for this item.
 *
 *
 * cchName (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stName field.
 * If the value is 0xFFFF then stName does not exist. Otherwise, the value MUST be greater than zero and less than or equal to 0x00FF.
 *
 *
 * MUST NOT be 0xFFFF when the PivotCache functionality level is less than 3, or for non-OLAP PivotTable view .
 *
 *
 * stName (variable): An XLUnicodeStringNoCch structure that specifies the name of this data item.
 * A value that is not NULL specifies that this string is used to override the name in the corresponding cache field.
 *
 *
 * MUST NOT exist if cchName is 0xFFFF. Otherwise, MUST exist and the length MUST equal cchName.
 *
 *
 * If this string is not NULL and the PivotTable view is a non-OLAP PivotTable view, this field MUST be unique within all SXDI records in this PivotTable view.
 */
class SxDI : XLSRecord(), XLSConstants {
    /**
     * serialVersionUID
     */
    /**
     * returns the pivot field index for this data item;
     * <br></br>the values in the source data associated with the associated cache field of the referenced pivot field are aggregated as specified in this record.
     *
     * @return
     */
    var pivotFieldIndex: Short = 0
        internal set
    /**
     * specifies a pivot field index used in calculations as specified by the Display Calculation function
     *
     * @return
     */
    var calculationPivotFieldIndex: Short = 0
        internal set
    internal var iiftab: Short = 0
    internal var df: Short = 0
    /**
     * specifies a pivot item index used in calculations as specified by the Display Calculation function
     *
     * @return
     */
    var calculationPivotItemIndex: Short = 0
        internal set
    internal var cchName: Short = 0
    /**
     * returns the index to the number format pattern for this data field
     *
     * @return
     */
    var numberFormat: Short = 0
        internal set
    internal var name: String? = null

    /**
     * returns the aggregation function for this data field
     *
     * @return
     * @see AGGREGATIONFUNCTIONS
     */
    /**
     * sets the aggregation function for this data field
     *
     * @see AGGREGATIONFUNCTIONS
     */
    var aggregationFunction: Int
        get() = iiftab.toInt()
        set(af) {
            iiftab = af.toShort()
            val b = ByteTools.shortToLEBytes(iiftab)
            this.getData()[2] = b[0]
            this.getData()[3] = b[1]
        }

    /**
     * A signed integer that specifies the calculation used to display the value of this data item.
     *
     * @return
     * @see DISPLAYTYPES
     */
    /**
     * sets the Display Calculation for the Data item:  A signed integer that specifies the calculation used to display the value of this data item.
     *
     * @see DISPLAYTYPES
     */
    var displayCalculation: Int
        get() = df.toInt()
        set(dc) {
            df = dc.toShort()
            val b = ByteTools.shortToLEBytes(df)
            this.getData()[4] = b[0]
            this.getData()[5] = b[1]
        }


    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, /* isxvdData */
            0, 0, /* iiftab */
            0, 0, /* df */
            0, 0, /* isxvd */
            0, 0, /* isxvi */
            0, 0, /* ifmt */
            -1, -1)/* cchname */

    /**
     * enum possible aggregation functions for Data axis fields
     * <br></br>Sum, Counta, Average, Max, Min, Product, Count, StdDev, StdDevP, Var, VarP
     */
    enum class AGGREGATIONFUNCTIONS private constructor(private val agf: String) {
        Sum("sum"),
        Count("count"),
        Average("average"),
        Max("max"),
        Min("min"),
        Product("product"),
        CountNums("countnums"),
        StdDev("stdDev"),
        StdDevP("stdDevp"),
        Var("var"),
        VarP("varp");


        companion object {

            operator fun get(s: String): Int {
                for (c in values()) {
                    if (c.agf == s)
                        return c.ordinal
                }
                return 0
            }
        }
    }

    /**
     * enum possible display types for Data axis fields
     *  * value	-- The data item value is displayed.
     *  * difference --Display as the difference between this data item value and the value of the pivot item
     *  * percentageValue -- Display as a percentage of the value of the pivot item
     *  * percentageDifference	-- Display as a percentage difference from the value of the pivot item
     *  * runningTotal	-- Display as the running total for successive pivot items in the pivot field
     *  * percentageTotalRow	-- Display as a percentage of the total for the row containing this data item.
     *  * percentageTotalCol -- Display as a percentage of the total for the column containing this data item.
     *  * grandTotal -- Display as a percentage of the grand total of the data item.
     *  * calculated 	--Calculate the value to display using the following formula:
     * <blockquote>((this data item value) * (grand total of grand totals)) / ((row grand total) * (column grand total))</blockquote>
     */
    enum class DISPLAYTYPES {
        value, difference, percentageValue, percentageDifference, runningTotal, percentageTotalRow, percentageTotalCol, grandTotal, calculated
    }

    override fun init() {
        super.init()
        pivotFieldIndex = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())    // pivot field index
        iiftab = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())    // aggregation function -- see  aggregationfunctions
        df = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())        // display calculation
        calculationPivotFieldIndex = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())    // specifies a pivot field index used in calculations as specified by the df field.
        calculationPivotItemIndex = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())    // A signed integer that specifies the pivot item used by df.
        numberFormat = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())    // number format index
        cchName = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        if (cchName.toInt() != -1) {
            val encoding = this.getByteAt(14)
            val tmp = this.getBytesAt(15, cchName * (encoding + 1))
            try {
                if (encoding.toInt() == 0)
                    name = String(tmp!!, XLSConstants.DEFAULTENCODING)
                else
                    name = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding PivotTable caption name in Sxvd: $e")
            }

        }
        if (DEBUGLEVEL > 3)
            Logger.logInfo("SXDI - isxvdData:$pivotFieldIndex iiftab:$iiftab df:$df isxvd:$calculationPivotFieldIndex isxvi:$calculationPivotItemIndex ifmt:$numberFormat name:$name")
    }

    /**
     * sets the pivot field index for this data item;
     * <br></br>the values in the source data associated with the associated cache field of the referenced pivot field are aggregated as specified in this record.
     *
     * @return
     */
    fun setPivotFieldIndex(fi: Int) {
        pivotFieldIndex = fi.toShort()
        val b = ByteTools.shortToLEBytes(pivotFieldIndex)
        this.getData()[0] = b[0]
        this.getData()[1] = b[1]
    }

    /**
     * adds the data field to the DATA axis and it's aggregation function ("sum" is default)
     *
     * @param fieldIndex
     * @param aggregationFunction
     * @param name
     */
    fun addDataField(fieldIndex: Int, aggregationFunction: String, name: String) {
        setPivotFieldIndex(fieldIndex)
        aggregationFunction = AGGREGATIONFUNCTIONS[aggregationFunction]
        setName(name)
    }

    /**
     * specifies a pivot field index used in calculations as specified by the Display Calculation function
     *
     * @return
     */
    fun setCalculationPivotFieldIndex(ci: Int) {
        calculationPivotFieldIndex = ci.toShort()
        val b = ByteTools.shortToLEBytes(calculationPivotFieldIndex)
        this.getData()[6] = b[0]
        this.getData()[7] = b[1]
    }

    /**
     * specifies a pivot item index used in calculations as specified by the Display Calculation function
     *
     * @return
     */
    fun setCalculationPivotItemIndex(ci: Int) {
        calculationPivotItemIndex = ci.toShort()
        val b = ByteTools.shortToLEBytes(calculationPivotItemIndex)
        this.getData()[8] = b[0]
        this.getData()[9] = b[1]
    }

    /**
     * sets the index to the number format pattern for this data field
     */
    fun setNumberFormat(i: Int) {
        numberFormat = i.toShort()
        val b = ByteTools.shortToLEBytes(numberFormat)
        this.getData()[10] = b[0]
        this.getData()[11] = b[1]
    }

    /**
     * returns the name of this data item; if not null, overrides the name in the corresponding cache field
     *
     * @return
     */
    fun getName(): String? {
        return name
    }

    /**
     * sets the name of this data item; if not null, overrides the name in the corresponding cache field
     * for this pivot item
     */
    fun setName(name: String?) {
        this.name = name
        var data = ByteArray(14)
        System.arraycopy(this.getData()!!, 0, data, 0, 13)
        if (name != null) {
            var strbytes: ByteArray? = null
            try {
                strbytes = this.name!!.toByteArray(charset(XLSConstants.DEFAULTENCODING))
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding pivot table name in SXVI: $e")
            }

            //update the lengths:
            cchName = strbytes!!.size.toShort()
            val nm = ByteTools.shortToLEBytes(cchName)
            data[12] = nm[0]
            data[13] = nm[1]

            // now append variable-length string data
            val newrgch = ByteArray(cchName + 1)    // account for encoding bytes
            System.arraycopy(strbytes, 0, newrgch, 1, cchName.toInt())

            data = ByteTools.append(newrgch, data)
        } else {
            data[12] = -1
            data[13] = -1
        }
        this.setData(data)
    }

    companion object {
        private val serialVersionUID = 2639291289806138985L

        val prototype: XLSRecord?
            get() {
                val di = SxDI()
                di.opcode = XLSConstants.SXDI
                di.setData(di.PROTOTYPE_BYTES)
                di.init()
                return di
            }
    }
}
