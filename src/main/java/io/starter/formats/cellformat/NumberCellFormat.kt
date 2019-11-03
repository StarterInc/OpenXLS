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
package io.starter.formats.cellformat

import io.starter.OpenXLS.Cell

import java.text.FieldPosition
import java.text.NumberFormat
import java.text.ParsePosition

class NumberCellFormat internal constructor(private val positive: String, private val negative: String, private val zero: String, private val string: String) : NumberFormat(), CellFormat {

    override fun format(input: Any, buffer: StringBuffer,
                        pos: FieldPosition): StringBuffer {
        var input = input
        if (input is String) {
            // hack to make useless @ pattern work
            if ("%s" == this.positive) {
                return buffer.append(input)
            }
            try {
                val d = Double(input.toString())
                input = d
            } catch (e: NumberFormatException) {
                return buffer.append(
                        String.format(this.string, input))
            }

        }
        if (input is Number) {
            val format: String
            var value = input.toDouble()

            if (value > 0)
                format = this.positive
            else if (value < 0) {
                format = this.negative
                value = Math.abs(value)
            } else
                format = this.zero

            // hack to make percentage formats work
            if (format.contains("%%")) {
                value *= 100.0
            }

            // hack to make useless @ pattern work
            return if ("%s" == format) {
                buffer.append(input)
            } else buffer.append(
                    String.format(format, java.lang.Double.valueOf(value)))

        } else {
            throw IllegalArgumentException("unsupported input type")
        }
    }

    override fun format(number: Double, buffer: StringBuffer,
                        pos: FieldPosition): StringBuffer {
        return buffer.append(this.format(java.lang.Double.valueOf(number)))
    }

    override fun format(number: Long, buffer: StringBuffer,
                        pos: FieldPosition): StringBuffer {
        return buffer.append(this.format(java.lang.Long.valueOf(number)))
    }

    override fun parse(source: String, parsePosition: ParsePosition): Number {
        throw UnsupportedOperationException()
    }

    override fun format(cell: Cell): String {
        return this.format(cell.`val`)
    }

    companion object {
        private val serialVersionUID = -7191923168789058338L
    }

}
