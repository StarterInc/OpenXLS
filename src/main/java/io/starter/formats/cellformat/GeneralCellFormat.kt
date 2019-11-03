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
import io.starter.OpenXLS.ExcelTools

import java.text.FieldPosition
import java.text.Format
import java.text.ParsePosition

class GeneralCellFormat// make the constructor package-private
internal constructor() : Format(), CellFormat {

    override fun format(obj: Any, buffer: StringBuffer,
                        pos: FieldPosition): StringBuffer {
        var obj = obj
        // try to parse strings as numbers
        if (obj is String)
            try {
                obj = java.lang.Double.valueOf(obj)
            } catch (ex: NumberFormatException) {
                // this is OK, it just wasn't a number
            }

        if (obj is Number) {
            val num = obj
            return if (num.toLong().toDouble() == num.toDouble()) {
                // it's an integer
                buffer.append(num.toLong())
            } else {
                // it's floating-point
                buffer.append(
                        ExcelTools.getNumberAsString(num.toDouble()))
            }
        }

        return buffer.append(obj.toString())
    }

    override fun parseObject(source: String, pos: ParsePosition): Any {
        throw UnsupportedOperationException()
    }

    override fun format(cell: Cell): String {
        return this.format(cell.`val`)
    }

    fun format(`val`: String): String {
        return this.format(`val`)
    }

    companion object {
        private val serialVersionUID = -3530672760714160988L
    }
}
