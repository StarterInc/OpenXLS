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
import io.starter.OpenXLS.CellHandle
import io.starter.OpenXLS.DateConverter
import java.text.SimpleDateFormat

class DateCellFormat internal constructor(date: String?, private val text_format: String) : SimpleDateFormat(date), CellFormat {

    override fun format(cell: Cell?): String? { // make sure to return the empty string for blank cells
// getting the calendar coerces to double and thus gets zero
        if (cell is CellHandle && cell.isBlank
                || "" == cell?.getVal()) return ""
        if (cell?.cellType  == Cell.TYPE_STRING) {
            return String.format(text_format, cell.getVal() as String)
        }
        val date = DateConverter.getCalendarFromCell(cell)
        return this.format(date.time)
    }

    companion object {
        private const val serialVersionUID = 1896075041723437260L
    }

}