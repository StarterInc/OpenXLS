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
package io.starter.OpenXLS

/**
 * Constants used in the generation of JSON output.
 */
internal interface JSONConstants {
    companion object {
        val JSON_CELL_VALUE = "v"
        val JSON_CELL_FORMATTED_VALUE = "fv"
        val JSON_CELL_FORMULA = "fm"
        val JSON_CELL = "Cell"
        val JSON_CELLS = "cs"
        val JSON_DATETIME = "DateTime"
        val JSON_DATEVALUE = "DateValue"
        val JSON_DOUBLE = "Double"
        val JSON_RANGE = "Range"
        val JSON_DATA = "d"
        val JSON_FLOAT = "Float"
        val JSON_INTEGER = "Integer"
        val JSON_LOCATION = "loc"
        val JSON_ROW = "Row"
        val JSON_ROW_BORDER_TOP = "BdrT"
        val JSON_ROW_BORDER_BOTTOM = "BdrB"
        val JSON_HEIGHT = "h"
        val JSON_STRING = "String"
        val JSON_STYLEID = "sid"
        val JSON_TYPE = "t"
        val JSON_FORMULA_HIDDEN = "fhd"
        val JSON_LOCKED = "lck"
        val JSON_HIDDEN = "Hidden"
        val JSON_VALIDATION_MESSAGE = "vm"
        val JSON_MERGEACROSS = "MergeAcross"
        val JSON_MERGEDOWN = "MergeDown"
        val JSON_MERGEPARENT = "MergeParent"
        val JSON_MERGECHILD = "MergeChild"
        val JSON_HREF = "HRef"
        val JSON_WORD_WRAP = "wrap"
        val JSON_RED_FORMAT = "negRed"
        val JSON_TEXT_ALIGN = "txtAlign"
    }
}
