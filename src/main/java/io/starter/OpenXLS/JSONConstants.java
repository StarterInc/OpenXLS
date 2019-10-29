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
package io.starter.OpenXLS;

/**
 * Constants used in the generation of JSON output.
 */
interface JSONConstants {
    String JSON_CELL_VALUE = "v";
    String JSON_CELL_FORMATTED_VALUE = "fv";
    String JSON_CELL_FORMULA = "fm";
    String JSON_CELL = "Cell";
    String JSON_CELLS = "cs";
    String JSON_DATETIME = "DateTime";
    String JSON_DATEVALUE = "DateValue";
    String JSON_DOUBLE = "Double";
    String JSON_RANGE = "Range";
    String JSON_DATA = "d";
    String JSON_FLOAT = "Float";
    String JSON_INTEGER = "Integer";
    String JSON_LOCATION = "loc";
    String JSON_ROW = "Row";
    String JSON_ROW_BORDER_TOP = "BdrT";
    String JSON_ROW_BORDER_BOTTOM = "BdrB";
    String JSON_HEIGHT = "h";
    String JSON_STRING = "String";
    String JSON_STYLEID = "sid";
    String JSON_TYPE = "t";
    String JSON_FORMULA_HIDDEN = "fhd";
    String JSON_LOCKED = "lck";
    String JSON_HIDDEN = "Hidden";
    String JSON_VALIDATION_MESSAGE = "vm";
    String JSON_MERGEACROSS = "MergeAcross";
    String JSON_MERGEDOWN = "MergeDown";
    String JSON_MERGEPARENT = "MergeParent";
    String JSON_MERGECHILD = "MergeChild";
    String JSON_HREF = "HRef";
    String JSON_WORD_WRAP = "wrap";
    String JSON_RED_FORMAT = "negRed";
    String JSON_TEXT_ALIGN = "txtAlign";
}
