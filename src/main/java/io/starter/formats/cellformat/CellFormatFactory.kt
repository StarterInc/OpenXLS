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

import io.starter.toolkit.StringTool

class CellFormatFactory private constructor() {
    companion object {
        @JvmStatic
        fun fromPatternString(pattern: String?): CellFormat {
            if (null == pattern || "" == pattern || "General".equals(pattern, ignoreCase = true)) {
                return GeneralCellFormat()
            }
            val pats = pattern.split(";".toRegex()).toTypedArray()
            val tester = StringTool.convertPatternExtractBracketedExpression(pats[0])
            val regex = Regex(".*(((y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}).*)+).*")
            if (tester.matches(regex)) {
                val string: String
                string = if (pats.size > 3) {
                    StringTool.convertPatternFromExcelToStringFormatter(pats[3], false)
                } else {
                    "%s"
                }
                return DateCellFormat(
                        StringTool.convertDatePatternFromExcelToStringFormatter(tester), string)
            }
            val positive: String
            val negative: String
            val zero: String
            val string: String
            positive = StringTool.convertPatternFromExcelToStringFormatter(pats[0], false)
            negative = StringTool.convertPatternFromExcelToStringFormatter(
                    pats[if (pats.size > 1) 1 else 0], true)
            zero = if (pats.size > 2) {
                StringTool.convertPatternFromExcelToStringFormatter(pats[2], false)
            } else {
                positive
            }
            string = if (pats.size > 3) {
                StringTool.convertPatternFromExcelToStringFormatter(pats[3], false)
            } else {
                "%s"
            }
            return NumberCellFormat(positive, negative, zero, string)
        }
    }

    init { // this is a static-only class
        throw UnsupportedOperationException()
    }
}