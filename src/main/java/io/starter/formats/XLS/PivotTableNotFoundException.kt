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


/**
 * **No PivotTable Found.**
 *
 *
 * Thrown when a PivotTableHandle is requested for a non-existent Pivot Table.
 *
 * @see PivotTableHandle
 */

class PivotTableNotFoundException(n: String) : Exception() {
    internal var pivotTableName = ""

    init {
        pivotTableName = n
    }

    override fun getMessage(): String {
        return this.toString()
    }

    override fun toString(): String {
        return "PivotTable Not Found in File. : '$pivotTableName'"
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4600384378296560405L
    }

}