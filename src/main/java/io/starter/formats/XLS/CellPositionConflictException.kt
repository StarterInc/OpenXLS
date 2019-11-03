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
 * thrown when trying to add a BiffRec to the position of an
 * existing Cell.
 */
class CellPositionConflictException(n: String) : java.lang.Exception() {
    internal var cellname = ""

    init {
        cellname = n
    }

    override fun getMessage(): String {
        // This method is derived from class java.lang.Throwable
        // to do: code goes here
        return this.toString()
    }

    override fun toString(): String {
        // This method is derived from class java.lang.Throwable
        // to do: code goes here
        return "BiffRec position conflicts with another Cell.  : '$cellname'"
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -923697641625221233L
    }

}