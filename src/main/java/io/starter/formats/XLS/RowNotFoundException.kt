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
 * **thrown when trying to access a Row and the Row is not Found.**
 *
 * @see Row
 *
 * @see WorkBook
 */

class RowNotFoundException(n: String) : java.lang.Exception() {
    internal var rowname = ""

    init {
        rowname = n
    }

    override fun getMessage(): String {
        // This method is derived from class java.lang.Throwable
        // to do: code goes here
        return this.toString()
    }

    override fun toString(): String {
        // This method is derived from class java.lang.Throwable
        // to do: code goes here
        return "Row Not Found in File. : '$rowname'"
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 1754346075847876028L
    }

}