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
 * Validation Exceptions are thrown when a cell value is set to a value that does not pass
 * validity  of an Excel validation record affecting said cell
 */
class ValidationException(eTitle: String, eText: String) : Exception() {

    /**
     * Returns the title of the validation error dialog.
     */
    var title = ""
    /**
     * Returns the body of the validation error dialog.
     */
    var text = ""


    init {
        title = eTitle
        text = eText
    }

    override fun getMessage(): String {
        return this.toString()
    }

    override fun toString(): String {
        return "$title: $text"
    }

    companion object {
        private val serialVersionUID = -6448974788123912538L
    }

}
