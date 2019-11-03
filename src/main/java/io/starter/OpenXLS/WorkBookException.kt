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
 * WorkBookInstantiationException is thrown when a workbook cannot be parsed for a particular reason.
 *
 *
 * Error codes can be retrieved with getErrorCode, which map to the static error ints
 */
class WorkBookException : io.starter.formats.XLS.WorkBookException {

    constructor(n: String, x: Int) : super(n, x) {}

    constructor(string: String, x: Int,
                e: Exception) : super(string, x, e) {
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -5313787084750169461L
        val DOUBLE_STREAM_FILE = 0
        val NOT_BIFF8_FILE = 1
        val LICENSING_FAILED = 2
        val UNSPECIFIED_INIT_ERROR = 3
        val RUNTIME_ERROR = 4
        val SMALLBLOCK_FILE = 5
        val WRITING_ERROR = 6
        val DECRYPTION_ERROR = 7
        val DECRYPTION_INCORRECT_PASSWORD = 8
        val ENCRYPTION_ERROR = 9
        val DECRYPTION_INCORRECT_FORMAT = 10
        val ILLEGAL_INIT_ERROR = 11
        val READ_ONLY_EXCEPTION = 12
        val SHEETPROTECT_INCORRECT_PASSWORD = 13
    }


}
