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
 *
 */
open class WorkBookException : RuntimeException {

    val errorCode: Int

    /**
     * Obsolete synonym for `getCause()`.
     *
     */
    val wrappedException: Exception?
        @Deprecated("Use {@link Throwable#getCause()} instead.")
        get() {
            val cause = cause
            return if (cause is Exception) cause else null
        }

    constructor(message: String, code: Int) : super(message) {
        errorCode = code
    }

    constructor(message: String, code: Int, cause: Throwable) : super(message, cause) {
        errorCode = code
    }

    override fun toString(): String {
        return "WorkBook initialization failed: '$message'"
    }

    companion object {
        private val serialVersionUID = 6406202417397276014L

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
    }
}
