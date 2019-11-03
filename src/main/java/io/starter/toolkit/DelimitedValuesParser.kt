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
package io.starter.toolkit

import java.io.IOException
import java.io.Reader

/**
 * Stream parser for delimiter-separated values formats.
 * These include comma separated values (CSV) and tab separated values (TSV).
 */
class DelimitedValuesParser(private val source: Reader) {

    /**
     * The delimiter used to separate values.
     */
    private val delimiter = '\t'

    /**
     * Contains the current value, if any.
     */
    private val value = StringBuilder()

    /**
     * The last token returned.
     */
    private var current: Token? = null

    /**
     * The next token to be returned.
     * This is used when a single character ends a token and is itself a token.
     */
    private var next: Token? = null

    /**
     * Represents the type of a token.
     */
    enum class Token {
        VALUE, NEWLINE, EOF
    }

    @Throws(IOException::class)
    operator fun next(): Token {
        // reset the value builder
        value.setLength(0)

        // if there's a token waiting, return it
        if (next != null) {
            current = next
            next = null
            return current
        }

        while (true) {
            val read = source.read()
            if (read == -1) {
                return if (value.length == 0)
                    current = Token.EOF
                else
                    current = Token.VALUE
            }

            if (read == delimiter.toInt()) return current = Token.VALUE

            if (read == '\n'.toInt()) {
                if (value.length > 0) {
                    if (value[value.length - 1] == '\r')
                        value.setLength(value.length - 1)
                    next = Token.NEWLINE
                    return current = Token.VALUE
                } else {
                    return current = Token.NEWLINE
                }
            }

            value.append(read.toChar())
        }
    }

    fun getValue(): String? {
        return if (value.length > 0) value.toString() else null
    }
}
