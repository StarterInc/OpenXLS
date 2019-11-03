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
/*
 * Created on May 7, 2007
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package io.starter.toolkit

import java.io.File

/**
 * @
 */
class FileFilter
/**
 *
 */
(fx: String) : java.io.FileFilter {

    internal var fix = "*" // the pattern to match

    init {
        this.fix = fx
    }

    /**
     * @see java.io.FileFilter.accept
     */
    override fun accept(pathname: File): Boolean {
        val fn = pathname.name
        return fn.endsWith(fix)
    }

    companion object {

        @JvmStatic
        fun main(args: Array<String>) {
        }
    }
}
