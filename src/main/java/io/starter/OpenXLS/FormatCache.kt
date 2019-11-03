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
/**
 * FormatCache.java
 */
package io.starter.OpenXLS

import io.starter.toolkit.Logger


/** Handles the caching of the Formats
 *
 * This class is no longer in use nor needed
 *
 */
// is this a valid class anymore?
@Deprecated("")
class FormatCache {

    internal var mpx: Map<*, *> = java.util.HashMap()

    /** Consolidate all identical formats to avoid too many formats errors
     *
     */
    @Deprecated("")
    fun pack() {
        val itx = mpx.keys.iterator()
        while (itx.hasNext()) {
            val oby = mpx.get(itx.next())
            var thisfmt = oby as FormatHandle
            thisfmt = this[thisfmt]
        }
    }

    /**
     * @return
     */
    @Deprecated(" ")
    operator fun get(fmx: FormatHandle): FormatHandle {
        val fmt = fmx.toString()
        if (!mpx.containsKey(fmt)) {
            Logger.logErr("missing in cache: FH $fmt")
        }
        return mpx.get(fmt) as FormatHandle
    }

    /**
     * @return
     */
    @Deprecated(" ")
    operator fun get(f: String): Any {
        return mpx.get(f)
    }

    /**
     * @return
     */
    @Deprecated(" ")
    fun getInt(f: String): Int {
        var findex = -1
        if (mpx.containsKey(f)) findex = (mpx.get(f) as Int).toInt()
        return findex
    }

}
