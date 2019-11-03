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
package io.starter.naming


import javax.naming.NamingEnumeration
import javax.naming.NamingException
import java.util.Enumeration

class NamingEnumerationImpl : NamingEnumeration<Any> {

    private var e: Enumeration<*>? = null

    internal fun setEnumeration(ex: Enumeration<*>) {
        e = ex
    }

    /* (non-Javadoc)
     * @see javax.naming.NamingEnumeration#close()
     */
    @Throws(NamingException::class)
    override fun close() {
        e = null
    }

    /* (non-Javadoc)
     * @see javax.naming.NamingEnumeration#hasMore()
     */
    @Throws(NamingException::class)
    override fun hasMore(): Boolean {
        return e!!.hasMoreElements()
    }

    /* (non-Javadoc)
     * @see javax.naming.NamingEnumeration#next()
     */
    @Throws(NamingException::class)
    override fun next(): Any {
        return e!!.nextElement()
    }

    /* (non-Javadoc)
     * @see java.util.Enumeration#hasMoreElements()
     */
    override fun hasMoreElements(): Boolean {
        return e!!.hasMoreElements()
    }

    /* (non-Javadoc)
     * @see java.util.Enumeration#nextElement()
     */
    override fun nextElement(): Any {
        return e!!.nextElement()
    }

}