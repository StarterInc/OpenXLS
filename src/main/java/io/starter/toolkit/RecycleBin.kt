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

interface RecycleBin {

    /**
     * get an unused Recyclable item from the bin
     *
     * @throws RecycleBinFullException
     */
    val item: Recyclable

    /**
     * get all of the items in this bin
     */
    val all: List<*>

    /**
     * add a Recyclable to the
     * bin
     */
    @Throws(RecycleBinFullException::class)
    fun addItem(r: Recyclable)

    /**
     * add a Recyclable to the
     * bin
     */
    @Throws(RecycleBinFullException::class)
    fun addItem(key: Any, r: Recyclable)

    /**
     * empty the current contents of the bin
     */
    fun empty()

    /**
     * set the maximum number of items for this
     * bin
     */
    fun setMaxItems(i: Int)


}