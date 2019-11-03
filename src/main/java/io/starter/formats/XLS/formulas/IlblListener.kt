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
package io.starter.formats.XLS.formulas

/**
 * An IlblListener is aware of changes to the Named range references.
 *
 *
 * Notably, the PtgName  recs...
 */
interface IlblListener {


    /**
     * @return Returns the Ilbl.
     */
    /**
     * @param Ilbl The Ilbl to set.
     */
    var ilbl: Short

    /**
     * get the name of the named range this refers to
     *
     * @return
     */
    val storedName: String

    /**
     * Add this to the ilbl listeners
     */
    fun addListener()

    /**
     * Store the name string for matching missing references
     *
     * @param name
     */
    fun storeName(name: String)
}