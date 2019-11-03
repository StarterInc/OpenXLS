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
 * Created on Dec 4, 2004
 *
 * To change the template for this generated file go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
package io.starter.formats.XLS

import java.io.OutputStream

/**
 * To change the template for this generated type comment go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
interface Book {

    val continueHandler: ContinueHandler

    val streamer: ByteStreamer

    val fileName: String

    /**
     * get a handle to the factory
     */
    var factory: WorkBookFactory


    fun addRecord(b: BiffRec, y: Boolean): BiffRec

    fun setFirstBof(b: Bof)

    /**
     * Stream the Book bytes to out
     *
     * @param out
     * @return
     */
    fun stream(out: OutputStream): Int

    /**
     * set the Debug level
     */
    fun setDebugLevel(i: Int)

    override fun toString(): String

    /**
     * Dec 15, 2010
     *
     * @param rec
     */
    fun getSheetFromRec(rec: BiffRec, l: Long?): Boundsheet

    /** set readiness -- is it done parsing?  */
    // public boolean isReady();
    // public void setReady(boolean t);

    /** the default session provides initial values
     * @return
     */
    //public BookSession getDefaultSession();
    //public void setDefaultSession(BookSession session);

}