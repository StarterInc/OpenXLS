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


/*

    Undocumented PTG, the only thing we know is "ptg DELETED" from MEFF pg 446.  I am going to treat
    as a PtgRefErr for now.

 * @see Ptg
 * @see Formula

*/
class PtgEndSheet : GenericPtg(), Ptg {

    override val isOperand: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = "End Sheet Error"

    override val value: Any?
        get() = "#REF!"

    override val length: Int
        get() = Ptg.PTG_ENDSHEET_LENGTH

    companion object {
        // Excel can handle PtgRefErrors within formulas, as long as they are not the result so...

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2395432053363123361L
    }


}