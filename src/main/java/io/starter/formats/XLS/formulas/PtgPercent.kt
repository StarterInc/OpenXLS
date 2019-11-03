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
   Ptg that indicates percentage, divides top operand by 100
    
 * @see Ptg
 * @see Formula

    
*/
package io.starter.formats.XLS.formulas

import io.starter.toolkit.Logger

/**
 *
 */
class PtgPercent : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true

    override val isUnaryOperator: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = "%"

    override val length: Int
        get() = Ptg.PTG_PERCENT_LENGTH

    /*  Operator specific calculate method, this one returns a single value sent to it.

     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        // 20090202 KSC: Handle array formulas
        val o = GenericPtg.getValuesFromPtgs(form)
        if (!o!![0].javaClass.isArray()) {
            //double[] dub = super.getValuesFromPtgs(form);
            // there should always be only two ptg's in this, error if not.
            if (o == null || o.size != 1) {
                // there should always be only one ptg in this, error if not.
                //if (form.length != 1){
                Logger.logWarn("calculating formula failed, wrong number of values in PtgPercent")
                return null
            }
        }
        // TODO: finish for Array formulas
        val res = (o[0] as Double).toDouble() / 100
        return PtgNumber(res)
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8559541841405018157L
    }

}
