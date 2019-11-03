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

import io.starter.OpenXLS.DateConverter
import io.starter.formats.XLS.BiffRec

object Calculator {
    /**
     * given a BiffRec Cell Record, an Object and an operator to compare to,
     * return true if the comparison passes, false otherwise
     *
     * @param c  BiffRec cell record
     * @param o  Object value - one of Double, String or Boolean
     * @param op - String operator - one of "=", ">", ">=", "<", "<=" or "<>
     * @return true if comparison of operator and value with cell value passes, false otherwise
     */
    fun compareCellValue(c: BiffRec, o: Any, op: String): Boolean {
        // doper types:  numeric:  ieee, rk
        //               string:   string doper
        // 				 boolean
        //				 error
        val compare: Int
        try {
            if (o is Boolean)
            // TODO: 1.5 use Boolean.compareTo
                compare = o.toString().compareTo(java.lang.Boolean.valueOf(c.booleanVal).toString())
            else if (o is String) {
                // use "matches" to handle wildcards
                if (o.toUpperCase().matches(c.stringVal.toUpperCase().toRegex()))
                    compare = 0    // equal or matches
                else
                    compare = -1    // doesn't equal
                //compare= ((String) o).toUpperCase().compareTo(c.getStringVal().toUpperCase());
            } else
            // it's a Double
                compare = (o as Double).compareTo(c.dblVal)
        } catch (e: Exception) {
            // report error?
            return false
        }

        if (op == "=") {
            return compare == 0
        } else if (op == "<") {
            return compare > 0
        } else if (op == "<=") {
            return compare >= 0
        } else if (op == ">") {
            return compare < 0
        } else if (op == ">=") {
            return compare <= 0
        } else if (op == "<>") {
            return compare != 0
        }
        return false
    }

    fun compareCellValue(`val`: Any, compareval: String, op: String): Boolean {
        // doper types:  numeric:  ieee, rk
        //               string:   string doper
        // 				 boolean
        //				 error
        var compare = -1
        try {
            if (`val` is Boolean)
            // TODO: 1.5 use Boolean.compareTo
                compare = java.lang.Boolean.valueOf(compareval).toString().compareTo(`val`.toString())
            else if (`val` is String) {
                if (compareval.indexOf('?') == -1 && compareval.indexOf('*') == -1)
                // if no wildcards
                    compare = compareval.compareTo(`val`.toUpperCase())
                else {    // use "matches" to handle wildcards
                    if (`val`.toUpperCase().matches(compareval.toRegex()))
                        compare = 0    // equal or matches
                    else
                        compare = -1    // doesn't equal
                }
            } else if (`val` is Number)
            // assume it's a number
                compare = Double(compareval).compareTo(`val`.toDouble())
            else
                return false
        } catch (e: Exception) {
            try {    // try date compare
                val dt = DateConverter.getXLSDateVal(java.util.Date(compareval))
                compare = dt.compareTo((`val` as Number).toDouble())
            } catch (ex: Exception) {    // just try string compare
                compare = compareval.compareTo(`val`.toString())
            }

        }

        if (op == "=") {
            return compare == 0
        } else if (op == "<") {
            return compare > 0
        } else if (op == "<=") {
            return compare >= 0
        } else if (op == ">") {
            return compare < 0
        } else if (op == ">=") {
            return compare <= 0
        } else if (op == "<>") {
            return compare != 0
        }
        return false
    }

    /*    } else if (op.equals("<")) {
    	passes= (compare < 0);
	} else if (op.equals("<=")) {
		passes= (compare <= 0);
	} else if (op.equals(">")) {
		passes= (compare > 0);
	} else if (op.equals(">=")) {
		passes= (compare >= 0);
	}
  */

    /**
     * translate Excel-style wildcards into Java wildcards in criteria string
     * plus handle qualified wildcard characters + percentages ...
     *
     * @param sCriteria criteria string
     * @return tranformed criteria string
     */
    fun translateWildcardsInCriteria(sCriteria: String): String {
        var sCriteria = sCriteria
        var s = ""    // handle wildcards
        var qualified = false
        var isalldigits = true
        for (i in 0 until sCriteria.length) {
            val c = sCriteria[i]
            if (c == '~') {
                qualified = true    // don't add tilde unless certain it's not qualifying a * or ?
            } else if (c == '*') {
                if (!qualified)
                    s += "."
                s += c
            } else if (c == '?') {
                if (!qualified)
                    s += "."
                s += c
            } else if (c == '%') { // translate percentage into decimals
                if (isalldigits) {
                    s = "0$s"
                    s = s.substring(s.length - 2, 2)
                    s = ".$s"
                }
            } else {
                if (qualified)
                // really add the tilde
                    s += '~'.toString()
                s += c
                qualified = false
                if (!Character.isDigit(c))
                    isalldigits = false
            }
        }
        sCriteria = s.toUpperCase()    // matching is case-insensitive
        return sCriteria
    }

    /**
     * given a criteria string that starts with an operator,
     * parse and return the index that the operator ends and the crtieria starts
     *
     * @param criteria
     * @return int i	position in criteria which actual criteria starts
     */
    fun splitOperator(criteria: String): Int {
        var i = 0
        while (i < criteria.length) {
            val c = criteria[i]
            if (Character.isJavaIdentifierPart(c))
                break
            else if (c == '*' || c == '?')
                break
            i++
        }
        return i
    }

    /**
     * takes a Reference Type Ptg and deferences and PtgNames, etc.
     * to return a PtgArea
     *
     * @param p
     * @return
     */
    @Throws(IllegalArgumentException::class)
    fun getRange(p: Ptg): PtgArea? {
        if (p is PtgArea)
            return p
        if (p is PtgName) {    // get source range
            var pr: Array<Ptg>? = null
            try {
                pr = p.name!!.cellRangePtgs
                return pr!![0] as PtgArea
            } catch (e: Exception) {
                try {    // if it's a PtgRef, convert to a PtgArea
                    if (pr!![0] !is PtgArea && pr[0] is PtgRef) {
                        val pa = PtgArea()
                        pa.parentRec = pr[0].parentRec
                        pa.location = pr[0].location
                        return pa
                    } else
                        throw IllegalArgumentException("Expected a reference-type operand")
                } catch (ex: Exception) {
                    throw IllegalArgumentException("Expected a reference-type operand")
                }

            }

        }
        return null
    }
}
