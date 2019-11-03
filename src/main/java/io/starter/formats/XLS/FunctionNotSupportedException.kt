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
package io.starter.formats.XLS

import io.starter.formats.XLS.formulas.FunctionConstants

import java.util.Locale

/**
 * **Formula function is not supported for calculation.**
 *
 * @see Formula
 */

class FunctionNotSupportedException(n: String) : java.lang.RuntimeException() {
    internal var functionName = ""

    init {
        functionName = n
    }

    override fun getMessage(): String {
        // This method is derived from class java.lang.Throwable
        // to do: code goes here
        return this.toString()
    }

    override fun toString(): String {
        // if it is a formula calc that is failing
        //		if (functionName.length() > 4){
        if (true) { // 20081203 KSC: functionName is offending function
            return "Function Not Supported: $functionName"
        }
        var fID = 0
        var f: String? = "Unknown Formula"
        try {
            fID = Integer.parseInt(functionName, 16)    // hex
            if (Locale.JAPAN == Locale.getDefault()) {
                f = FunctionConstants.getJFunctionString(fID.toShort())
            }
            if (f == "Unknown Formula") {
                f = FunctionConstants.getFunctionString(fID.toShort())
            }
            if (f!!.length == 0) {
                if (fID == FunctionConstants.xlfADDIN)
                    f = "AddIn Formula"
                else
                    f = "Unknown Formula"
            } else {
                f += ")"    // add ending paren
            }
        } catch (e: Exception) {
        }

        return "Function: $f $functionName is not implemented."
    }

    companion object {


        private val serialVersionUID = 3569219212252117988L
    }
}