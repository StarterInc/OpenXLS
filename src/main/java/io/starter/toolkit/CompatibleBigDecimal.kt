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

/**
 * CompatibleBigDecimal.java
 *
 *
 *
 *
 * CompatibleBigDecimal deals with the java 1.4-1.5 transition error of BigDecimal.toString().
 *
 *
 * Prior to 1.5, BigDecimal.toString would not return Scientific Notation.  JDK1.5 now allows Scientific Notation
 * to be returned in some cases for this method.  A new method has been created, .toPlainString that mimics the
 * behavior of the old .toString.   As we do not know the runtime JDK, and returning correctly formatted numbers
 * is critical to the functionality of OpenXLS, this class mimics 1.4 functionality for toString under either JDK.
 *
 *
 * A static method is used to increase performance for multiple calls.
 */


import java.lang.reflect.Method
import java.math.BigDecimal

class CompatibleBigDecimal : BigDecimal {

    /**
     * Constructor
     */
    constructor(bd: BigDecimal) : super(bd.unscaledValue(), bd.scale()) {}

    constructor(num: String) : super(num) {}

    /**
     * Compatible toString functionality
     *
     * @see java.math.BigDecimal.toString
     */
    fun toCompatibleString(): String? {
        if (_methodToString != null) {
            try {
                return _methodToString!!.invoke(this, *null as Array<Any>?) as String
            } catch (e: Exception) {
                Logger.logWarn("Error in calling CompatibleBigDecimal.toString")
            }

        }
        return null
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -6816994951413033200L
        private var _methodToString: Method? = null

        // Create the static method we can access later.  This could be .toPlainString or .toString.
        init {
            try {
                _methodToString = BigDecimal::class.java!!.getMethod("toPlainString", *null as Array<Class<*>>?)
            } catch (e: NoSuchMethodException) {
                try {
                    _methodToString = BigDecimal::class.java!!.getMethod("toString", *null as Array<Class<*>>?)
                } catch (ex: NoSuchMethodException) {
                    Logger.logWarn("Error creating toString method in CompatibleBigDecimal")
                }

            }

        }
    }

}
