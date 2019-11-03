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

import io.starter.toolkit.CompatibleVector

import javax.naming.InvalidNameException
import javax.naming.Name
import java.util.Enumeration

/*
	Name add(int posn, String comp)
			  Adds a single component at a specified position within this name.
	 Name add(String comp)
			  Adds a single component to the end of this name.
	 Name addAll(int posn, Name n)

	 Name addAll(Name suffix)
			  Adds the components of a name -- in order -- to the end of this name.
	 Object clone()
			  Generates a new copy of this name.
	 int compareTo(Object obj)
			  Compares this name with another name for order.
	 boolean endsWith(Name n)
			  Determines whether this name ends with a specified suffix.
	 String get(int posn)
			  Retrieves a component of this name.
	 Enumeration getAll()
			  Retrieves the components of this name as an enumeration of strings.
	 boolean isEmpty()
			  Determines whether this name is empty.
	 Object remove(int posn)
			  Removes a component from this name.
	 int size()
			  Returns the number of components in this name.
	 boolean startsWith(Name n)
			  Determines whether this name starts with a specified prefix.
*/

class NameImpl : Name {
    /**
     * @return
     */
    /**
     * @param vector
     */
    internal var vals = CompatibleVector() /*
     * (non-Javadoc)
     *
     *
     * * @see javax.naming.Name#clone()
     */

    override fun clone(): Any {
        val nimple = NameImpl()
        val newvals = CompatibleVector()
        newvals.addAll(vals)
        nimple.vals = newvals
        return nimple
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#remove(int)
     */
    @Throws(InvalidNameException::class)
    override fun remove(arg0: Int): Any {
        return vals.removeAt(arg0)
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#get(int)
     */
    override fun get(arg0: Int): String {
        return vals[arg0].toString()
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#getAll()
     */
    override fun getAll(): Enumeration<*> {
        return vals.elements()
    }

    /*
     * Creates a name whose components consist of a prefix of the components of this
     * name.
     *
     * @see javax.naming.Name#getPrefix(int)
     */
    override fun getPrefix(arg0: Int): Name? {
        return null
    }

    /*
     * Creates a name whose components consist of a suffix of the components in this
     * name.
     *
     * @see javax.naming.Name#getSuffix(int)
     */
    override fun getSuffix(arg0: Int): Name? {
        return null
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#add(java.lang.String)
     */
    @Throws(InvalidNameException::class)
    override fun add(arg0: String): Name? {
        return null
    }

    /*
     * Adds the components of a name -- in order -- at a specified position within
     * this name.
     *
     * @see javax.naming.Name#addAll(int, javax.naming.Name)
     */
    @Throws(InvalidNameException::class)
    override fun addAll(arg0: Int, arg1: Name): Name {
        this.vals.addAll(arg0, (arg1 as NameImpl).vals)
        return this
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#addAll(javax.naming.Name)
     */
    @Throws(InvalidNameException::class)
    override fun addAll(arg0: Name): Name {
        this.vals.addAll((arg0 as NameImpl).vals)
        return this
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#size()
     */
    override fun size(): Int {
        return vals.size
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#isEmpty()
     */
    override fun isEmpty(): Boolean {
        return vals.size > 0
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#compareTo(java.lang.Object)
     */
    override fun compareTo(arg0: Any): Int {
        return this.compareTo(arg0)
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#endsWith(javax.naming.Name)
     */
    override fun endsWith(arg0: Name): Boolean {
        val ob1 = arg0.get(arg0.size() - 1)
        val ob2 = this.get(this.size() - 1)
        return ob1 == ob2
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#startsWith(javax.naming.Name)
     */
    override fun startsWith(arg0: Name): Boolean {
        val ob1 = arg0.get(0)
        val ob2 = this.get(0)
        return ob1 == ob2
    }

    /*
     * (non-Javadoc)
     *
     * @see javax.naming.Name#add(int, java.lang.String)
     */
    @Throws(InvalidNameException::class)
    override fun add(arg0: Int, arg1: String): Name {
        vals[arg0] = arg1
        return this
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4387233472850688497L
    }

}