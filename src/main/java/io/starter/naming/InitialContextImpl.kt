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

import javax.naming.*
import java.util.Hashtable

/**
 * A basic JNDI Context which holds a flat lookup of names
 */
open class InitialContextImpl : javax.naming.Context {

    internal var nameParser: NameParser = NameParserImpl()
    protected var env: Hashtable<Comparable<*>, Any>

    private var closed = false

    init {
        if (System.getProperties()[CONTEXT_ID] != null)
            this.env = System.getProperties()[CONTEXT_ID] as Hashtable<Comparable<*>, Any>
        else {
            val loadme = System.getProperty(LOAD_CONTEXT)
            env = Hashtable() // 20070518 KSC: Moved so gets init even if no LOAD_CONTEXT
            if (loadme != null) {
                if (loadme == "true") {
                    // env = new Hashtable(); KSC: See above
                    // this breaks properties
                    System.getProperties()[CONTEXT_ID] = env
                }
            }
        }
    }

    // check return... -jm
    @Throws(NamingException::class)
    override fun addToEnvironment(propName: String, propVal: Any): Any {
        if (env.contains(propVal)) {
            throw NamingException("Object $propName already exists in NamingContext.")
        } else {
            env[propName] = propVal
            return propVal
        }
    }

    // we use string to bind -- is that bad?
    @Throws(NamingException::class)
    override fun bind(name: Name, obj: Any) {
        val str = name.toString()
        this.bind(str, obj)
    }

    @Throws(NamingException::class)
    override fun bind(name: String, obj: Any) {
        try {
            this.addToEnvironment(name, obj)
        } catch (e: NamingException) {
            env.remove(obj)
            env[name] = obj // override
        }

    }

    @Throws(NamingException::class)
    override fun close() {
        closed = true
    }

    // ?
    @Throws(NamingException::class)
    override fun composeName(name: Name, prefix: Name): Name {
        val retval = NameImpl()
        retval.addAll(prefix)
        retval.addAll(name)
        return retval
    }

    @Throws(NamingException::class)
    override fun composeName(name: String, prefix: String): String {
        val sb = StringBuffer()
        sb.append(name)
        sb.append(prefix)
        return sb.toString()
    }

    @Throws(NamingException::class)
    override fun getEnvironment(): Hashtable<Comparable<*>, Any> {
        return env
    }

    @Throws(NamingException::class)
    override fun getNameParser(name: String): NameParser {
        this.nameParser.parse(name)
        return this.nameParser
    }

    @Throws(NamingException::class)
    override fun getNameParser(name: Name): NameParser {
        return this.nameParser
    }

    @Throws(NamingException::class)
    override fun lookup(name: Name): Any {
        return env[name]
    }

    @Throws(NamingException::class)
    override fun lookup(name: String): Any {
        return env[name]
    }

    @Throws(NamingException::class)
    override fun lookupLink(name: Name): Any {
        return env[name]
    }

    @Throws(NamingException::class)
    override fun lookupLink(name: String): Any {
        return env[name]
    }

    @Throws(NamingException::class)
    override fun rebind(name: Name, obj: Any) {
        this.bind(name, obj)
    }

    @Throws(NamingException::class)
    override fun rebind(name: String, obj: Any) {
        this.bind(name, obj)
    }

    @Throws(NamingException::class)
    override fun removeFromEnvironment(propName: String): Any {
        return env.remove(propName)
    }

    @Throws(NamingException::class)
    override fun rename(oldName: String, newName: String) {
        val ob = env[oldName]
        env.remove(oldName)
        env[newName] = ob
    }

    @Throws(NamingException::class)
    override fun rename(oldName: Name, newName: Name) {
        val ob = env[oldName]
        env.remove(oldName)
        env[newName] = ob
    }

    @Throws(NamingException::class)
    override fun unbind(name: String) {
        try {
            env.remove(env[name])
        } catch (e: Exception) {
            throw NamingException(e.toString())
        }

    }

    @Throws(NamingException::class)
    override fun unbind(name: Name) {
        try {
            env.remove(env[name])
        } catch (e: Exception) {
            throw NamingException(e.toString())
        }

    }

    // TODO: Implement the following mehods -jm 9/27/2004

    @Throws(NamingException::class)
    override fun list(name: String): NamingEnumeration<*>? {
        return null
    }

    @Throws(NamingException::class)
    override fun list(name: Name): NamingEnumeration<*>? {
        return null
    }

    @Throws(NamingException::class)
    override fun listBindings(name: Name): NamingEnumeration<*>? {
        return null
    }

    @Throws(NamingException::class)
    override fun listBindings(name: String): NamingEnumeration<*>? {
        return null
    }

    @Throws(NamingException::class)
    override fun createSubcontext(name: Name): Context? {
        // This method is derived from interface javax.naming.Context
        // to do: code goes here
        return null
    }

    @Throws(NamingException::class)
    override fun createSubcontext(name: String): Context? {
        // This method is derived from interface javax.naming.Context
        // to do: code goes here
        return null
    }

    @Throws(NamingException::class)
    override fun destroySubcontext(name: String) {
        // This method is derived from interface javax.naming.Context
        // to do: code goes here
    }

    @Throws(NamingException::class)
    override fun destroySubcontext(name: Name) {
        // This method is derived from interface javax.naming.Context
        // to do: code goes here
    }

    @Throws(NamingException::class)
    override fun getNameInNamespace(): String? {
        // This method is derived from interface javax.naming.Context
        // to do: code goes here
        return null
    }

    companion object {

        // provide persistence between instantiations
        var CONTEXT_ID = "io.starter.naming.InitialContextImpl_instance"
        var LOAD_CONTEXT = "io.starter.naming.load_context"
    }
}