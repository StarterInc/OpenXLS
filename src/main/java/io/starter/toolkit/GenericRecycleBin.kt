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

import java.util.*

/**
 * A recycling cache, items are checked at intervals
 */
abstract class GenericRecycleBin : java.lang.Thread(), Map<Any, Any>, io.starter.toolkit.RecycleBin {
    protected var map: MutableMap<Any, Any> = java.util.HashMap()
    protected var active = Vector<Any>()
    protected var spares = Stack<Recyclable>()

    /**
     * returns number of items in cache
     *
     * @return
     */
    val numItems: Int
        get() = active.size

    override val all: List<Any>
        @Synchronized get() = active

    /**
     * returns a new or recycled item from the spares pool
     *
     * @see io.starter.toolkit.RecycleBin.getItem
     */
    override// spares contains the recycled
    // technically infinite loop until exception thrown
    val item: Recyclable?
        @Synchronized @Throws(RecycleBinFullException::class)
        get() {
            var active: Recyclable? = null
            if (spares.size > 0) {
                active = spares.pop()
                addItem(active)
                return active
            }
            recycle()
            return item
        }

    protected var MAXITEMS = -1 // no limit is default

    val spareCount: Int
        get() = spares.size

    /**
     * add an item
     */
    @Throws(RecycleBinFullException::class)
    override fun addItem(r: Recyclable?) {
        if (MAXITEMS == -1 || map.size < MAXITEMS) {
            addItem(Integer.valueOf(map.size), r)
        } else {
            throw RecycleBinFullException()
        }
    }

    @Throws(RecycleBinFullException::class)
    override fun addItem(key: Any?, r: Recyclable?) {
        // recycle();
        if (MAXITEMS == -1 || map.size < MAXITEMS) {
            active.add(r)
            map[key] = r

        } else {
            throw RecycleBinFullException()
        }
    }

    /**
     * iterate all active items and try to recycle
     */
    @Synchronized
    fun recycle() {
        val rs = arrayOfNulls<Recyclable>(active.size)
        active.copyInto(rs)
        for (t in rs.indices) {
            try {
                val rb = rs[t]
                if (!rb.inUse()) {
                    // recycle it
                    rb.recycle()

                    // remove from active and lookup
                    active.remove(rb)
                    map.remove(rb)

                    // put in spares
                    spares.push(rb)

                }
            } catch (ex: Exception) {
                Logger.logErr("recycle failed", ex)
            }

        }

    }

    override fun empty() {
        map.clear()
        active.clear()
    }

    /**
     * max number of items to be put in this bin.
     */
    override fun setMaxItems(i: Int) {
        MAXITEMS = i
    }

    fun getMaxItems(): Int {
        return MAXITEMS
    }

    override fun clear() {
        map.clear()
        active.clear()
    }

    override fun containsKey(key: Any): Boolean {
        return map.containsKey(key)

    }

    override fun containsValue(value: Any): Boolean {
        return map.containsValue(value)
    }

    override fun entrySet(): Set<*> {
        return map.entries
    }

    override fun equals(o: Any?): Boolean {
        return map == o
    }

    override fun get(key: Any): Any {
        return map[key]
    }

    override fun hashCode(): Int {
        return map.hashCode()
    }

    override fun isEmpty(): Boolean {
        return map.isEmpty()
    }

    override fun keySet(): Set<Any> {
        return map.keys
    }

    override fun put(arg0: Any, arg1: Any): Any? {
        active.add(arg1)
        return map.put(arg0, arg1)
    }

    override fun putAll(arg0: Map<*, *>) {
        active.addAll(arg0.entries)
        map.putAll(arg0)
    }

    override fun remove(key: Any): Any {
        active.remove(map[key])
        return map.remove(key)
    }

    override fun size(): Int {
        return map.size
    }

    override fun toString(): String {
        return map.toString()
    }

    override fun values(): Collection<Any> {
        return map.values
    }

    fun getMap(): Map<*, *> {
        return map
    }

    fun setMap(_map: java.util.HashMap<*, *>) {
        map = _map
    }

}