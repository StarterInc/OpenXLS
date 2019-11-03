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

import java.io.IOException
import java.io.OutputStream
import java.util.concurrent.locks.ReentrantReadWriteLock

/**
 * An `IndirectOutputStream` forwards all requests unmodified to
 * another output stream which may be changed at runtime. By default an
 * `IOException` will be thrown if no sink is configured when a
 * request is received. The stream may be configured to drop such requests
 * instead.
 */
class IndirectOutputStream
/**
 * Creates a new `IndirectOutputStream` with the given sink
 * and behavior when no sink is present.
 *
 * @param sink    the initial sink
 * @param discard whether to discard requests when no sink is present
 */
@JvmOverloads constructor(private var sink: OutputStream? = null, private var discardOnNull: Boolean = false) : OutputStream() {
    private val lock = ReentrantReadWriteLock()

    /**
     * Gets the currently configured sink.
     *
     * @return the stream to which requests are currently being forwarded
     * or `null` if no sink present
     */
    fun getSink(): OutputStream? {
        // don't bother with a read lock, a single read is atomic
        return sink
    }

    /**
     * Sets the stream to which requests are forwarded.
     *
     * @param sink the stream to which requests should be forwarded
     * or `null` to remove the current sink
     */
    fun setSink(sink: OutputStream) {
        lock.writeLock().lock()
        try {
            this.sink = sink
        } finally {
            lock.writeLock().unlock()
        }
    }

    /**
     * Gets the current behavior when no wink is present.
     *
     * @return whether requests will be discarded when no sink is present
     */
    fun discardOnNoSink(): Boolean {
        // don't bother with a read lock, a single read is atomic
        return discardOnNull
    }

    /**
     * Sets the behavior when no sink is present.
     *
     * @param discard whether to discard requests when no sink is present
     */
    fun discardOnNoSink(discard: Boolean) {
        // don't bother with a write lock, a single write is atomic
        // as are all uses of this field
        this.discardOnNull = discard
    }

    @Throws(IOException::class)
    private fun checkSink(): Boolean {
        return if (null == sink) {
            if (discardOnNull)
                true
            else
                throw IOException("sink not connected")
        } else false
    }

    @Synchronized
    @Throws(IOException::class)
    override fun write(b: Int) {
        lock.readLock().lock()
        try {
            if (this.checkSink()) return
            sink!!.write(b)
        } finally {
            lock.readLock().unlock()
        }
    }

    @Synchronized
    @Throws(IOException::class)
    override fun write(b: ByteArray) {
        lock.readLock().lock()
        try {
            if (this.checkSink()) return
            sink!!.write(b)
        } finally {
            lock.readLock().unlock()
        }
    }

    @Synchronized
    @Throws(IOException::class)
    override fun write(b: ByteArray, off: Int, len: Int) {
        lock.readLock().lock()
        try {
            if (this.checkSink()) return
            sink!!.write(b, off, len)
        } finally {
            lock.readLock().unlock()
        }
    }

    @Synchronized
    @Throws(IOException::class)
    override fun flush() {
        lock.readLock().lock()
        try {
            if (this.checkSink()) return
            sink!!.flush()
        } finally {
            lock.readLock().unlock()
        }
    }

    @Synchronized
    @Throws(IOException::class)
    override fun close() {
        lock.readLock().lock()
        try {
            if (this.checkSink()) return
            sink!!.close()
        } finally {
            lock.readLock().unlock()
        }
    }
}
/**
 * Creates a new `IndirectOutputStream` with no sink which
 * fails when no sink is present.
 */
/**
 * Creates a new `IndirectOutputStream` with the given sink
 * which fails when no sink is present.
 *
 * @param sink the initial sink
 */
