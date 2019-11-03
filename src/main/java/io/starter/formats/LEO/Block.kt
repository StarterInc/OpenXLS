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
package io.starter.formats.LEO

import java.io.OutputStream
import java.nio.ByteBuffer


/**
 * the basic unit of data in a LEO file.  Can either be BIG or SMALL.
 */
interface Block : Iterator<*> {

    /**
     * @return
     */
    val isXBAT: Boolean

    /**
     * get the size of the Block data in bytes
     *
     * @return block data size
     */
    val blockSize: Int

    /**
     * get the index of this Block in the storage
     * Vector
     */
    val blockIndex: Int

    /**
     * set the storage for this Block
     */
    /**
     * set the storage for this Block
     */
    var storage: Storage

    /**
     * returns whether this block has been
     * added to the output stream
     */
    /**
     * sets whether this block has been
     * added to the output stream
     */
    var streamed: Boolean

    /**
     * returns the int representing the block type
     */
    val blockType: Int

    /**
     * returns true if this is a Block Depot block
     * that needs to be ignored when reading byte storages
     */
    val isSpecialBlock: Boolean

    /**
     * returns true if this is a Block Depot block
     * that needs to be ignored when reading byte storages
     */
    /**
     * set to true if this is a Block Depot block
     */
    var isDepotBlock: Boolean

    /**
     * returns whether this Block has been read yet...
     */
    /**
     * set whether this Block has been read yet...
     */
    var initialized: Boolean

    /**
     * get the data bytes  on this Block
     */
    val byteBuffer: ByteBuffer

    /**
     * get the data bytes  on this Block
     */
    val bytes: ByteArray

    /**
     * return the original BB position in the file
     */
    val originalPos: Int

    /**
     * return the original BB position in the file
     */
    /**
     * set the original BB position in the file
     */
    var originalIdx: Int

    /**
     * @param b
     */
    fun setIsExtraSector(b: Boolean)

    /**
     * link to the vector of blocks for the storage
     */
    fun setBlockVector(v: List<*>)

    /**
     * link the next Block in the chain
     */
    fun setNextBlock(b: Block)

    /**
     * init the Block Data
     */
    fun init(d: ByteBuffer, origidx: Int, origp: Int)

    /**
     * set the data bytes  on this Block
     */
    fun setBytes(b: ByteBuffer)

    /**
     * return the byte Array for this BLOCK
     */
    fun getBytes(start: Int, end: Int): ByteArray

    /**
     * write the data bytes on this Block to out
     */
    fun writeBytes(out: OutputStream)

    companion object {

        val SMALL = 0
        val BIG = 1
    }
}