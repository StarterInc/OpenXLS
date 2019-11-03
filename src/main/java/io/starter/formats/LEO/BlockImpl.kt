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

import io.starter.toolkit.CompatibleVectorHints
import io.starter.toolkit.Logger

import java.io.OutputStream
import java.io.Serializable
import java.nio.ByteBuffer


/**
 * LEO File Block Information Record
 *
 *
 * These blocks of data contain information related to the
 * LEO file format data blocks.
 */
abstract class BlockImpl : io.starter.formats.LEO.Block, CompatibleVectorHints, Serializable {
    /*allows the block to be populated with a byte array
     rather than just a bytebuffer, easing debugging
    **/
    internal var DEBUG = false

    /**
     * methods from CompatibleVectorHints
     */
    /**
     * provide a hint to the CompatibleVector
     * about this objects likely position.
     */
    /**
     * set index information about this
     * objects likely position.
     */
    @Transient
    override var recordIndexHint = -1
    @Transient
    internal var lastidx = -1
    /**
     * return the ByteBuffer for this BLOCK
     */
    @Transient
    override var byteBuffer: ByteBuffer? = null
        internal set // new byte[SIZE];
    private var blockvec: MutableList<*>? = null

    // implement iterator for use in
    // chaining blocks
    var nextblock: Block? = null
    /**
     * @return
     */
    override var isXBAT = false
        private set

    /**
     * return the original position of this BIGBLOCK
     */
    /**
     * set the original BB position in the file
     */
    override var originalIdx: Int = 0
    /**
     * return the original position of this
     * BIGBLOCK record in the array of BIGBLOCKS
     * that make up the file.
     */
    override var originalPos: Int = 0
        internal set
    internal var isBBDepotBlock = false
    internal var isSBDepotBlock = false
    /**
     * returns true if this is a Block Depot block
     * that needs to be ignored when reading byte storages
     */
    override var isSpecialBlock = false
        internal set
    /**
     * returns whether this Block has been read yet...
     */
    /**
     * set whether this Block has been read yet...
     */
    override var initialized = false
    /** returns the original BB pos
     *
     * public int getOriginalPos(){return  this.originalpos;}
     */
    /**
     * returns whether this block has been
     * added to the output stream
     */
    /**
     * sets whether this block has been
     * added to the output stream
     */
    override var streamed = false
    /**
     * get the storage for this Block
     */
    /**
     * set the storage for this Block
     */
    override var storage: Storage? = null

    override val blockSize: Int
        get() = if (this.blockType == Block.BIG)
            BIGBLOCK.SIZE
        else if (this.blockType == Block.SMALL)
            SMALLBLOCK.SIZE
        else
            0

    /**
     * return the byte Array for this BLOCK
     */
    override// why is this hitting????????
    // CAN HAPPEN ON OUT-OF-SPEC FILES whom have last block size < 512
    val bytes: ByteArray
        get() {
            var SIZE = 0
            if (this.blockType == Block.BIG)
                SIZE = BIGBLOCK.SIZE
            else if (this.blockType == Block.SMALL)
                SIZE = SMALLBLOCK.SIZE
            val ret = ByteArray(SIZE)

            val capcheck = byteBuffer!!.capacity()
            if (capcheck <= originalPos)
                originalPos = 0
            if (capcheck < SIZE + originalPos)
                SIZE = capcheck - originalPos
            try {
                byteBuffer!!.position(originalPos)
                byteBuffer!!.get(ret, 0, SIZE)
            } catch (e: Exception) {
                Logger.logWarn("BlockImpl.getBytes(0,$SIZE): $e")
            }

            return ret
        }

    /**
     * get the index of this Block in the storage
     * Vector
     */
    override val blockIndex: Int
        get() = if (blockvec == null) this.recordIndexHint else blockvec!!.indexOf(this)

    /**
     * returns true if this is a Block Depot block
     * that needs to be ignored when reading byte storages
     */
    /**
     * set to true if this is a Block Depot block
     */
    override var isDepotBlock: Boolean
        get() = isBBDepotBlock
        set(b) {
            isSpecialBlock = b
            isBBDepotBlock = b
        }


    fun close() {
        if (blockvec != null) {
            blockvec!!.clear()
            blockvec = null
        }
        if (nextblock != null && nextblock !== this) {
            nextblock = null
        }
        storage = null
        if (byteBuffer != null) {
            byteBuffer!!.clear()
            byteBuffer = null
        }
    }

    /**
     * Write the entire bytes directly to out
     *
     * @see io.starter.formats.LEO.Block.writeBytes
     */
    override fun writeBytes(out: OutputStream) {
        try {
            out.write(bytes)
        } catch (exp: Exception) {
            Logger.logErr("BlockImpl.writeBytes failed.", exp)
        }

    }


    //byte[] delbytes = null;

    /**
     * return the byte Array for this BLOCK
     */
    override fun getBytes(start: Int, end: Int): ByteArray {
        var start = start
        //if (delbytes == null) delbytes = getBytes();
        var SIZE = end - start
        if (end > this.blockSize)
            throw RuntimeException("WARNING: BlockImpl.getBytes(): read position > block size:$SIZE$start")

        val ret = ByteArray(SIZE)
        val capcheck = byteBuffer!!.capacity()
        // TODO: track why this is occurring
        if (capcheck <= SIZE)
            originalPos = 0
        // CAN HAPPEN ON OUT-OF-SPEC FILES whom have last block size < 512
        if (capcheck < SIZE + originalPos)
            SIZE = capcheck - originalPos
        try {
            start += originalPos
            byteBuffer!!.position(start)
            byteBuffer!!.get(ret, 0, SIZE)
        } catch (e: Exception) {
            Logger.logWarn("BlockImpl.getBytes() start: $start size: $SIZE: $e")
        }

        return ret
    }

    /**
     * set the data bytes  on this Block
     */
    override fun setBytes(b: ByteBuffer) {
        byteBuffer = b
        //if(DEBUG) {
        // Logger.logInfo("Debugging turned on in BlockImpl.setBytes");
        //delbytes = getBytes();
        //}
    }

    /**
     * link to the vector of blocks for the storage
     */
    override fun setBlockVector(v: MutableList<*>) {
        blockvec = v
    }

    override fun setNextBlock(b: Block) {
        nextblock = b
        /*		if (LEOFile.DEBUG)
			Logger.logInfo(
				"INFO: BlockImpl setNextBlock(): " + b.toString());*/
    }

    override fun hasNext(): Boolean {
        return nextblock != null
    }

    override fun next(): Any? {
        return nextblock
    }

    override fun remove() {
        this.storage!!.removeBlock(this)
        nextblock = null
    }

    /**
     * init the BIGBLOCK Data
     */
    override fun init(d: ByteBuffer, origidx: Int, origp: Int) {
        originalIdx = origidx
        originalPos = origp
        this.setBytes(d)
    }

    /**
     * true if ths block is represents an extra DIFAT sector
     *
     * @param b
     */
    override fun setIsExtraSector(b: Boolean) {
        isXBAT = b
    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = 4833713921208834278L
    }

}