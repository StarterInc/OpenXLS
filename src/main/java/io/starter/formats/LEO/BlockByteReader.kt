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

import io.starter.formats.XLS.WorkBookException
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.ByteArrayOutputStream
import java.io.Serializable
import java.nio.*
import java.util.AbstractList
import java.util.ArrayList

/**
 * Provide a translation layer between the block vector and a byte array.
 *
 *
 *
 *
 * A record (Storage, XLSRecord) can retrieve bytes from Scattered blocks
 * transparently.
 *
 *
 * The BlockByteReader Allocates a ByteBuffer containing only byte references
 * contained within Blocs assigned to the implementation class.
 *
 *
 * A Class using this reader will either subclass the reader and manage its
 * blocks, or interact with a shared reader.
 *
 *
 * In the case of a shared reader,
 */
open class BlockByteReader : Serializable {
    /**
     * @return
     */
    /**
     * Only add the offset when we are fetching
     * data from 'within' a record, ie: rkdata.get(i)
     * as opposed to when we are traversing the entire collection
     * of bytes as in WorkBookFactory.parse()...
     *
     *
     * * @param b
     */
    var applyRelativePosition = true

    /**Assign blocks to the recs
     *
     * @param rec
     * @param startpos
     * @return
     */
    /* Not used
	public int setBlocks(BlockByteConsumer rec, int startpos) {
		if (rec.getBlocks() != null)
			return rec.getFirstBlock();

		int reclen = rec.getLength();

		int numblocks = -1, firstblock = -1, lastblock = -1;

		// Handle assignment of blocks
		//	span blocks on either side of the block boundary
		if (this.getApplyRelativePosition()) {
			firstblock = rec.getOffset();
			firstblock /= BIGBLOCK.SIZE;
			lastblock = reclen + rec.getOffset();
			lastblock /= BIGBLOCK.SIZE;
		} else {
			firstblock = startpos;
			firstblock /= BIGBLOCK.SIZE;
			lastblock = (reclen + startpos);
			lastblock /= BIGBLOCK.SIZE;
		}

 		numblocks = lastblock - (firstblock - 1);

		if (numblocks < 1)
			numblocks = 1;

		Block[] blks = new Block[numblocks];
		for (int t = 0; t < numblocks; t++) {
			try { // inlining byte read
				int getblk = firstblock + t;
				//Logger.logInfo("GETTING BLOCK: "+ getblk);
				blks[t] = (Block) this.getBlockmap().get(getblk);
			} catch (Exception a) {
				Logger.logWarn(
					"ERROR: Bytes for block:"
						+ blks[t]
						+ " failed: "
						+ a
						+ " Output Corrupted.");
			}
		}
		rec.setFirstBlock(firstblock);
		rec.setLastBlock(lastblock);
		rec.setBlocks(blks);
		return firstblock;
	}
*/

    /**
     * For whatever reason, get the blockmap
     *
     * @return the map of blocks
     */
    var blockmap: List<*> = ArrayList()
        private set
    val isReadOnly = false
    /**
     * @return
     */
    /**
     * @param i
     */
    var length = -1
    /**
     * @return
     */
    /**
     * @param buffer
     */
    @Transient
    var backingByteBuffer = ByteBuffer.allocate(0)

    /**
     * @return
     */
    val char: Char
        get() = backingByteBuffer.char

    /**
     * @return
     */
    val double: Double
        get() = backingByteBuffer.double

    /**
     * @return
     */
    val float: Float
        get() = backingByteBuffer.float

    /**
     * @return
     */
    val int: Int
        get() = backingByteBuffer.int

    /**
     * @return
     */
    val long: Long
        get() = backingByteBuffer.long

    /**
     * @return
     */
    val short: Short
        get() = backingByteBuffer.short

    /**
     * @return
     */
    val isDirect: Boolean
        get() = backingByteBuffer.isDirect

    protected constructor() {
        // empty constructor...
    }

    constructor(blokz: List<*>, len: Int) {
        this.blockmap = blokz
        this.length = len
    }

    /**
     * Allows for getting of header bytes without setting blocks on a rec
     *
     * @param startpos
     * @return
     */
    fun getHeaderBytes(startpos: Int): ByteArray {
        var startpos = startpos
        try {
            var SIZE = BIGBLOCK.SIZE // normal case
            if (this.length < StorageTable.BIGSTORAGE_SIZE) {
                SIZE = SMALLBLOCK.SIZE
            }
            val block = startpos / SIZE
            val check = startpos % SIZE
            // handle EOF that falls right on boundary
            if (check + 4 > SIZE && blockmap.size - 1 == block) {
                // Last EOF falls within 4 bytes of 512 boundary... junkrec
                return byteArrayOf(0x0, 0x0, 0x0, 0x0)
            } else if (check + 4 > SIZE) { // SPANNER!
                var bx = this.blockmap[block] as Block
                var l1 = SIZE * (block + 1) - startpos
                val s2 = startpos % SIZE
                val b1 = bx.getBytes(s2, s2 + l1)
                bx = this.blockmap[block + 1] as Block
                l1 = 4 - l1
                val b2 = bx.getBytes(0, l1)
                return ByteTools.append(b2, b1)
            }

            val bx = this.blockmap[block] as Block
            startpos -= block * SIZE
            return bx.getBytes(startpos, startpos + 4)
        } catch (e: RuntimeException) {
            throw WorkBookException("Smallblock based workbooks are unsupported in OpenXLS: see http://extentech.com/uimodules/docs/docs_detail.jsp?showall=true&meme_id=195", WorkBookException.SMALLBLOCK_FILE)
        }

    }


    /* Return the byte from the blocks at the proper locations...
     *
     * @see java.nio.ByteBuffer#get()
     */
    operator fun get(rec: BlockByteConsumer, startpos: Int): Byte {
        return get(rec, startpos, 1)[0]
    }

    /* Return the bytes from the blocks at the proper locations...
     *
     * as opposed to when we are traversing the entire collection of bytes as in WorkBookFactory.parse()
     *
     * @ see java.nio.ByteBuffer # get()
     */
    operator fun get(rec: BlockByteConsumer, startpos: Int, len: Int): ByteArray {
        var startpos = startpos
        rec.byteReader = this
        //	we only want to add the offset when
        //  we are fetching data from 'within' a record, ie: rkdata.get(i)
        val recoffy = rec.offset
        if (this.applyRelativePosition) {

            startpos += 4 // add the offset
        }
        startpos += recoffy
        // reality checks
        if (false)
        // ((startpos + len) > getLength())
            Logger.logWarn(
                    "WARNING: BlockByteReader.get(rec,"
                            + startpos
                            + ","
                            + rec.length
                            + ") error.  Attempt to read past end of Block buffer.")

        // return the bytes from the rec
        return getRecBytes(rec, startpos, len)
    }

    /* Handles the spanning of Record Bytes over Block boundaries
     * then returns requested bytes.
     *
     */
    private fun getRecBytes(
            rec: BlockByteConsumer,
            startpos: Int,
            len: Int): ByteArray {
        if (startpos < 0 || startpos > startpos + len)
            throw RuntimeException(
                    "ERROR: BBR.getRecBytes("
                            + rec.javaClass.getName()
                            + ","
                            + startpos
                            + ","
                            + (startpos + len)
                            + ") failed - OUT OF BOUNDS.")
        // get the block byte boundaries
        val pos = this.getReadPositions(startpos, len)
        val numblocks = pos.size / 3
        var blkdef = 0
        //	backingByteBuffer = blokx.getByteBuffer();
        //	Temporarily use BAOS...
        val out = ByteArrayOutputStream()
        for (t in 0 until numblocks) {
            try { // inlining byte read
                val b1 = this.blockmap[pos[blkdef++]] as Block
                out.write(b1.getBytes(pos[blkdef++], pos[blkdef++]))
                if (false)
                    Logger.logInfo(
                            "INFO: BBR.getRecBytes() "
                                    + rec.javaClass.getName()
                                    + " ACCESSING DATA for block:"
                                    + b1.blockIndex
                                    + ":"
                                    + pos[0]
                                    + "-"
                                    + pos[1])
            } catch (a: Exception) {
                Logger.logWarn(
                        "ERROR: BBR.getRecBytes streaming $rec bytes for block failed: $a")
            }

        }
        return out.toByteArray()
    }

    /**
     * Gets the list of blocks needed to read the given sequence.
     */
    fun getReadPositions(startpos: Int, len: Int): IntArray {
        return getReadPositions(startpos, len,
                this.length >= StorageTable.BIGSTORAGE_SIZE)
    }

    /**
     * Gets the mapping from stream offsets to file offsets over the given
     * range.
     *
     *
     * This returns an array of integers arranged in pairs. The first value of
     * each pair is an offset from `start` and the second is the
     * corresponding offset in the source file.
     */
    fun getFileOffsets(start: Int, size: Int): IntArray {
        val smap = this.getReadPositions(start, size)
        val fmap = IntArray(smap.size / 3 * 2)

        var offset = 0
        var fidx = 0
        var block: Block? = null
        var prev: Block?
        var sidx = 0
        while (sidx < smap.size) {
            prev = block
            block = this.blockmap[smap[sidx]] as Block

            if (prev == null || block.originalPos + smap[sidx + 1] != prev.originalPos + smap[sidx - 1]) {
                fmap[fidx++] = offset
                fmap[fidx++] = block.originalPos + smap[sidx + 1]
            }

            offset += smap[sidx + 2] - smap[sidx + 1]
            sidx += 3
        }

        val ret = IntArray(fidx)
        System.arraycopy(fmap, 0, ret, 0, fidx)
        return ret
    }

    fun getFileOffsetString(start: Int, size: Int): String {
        var ret = ""
        val map = this.getFileOffsets(start, size)

        var idx = 0
        while (idx < map.size) {
            ret += ((if (idx != 0) " " else "")
                    + Integer.toHexString(map[idx + 0]).toUpperCase() + ":"
                    + Integer.toHexString(map[idx + 1]).toUpperCase())
            idx += 2
        }

        return ret
    }

    /**
     * Set the map of Blocks contained in this reader
     *
     * @param list the map of blocks
     */
    fun setBlockmap(list: AbstractList<*>) {
        blockmap = list
    }

    /**
     * @return
     */
    fun array(): ByteArray {
        return backingByteBuffer.array()
    }

    /**
     * @return
     */
    fun arrayOffset(): Int {
        return backingByteBuffer.arrayOffset()
    }

    /**
     * @return
     */
    fun asCharBuffer(): CharBuffer {
        return backingByteBuffer.asCharBuffer()
    }

    /**
     * @return
     */
    fun asDoubleBuffer(): DoubleBuffer {
        return backingByteBuffer.asDoubleBuffer()
    }

    /**
     * @return
     */
    fun asFloatBuffer(): FloatBuffer {
        return backingByteBuffer.asFloatBuffer()
    }

    /**
     * @return
     */
    fun asIntBuffer(): IntBuffer {
        return backingByteBuffer.asIntBuffer()
    }

    /**
     * @return
     */
    fun asLongBuffer(): LongBuffer {
        return backingByteBuffer.asLongBuffer()
    }

    /**
     * @return
     */
    fun asReadOnlyBuffer(): ByteBuffer {
        return backingByteBuffer.asReadOnlyBuffer()
    }

    /**
     * @return
     */
    fun asShortBuffer(): ShortBuffer {
        return backingByteBuffer.asShortBuffer()
    }

    /**
     * @return
     */
    fun capacity(): Int {
        return backingByteBuffer.capacity()
    }

    /**
     * @return
     */
    fun clear(): Buffer {
        return backingByteBuffer.clear()
    }

    /**
     * @return
     */
    fun compact(): ByteBuffer {
        return backingByteBuffer.compact()
    }

    /**
     * @param arg0
     * @return public int compareTo(Object arg0) {
     * return backingByteBuffer.compareTo(arg0);
     * }
     */

    /**
     * @return
     */
    fun duplicate(): ByteBuffer {
        return backingByteBuffer.duplicate()
    }

    /* (non-Javadoc)
     * @see java.lang.Object#equals(java.lang.Object)
     */
    override fun equals(arg0: Any?): Boolean {
        return backingByteBuffer == arg0
    }

    /**
     * @return
     */
    fun flip(): Buffer {
        return backingByteBuffer.flip()
    }

    /**
     * @return
     */
    fun get(): Byte {
        return backingByteBuffer.get()
    }

    /**
     * @param arg0
     * @return
     */
    operator fun get(arg0: ByteArray): ByteBuffer {
        return backingByteBuffer.get(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @param arg2
     * @return
     */
    operator fun get(arg0: ByteArray, arg1: Int, arg2: Int): ByteBuffer {
        return backingByteBuffer.get(arg0, arg1, arg2)
    }

    /**
     * @param arg0
     * @return
     */
    fun getChar(arg0: Int): Char {
        return backingByteBuffer.getChar(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun getDouble(arg0: Int): Double {
        return backingByteBuffer.getDouble(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun getFloat(arg0: Int): Float {
        return backingByteBuffer.getFloat(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun getInt(arg0: Int): Int {
        return backingByteBuffer.getInt(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun getLong(arg0: Int): Long {
        return backingByteBuffer.getLong(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun getShort(arg0: Int): Short {
        return backingByteBuffer.getShort(arg0)
    }

    /**
     * @return
     */
    fun hasArray(): Boolean {
        return backingByteBuffer.hasArray()
    }

    /* (non-Javadoc)
     * @see java.lang.Object#hashCode()
     */
    override fun hashCode(): Int {
        return backingByteBuffer.hashCode()
    }

    /**
     * @return
     */
    fun hasRemaining(): Boolean {
        return backingByteBuffer.hasRemaining()
    }

    /**
     * @return
     */
    fun limit(): Int {
        return backingByteBuffer.limit()
    }

    /**
     * @param arg0
     * @return
     */
    fun limit(arg0: Int): Buffer {
        return backingByteBuffer.limit(arg0)
    }

    /**
     * @return
     */
    fun mark(): Buffer {
        return backingByteBuffer.mark()
    }

    /**
     * @return
     */
    fun order(): ByteOrder {
        return backingByteBuffer.order()
    }

    /**
     * @param arg0
     * @return
     */
    fun order(arg0: ByteOrder): ByteBuffer {
        return backingByteBuffer.order(arg0)
    }

    /**
     * @return
     */
    fun position(): Int {
        return backingByteBuffer.position()
    }

    /**
     * @param arg0
     * @return
     */
    fun position(arg0: Int): Buffer {
        return backingByteBuffer.position(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun put(arg0: Byte): ByteBuffer {
        return backingByteBuffer.put(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun put(arg0: ByteArray): ByteBuffer {
        return backingByteBuffer.put(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @param arg2
     * @return
     */
    fun put(arg0: ByteArray, arg1: Int, arg2: Int): ByteBuffer {
        return backingByteBuffer.put(arg0, arg1, arg2)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun put(arg0: Int, arg1: Byte): ByteBuffer {
        return backingByteBuffer.put(arg0, arg1)
    }

    /**
     * @param arg0
     * @return
     */
    fun put(arg0: ByteBuffer): ByteBuffer {
        return backingByteBuffer.put(arg0)
    }

    /**
     * @param arg0
     * @return
     */
    fun putChar(arg0: Char): ByteBuffer {
        return backingByteBuffer.putChar(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun putChar(arg0: Int, arg1: Char): ByteBuffer {
        return backingByteBuffer.putChar(arg0, arg1)
    }

    /**
     * @param arg0
     * @return
     */
    fun putDouble(arg0: Double): ByteBuffer {
        return backingByteBuffer.putDouble(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun putDouble(arg0: Int, arg1: Double): ByteBuffer {
        return backingByteBuffer.putDouble(arg0, arg1)
    }

    /**
     * @param arg0
     * @return
     */
    fun putFloat(arg0: Float): ByteBuffer {
        return backingByteBuffer.putFloat(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun putFloat(arg0: Int, arg1: Float): ByteBuffer {
        return backingByteBuffer.putFloat(arg0, arg1)
    }

    /**
     * @param arg0
     * @return
     */
    fun putInt(arg0: Int): ByteBuffer {
        return backingByteBuffer.putInt(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun putInt(arg0: Int, arg1: Int): ByteBuffer {
        return backingByteBuffer.putInt(arg0, arg1)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun putLong(arg0: Int, arg1: Long): ByteBuffer {
        return backingByteBuffer.putLong(arg0, arg1)
    }

    /**
     * @param arg0
     * @return
     */
    fun putLong(arg0: Long): ByteBuffer {
        return backingByteBuffer.putLong(arg0)
    }

    /**
     * @param arg0
     * @param arg1
     * @return
     */
    fun putShort(arg0: Int, arg1: Short): ByteBuffer {
        return backingByteBuffer.putShort(arg0, arg1)
    }

    /**
     * @param arg0
     * @return
     */
    fun putShort(arg0: Short): ByteBuffer {
        return backingByteBuffer.putShort(arg0)
    }

    /**
     * @return
     */
    fun remaining(): Int {
        return backingByteBuffer.remaining()
    }

    /**
     * @return
     */
    fun reset(): Buffer {
        return backingByteBuffer.reset()
    }

    /**
     * @return
     */
    fun rewind(): Buffer {
        return backingByteBuffer.rewind()
    }

    /**
     * @return
     */
    fun slice(): ByteBuffer {
        return backingByteBuffer.slice()
    }

    /* (non-Javadoc)
     * @see java.lang.Object#toString()
     */
    override fun toString(): String {
        return backingByteBuffer.toString()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = 4845306509411520019L


        /**
         * returns the lengths of the two byte
         * <pre>
         * ie: start 10 len 8
         * blk0 = 10-18
         * 0,10,18
         *
         * ie: start 514 len 1
         * blk1 = 2-3
         * 1,2,3
         *
         * ie: start 480 len 1124
         *
         * blk0 = 480-512
         * blk1 = 512-1024
         * blk2 = 1024-1536
         * blk3 = 1536-1604
         *
         * [0, 480, 512,1, 0, 512, 2, 0, 512, 3, 0, 68]
         *
         * 8139, 693
        </pre> *
         *
         * @param startpos
         * @return
         */
        fun getReadPositions(startpos: Int, len: Int, BIGBLOCKSTORAGE: Boolean): IntArray {
            var len = len
            //	Logger.logInfo("BBR.getReadPositions()"+startpos+":"+endpos);

            // 20100323 KSC: handle small blocks
            var SIZE = BIGBLOCK.SIZE // normal case
            if (!BIGBLOCKSTORAGE)
                SIZE = SMALLBLOCK.SIZE


            var firstblock = startpos / SIZE
            val lastblock = (startpos + len) / SIZE
            var numblocks = lastblock - firstblock
            numblocks++
            val origlen = len

            var ct = startpos / SIZE


            var pos1 = startpos
            var ret = IntArray(numblocks * 3)
            var t = 0
            // for each block, create 2 byte positions
            while (len > 0) {
                if (t >= ret.size) {
                    Logger.logWarn("BlockByteReader.getReadPositions() wrong guess on NumBlocks.")
                    numblocks++
                    val retz = IntArray(numblocks * 3)
                    System.arraycopy(ret, 0, retz, 0, ret.size)
                    ret = retz
                }
                ret[t++] = firstblock++
                var check = pos1 % SIZE //leftover
                check += len
                if (check > SIZE) { // SPANNER!
                    pos1 = startpos - SIZE * ct
                    if (pos1 < 0) pos1 = 0
                    // int s1 = pos1- ((SIZE)*(ct));
                    var s2 = pos1 % SIZE
                    if (s2 < 0) {
                        s2 = 0
                        pos1 = 0
                    }
                    ret[t++] = s2
                    ret[t++] = SIZE
                } else {
                    pos1 = startpos - SIZE * ct
                    val strt = startpos - SIZE * ct
                    if (strt < 0) {
                        ret[t++] = 0
                        ret[t++] = len
                    } else {
                        ret[t++] = strt
                        ret[t++] = startpos + origlen - SIZE * ct
                    }

                }
                ct++
                val ctdn = ret[t - 1] - pos1
                len -= ctdn
                pos1 = 0//startpos;
            }
            return ret
        }

        /**
         * @param arg0
         * @return
         */
        fun allocate(arg0: Int): ByteBuffer {
            return ByteBuffer.allocate(arg0)
        }

        /**
         * @param arg0
         * @return
         */
        fun allocateDirect(arg0: Int): ByteBuffer {
            return ByteBuffer.allocateDirect(arg0)
        }

        /**
         * @param arg0
         * @return
         */
        fun wrap(arg0: ByteArray): ByteBuffer {
            return ByteBuffer.wrap(arg0)
        }

        /**
         * @param arg0
         * @param arg1
         * @param arg2
         * @return
         */
        fun wrap(arg0: ByteArray, arg1: Int, arg2: Int): ByteBuffer {
            return ByteBuffer.wrap(arg0, arg1, arg2)
        }
    }

}
