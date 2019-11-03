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

import io.starter.toolkit.ByteTools

import java.io.ByteArrayInputStream
import java.io.OutputStream
import java.nio.ByteBuffer
import java.nio.ByteOrder


/**
 * Dutifully produces LEO file Blocks.  Complains not.
 */
object BlockFactory {
    /**
     * get a new, empty Block
     */
    fun getPrototypeBlock(type: Int): Block {
        var retblock: Block? = null
        var dta: ByteBuffer? = null
        when (type) {
            Block.BIG -> {
                dta = ByteBuffer.allocate(BIGBLOCK.SIZE)
                retblock = BIGBLOCK()
            }
            Block.SMALL -> {
                dta = ByteBuffer.allocate(SMALLBLOCK.SIZE)
                retblock = SMALLBLOCK()
            }
        }
        dta!!.order(ByteOrder.LITTLE_ENDIAN)
        //for (int x = 0; x < dta.limit(); x++)
        //	dta.put((byte) - 1);
        retblock!!.setBytes(dta)
        return retblock
    }

    /**
     * transform the byte array into Block records.
     */
    fun getBlocksFromOutputStream(bbuf: OutputStream?, blen: Int, type: Int): Array<Block>? {
        var SIZE = 0
        when (type) {
            Block.BIG -> SIZE = 512
            Block.SMALL -> SIZE = 64
        }
        val sz = LEOFile.getSizeInBlocks(blen, SIZE)

        if (bbuf == null)
            return null
        // int len = (blen-3) / SIZE;
        var len = blen / SIZE
        var pos = 0
        var sizeDiff = 0
        var size = 0

        if (len * SIZE < blen) {
            len++
            sizeDiff = len * SIZE - blen
        }
        val blockarr = arrayOfNulls<Block>(len)
        val ins: ByteArrayInputStream? = null

        // KSC: made a bit simpler upon padding situations ...
        // get ALL blockVect (512 byte chunks of file)
        for (i in 0 until len) {
            val bbd = getPrototypeBlock(type)
            val bb = bbd.bytes
            var filepos = i * SIZE
            if (type == Block.SMALL)
            // smallblocks don't need file offset as they are allocated differently than bigblocks
                filepos = 0

            size = SIZE
            // make simpler:
            if (blen - pos < size)
                size = blen - pos
            ins!!.read(bb, pos, size)
            bbd.init(ByteBuffer.wrap(bb), i, filepos)
            pos += SIZE
            blockarr[i] = bbd
        }
        return blockarr
    }

    /**
     * transform the byte array into Block records.
     */
    fun getBlocksFromByteArray(bbuf: ByteArray?, type: Int): Array<Block>? {
        var bbuf = bbuf
        var SIZE = 0
        when (type) {
            Block.BIG -> SIZE = 512
            Block.SMALL -> SIZE = 64
        }
        if (bbuf == null)
            return null
        //		int sz = LEOFile.getSizeInBlocks(bbuf.length,SIZE);

        var len = bbuf.size / SIZE
        var pos = 0
        var sizeDiff = 0
        var size = 0

        if (len * SIZE < bbuf.size) {
            len++    // PAD - this most usually hits when called from buildSSAT
            sizeDiff = len * SIZE - bbuf.size
            val bb = ByteArray(sizeDiff)
            bbuf = ByteTools.append(bb, bbuf)
        }
        val blockarr = arrayOfNulls<Block>(len)

        var start = 1    // for BigBlocks, skip 1st block, for small
        if (type == Block.SMALL)
            start = 0
        // get ALL blockVect (512 byte chunks of file)
        for (i in 0 until len) {
            val bbd = getPrototypeBlock(type)
            val bb = bbd.bytes

            /*			int filepos = (i + start) * SIZE;
			if (type==Block.SMALL)	// smallblocks don't need file offset as they are allocated differently than bigblocks
				filepos= 0;
*/
            size = SIZE
            // make it simpler:
            if (bbuf!!.size - pos < size)
                size = bbuf.size - pos    // account for leftovers (Block padding)
            /*if (i + start == len) {
				size -= sizeDiff; // account for leftovers (Block padding)
			}*/
            // ExcelTools.benchmark("BlockFactory initting new Block@: " + pos + " sz: " +size + " len:" + bbuf.length);
            System.arraycopy(bbuf, pos, bb, 0, size)
            bbd.init(ByteBuffer.wrap(bb), i, 0)    // THIS IS A BYTEARRAY ALL OFFSETS ARE 0 filepos);
            pos += size
            blockarr[i] = bbd
        }
        return blockarr
    }
}