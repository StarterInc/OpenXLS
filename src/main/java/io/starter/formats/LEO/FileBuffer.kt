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
/*
 * Created on Apr 25, 2005
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package io.starter.formats.LEO

import io.starter.toolkit.TempFileManager

import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.nio.ByteBuffer
import java.nio.ByteOrder
import java.nio.channels.FileChannel

/**
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
/**
 *
 */
class FileBuffer {

    /**
     * @return Returns the buffer.
     */
    /**
     * @param buffer The buffer to set.
     */
    @Transient
    var buffer: ByteBuffer? = null
    private var tempfile: File? = null
    internal var channel: FileChannel? = null
    internal var input: FileInputStream? = null

    @Throws(IOException::class)
    fun close() {
        input!!.close()
        channel!!.close()
        if (tempfile != null) {
            tempfile!!.deleteOnExit()
            tempfile!!.delete()
        }
        tempfile = null
    }

    companion object {

        fun readFile(fpath: String): FileBuffer {
            try {
                val fx0 = File(fpath)
                return readFile(fx0)
            } catch (e: Throwable) {
                throw InvalidFileException("LEO FileBuffer.readFile() failed: $e")
            }

        }

        fun readFile(fx0: File): FileBuffer {
            try {
                val fb = FileBuffer()
                fb.input = FileInputStream(fx0)
                fb.channel = fb.input!!.channel
                val fileLength = fb.channel!!.size().toInt()
                // MappedByteBuffer
                fb.buffer = fb.channel!!.map(FileChannel.MapMode.READ_ONLY, 0, fileLength.toLong())
                fb.buffer!!.order(ByteOrder.LITTLE_ENDIAN)
                return fb
            } catch (e: Throwable) {
                throw InvalidFileException("LEO FileBuffer.readFile() failed: $e")
            }

        }

        //TODO: reimplement temp files and deal with cleanup -jm
        fun readFileUsingTemp(fpath: String): FileBuffer {
            return readFileUsingTemp(File(fpath))
        }

        fun readFileUsingTemp(fx0: File): FileBuffer {
            try {
                val fb = FileBuffer()
                // create Temp file and populate
                fb.tempfile = TempFileManager.createTempFile("LEOFile_", ".tmp")
                fb.tempfile!!.delete()

                val input0 = FileInputStream(fx0)
                val channel0 = input0.channel
                val output0 = FileOutputStream(fb.tempfile!!)
                val channel1 = output0.channel
                channel0.transferTo(0, fx0.length(), channel1)

                channel0.close()
                channel1.close()
                input0.close()
                output0.close()

                fb.input = FileInputStream(fb.tempfile!!)
                fb.channel = fb.input!!.channel
                val fileLength = fb.channel!!.size().toInt()
                // MappedByteBuffer
                fb.buffer = fb.channel!!.map(FileChannel.MapMode.READ_ONLY, 0, fileLength.toLong())
                fb.buffer!!.order(ByteOrder.LITTLE_ENDIAN)
                return fb
            } catch (e: Throwable) {
                throw InvalidFileException("LEO FileBuffer.readFile() failed: $e")
            }

        }
    }
}// TODO Auto-generated constructor stub
