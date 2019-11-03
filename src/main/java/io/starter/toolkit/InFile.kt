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

import java.io.*

class InFile @Throws(FileNotFoundException::class)
constructor(filename: String) : DataInputStream(BufferedInputStream(FileInputStream(File(filename)))) {

    internal var sb = StringBuffer()

    /* public InFile(String filename)
    throws FileNotFoundException {
        super(new BufferedInputStream(new FileInputStream(new File(filename))));
    }
    */

    @Throws(FileNotFoundException::class)
    constructor(file: File) : this(file.path) {
    }

    /**
     * Reads File from Disk
     *
     * @param fname path to file
     */
    fun readFile(): String {
        try {
            while (this.available() != 0) {
                sb.append(this.readLine())
            }
        } catch (e: FileNotFoundException) {
            Logger.logInfo("FNF Exception in InFile: $e")
        } catch (e: IOException) {
            Logger.logInfo("IO Exception in InFile: $e")
        }

        return sb.toString()
    }

    companion object {


        /**
         * Gets a byte arrray from a file
         *
         * @param file File the file to get bytes from
         * @return byte[] Returns byte[] array file contents
         */
        @Throws(IOException::class)
        fun getBytesFromFile(file: File): ByteArray {
            val fis = FileInputStream(file)
            val length = file.length()
            val ret = ByteArray(length.toInt())
            var offset = 0
            var numRead = 0
            while (offset < ret.size && (numRead = fis.read(ret, offset, ret.size - offset)) >= 0) {
                offset += numRead
            }
            if (offset < ret.size) {
                throw IOException("Read file failed -- all bytes not retreived. " + file.name)
            }
            fis.close()
            return ret

        }
    }

}

