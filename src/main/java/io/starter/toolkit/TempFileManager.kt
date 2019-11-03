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
/**
 * TempFileManager.java
 *
 *
 *
 *
 * Feb 27, 2012
 */
package io.starter.toolkit

import io.starter.OpenXLS.DocumentHandle

import java.io.*

/**
 * The TempFileManager allows for consolidated handling of all TempFiles used by OpenXLS.
 *
 * TempFileManager is pluggable and allows you to implement a custom TempFileGenerator and install using
 * System properties.
 *
 * ie:
 *
 * System.setProperty(TempFileManager.TEMPFILE_MANAGER_CLASSNAME, "com.acme.CustomTempFileGenerator");
 * WorkBookHandle bkx = new WorkBookHandle("test.xlsx"); // use custom tempfile generator
 *
 *
 */
object TempFileManager {

    var TEMPFILE_MANAGER_CLASSNAME = "io.starter.OpenXLS.tempfilemanager"

    @Throws(IOException::class)
    fun createTempFile(prefix: String, extension: String): File {
        val tmpfu = System.getProperty(TEMPFILE_MANAGER_CLASSNAME)
        if (tmpfu != null) {
            try {
                val tgen = Class.forName(tmpfu).newInstance() as TempFileGenerator
                return tgen.createTempFile(prefix, extension)
            } catch (e: Exception) {
                Logger.logErr("Could not load custom TempFileGenerator: $tmpfu. Falling back to default TempFileGenerator.")
            }

        }
        return DefaultTempFileGeneratorImpl().createTempFile(prefix, extension)
    }

    /**
     * Write an InputStream to disk, and return as a file handle.
     *
     * @param input
     * @param prefix
     * @param extension
     * @return
     */
    @Throws(IOException::class)
    fun createTempFile(input: InputStream, prefix: String, extension: String): File {
        val tmpfile = TempFileManager.createTempFile(prefix, extension)
        JFileWriter.writeToFile(input, tmpfile)
        return tmpfile
    }

    /**
     * Feb 27, 2012
     * @param string
     * @param string2
     * @param dir
     * @return
     * @throws IOException
     */
    @Throws(IOException::class)
    fun createTempFile(prefix: String, extension: String, dir: File): File {
        var prefix = prefix
        prefix = dir.absolutePath + prefix
        return createTempFile(prefix, extension)
    }

    @Throws(IOException::class)
    fun writeToTempFile(prefix: String, extension: String, doc: DocumentHandle): File {
        val file = createTempFile(prefix, extension)

        val stream = BufferedOutputStream(FileOutputStream(file))
        doc.write(stream, DocumentHandle.FORMAT_NATIVE)

        stream.flush()
        stream.close()

        return file
    }
}
