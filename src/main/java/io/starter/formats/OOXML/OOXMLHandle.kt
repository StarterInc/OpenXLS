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
package io.starter.formats.OOXML

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.CellNotFoundException
import io.starter.formats.XLS.OOXMLReader
import io.starter.formats.XLS.OOXMLWriter
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParserException

import java.io.File
import java.io.FileOutputStream
import java.io.IOException

/**
 *
 */
object OOXMLHandle {

    /**
     * generates OOXML for a workbook
     * Creates the ZIP file and writes all files into proper directory structure
     * Will create either an .xlsx or an .xlsm output, depending upon
     * whether WorkBookHandle bk contains macros
     *
     * @param bk   workbookhandle
     * @param path output filename and path
     */
    @Throws(IOException::class)
    fun getOOXML(bk: WorkBookHandle, path: String) {
        var path = path
        try {
            if (!io.starter.formats.XLS.OOXMLAdapter.hasMacros(bk))
                path = StringTool.replaceExtension(path, ".xlsx")
            else
            // it's a macro-enabled workbook
                path = StringTool.replaceExtension(path, ".xlsm")

            val fout = java.io.File(path)
            val dirs = fout.parentFile
            if (dirs != null && !dirs.exists())
                dirs.mkdirs()
            val oe = OOXMLWriter()
            oe.getOOXML(bk, FileOutputStream(path))
        } catch (e: Exception) {
            throw IOException("Error parsing OOXML file: $e")
        }

    }

    /**
     * OOXML parseNBind - reads in an OOXML (Excel 7) workbook
     *
     * @param bk    WorkBookHandle - workbook to input
     * @param fName OOXML filename (must be a ZIP file in OPC format)
     * @throws XmlPullParserException
     * @throws IOException
     * @throws CellNotFoundException
     */
    @Throws(XmlPullParserException::class, IOException::class, CellNotFoundException::class)
    fun parseNBind(bk: WorkBookHandle, fName: String) {
        val oe = OOXMLReader()
        oe.parseNBind(bk, fName)
    }
}
