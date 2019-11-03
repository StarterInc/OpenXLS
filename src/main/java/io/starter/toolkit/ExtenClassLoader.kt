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

import java.io.File
import java.io.FileInputStream
import java.io.IOException

class ExtenClassLoader : java.lang.ClassLoader {
    private val targetClassName: String

    private var wd = ""

    constructor(target: String) {
        targetClassName = target
    }

    constructor() {}

    fun setDirectory(_wd: String) {
        wd = _wd
    }

    protected fun loadClassFromFile(name: String): ByteArray {
        var name = name
        var classBytes: ByteArray? = null
        try {
            var file: File? = null
            var stream: FileInputStream? = null
            //  name = name.substring(name.indexOf(wd)+wd.length()); // strip the working directory
            name = this.wd + "/" + name
            name = StringTool.replaceChars(".", name, "/")
            file = File("$name.class")
            classBytes = ByteArray(file.length().toInt())
            stream = FileInputStream(file)
            stream.read(classBytes)
            stream.close()
        } catch (io: IOException) {
        }

        return classBytes
    }

    @Synchronized
    @Throws(ClassNotFoundException::class)
    override fun loadClass(name: String): Class<*>? {
        return loadClass(name, false)
    }

    @Synchronized
    @Throws(ClassNotFoundException::class)
    public override fun loadClass(name: String, resolve: Boolean): Class<*>? {
        var loadedClass: Class<*>?
        val bytes: ByteArray?
        if (name != targetClassName) {
            try {
                loadedClass = super.findSystemClass(name)
                return loadedClass
            } catch (e: ClassNotFoundException) {
            }

        }
        bytes = loadClassFromFile(name)
        if (bytes == null) {
            throw ClassNotFoundException()
        }
        loadedClass = defineClass(name, bytes, 0, bytes.size)
        if (loadedClass == null) {
            Logger.logInfo("Class cannot be loaded: $name")
            throw ClassFormatError()
        }
        if (resolve) {
            resolveClass(loadedClass)
        }

        return loadedClass
    }
}