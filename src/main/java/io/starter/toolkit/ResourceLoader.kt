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

import io.starter.OpenXLS.GetInfo
import io.starter.naming.InitialContextImpl

import java.io.*
import java.lang.reflect.Method
import java.net.URL
import java.net.URLDecoder
import java.util.Enumeration
import java.util.MissingResourceException
import java.util.Properties
import java.util.zip.ZipEntry
import java.util.zip.ZipInputStream
import java.util.zip.ZipOutputStream

/**
 * Resource Loader which implements a basic JNDI Context and performs:
 *
 *  * Classloading mapped to variable names in properties files allows for easy
 * abstraction of implementation classes
 *  * Configuration strings loaded from properties files
 *  * Arbitrary resource binding
 */
class ResourceLoader : InitialContextImpl, Serializable, javax.naming.Context {
    private var resloc = ""
    private var propsfile: File? = null
    private val resources = Properties()

    val keys: Enumeration<*>
        get() = if (!snagged) resources.keys() else env.keys()

    private var snagged = false

    override fun toString(): String {
        return "Extentech ResourceLoader v." + ResourceLoader.version
    }

    fun getObject(key: String): Any {
        return if (!snagged) resources[key] else env[key]
    }

    /**
     * put the properties file vals in the ResourceLoader
     */
    private fun snagVals() {
        snagged = true
        val a = resources.keys()
        while (a.hasMoreElements()) {
            val mystr = a.nextElement() as String
            env[mystr] = resources[mystr]
        }
    }

    /**
     * Constructor which takes a path to the properties file containing the initial
     * ResourceLoader values.
     *
     *
     * Uses the resources from the proper locale.
     *
     * @param s
     */
    constructor(s: String) : super() {
        var s = s
        if (true) {
            if (s.indexOf("resources/") == -1)
                s = "resources/$s" // StringTool.strip(s,".properties");
        }
        Logger.logInfo("ResourceLoader INIT: $s")

        resloc = s
        try {
            try {
                propsfile = File("$s.properties")
                val fis = FileInputStream(propsfile!!)
                resources.load(fis)
            } catch (e: Exception) {
                try {
                    // propsfile.mkdirs();

                    propsfile!!.createNewFile()
                    val fis = FileInputStream(propsfile!!)
                    resources.load(fis)
                } catch (ex: Exception) {
                    Logger.logWarn("Could not init Resourceloader from: " + propsfile!!.absolutePath)
                }

            }

            // handle private values
            var hidevals = false
            try {
                if (resources["public"] != null)
                    if (resources["visibility"] == "private")
                        hidevals = true
            } catch (mre: MissingResourceException) { // do not load properties into JNDI environment
                // Logger.logInfo("ResourceLoader - getting resources failed: " +
                // mre.toString());
            }

            if (!hidevals)
                this.snagVals()
        } catch (mre: MissingResourceException) {
            Logger.logErr("ResourceLoader getting resources failed: $mre")
        }

    }

    constructor() : super() {
        // empty constructor
    }

    /**
     * Returns a String from the properties file
     *
     * @param nm
     * @return
     */
    fun getResourceString(nm: String): String? {
        var str: String
        try {
            str = resources[nm].toString()
        } catch (mre: Exception) {
            str = ""
            // Logger.logWarn("Resource string: " + nm + " not found in: " + this.resloc);
        }

        return str
    }

    /**
     * Sets a String value in the properties file
     *
     * @param nm
     * @return
     */
    fun setResourceString(nm: String, v: String) {
        try {
            resources.setProperty(nm, v)
            val fos = FileOutputStream(propsfile!!)
            resources.store(fos, null)
            fos.flush()
            fos.close()
        } catch (mre: Exception) {
            Logger.logWarn("Resource string: " + nm + " could not be set to " + v + " in:" + this.resloc)
        }

    }

    /**
     * Returns an Array of Objects which are class loaded based on a comma-delimited
     * list of class names listed in the properties file.
     */
    fun getObjects(propname: String): Array<Any>? {
        val objnames = getResourceString(propname)
        if (objnames != null) {
            val obj = arrayOfNulls<Any>(1)
            obj[0] = loadClass(objnames)
            return obj
        }
        return null
    }

    /**
     * Sets the debugging level for the ResourceLoader
     *
     * @param b
     */
    fun setDebug(b: Boolean) {
        DEBUG = b
    }

    companion object {

        /**
         *
         */
        private const val serialVersionUID = 12345245254L
        var DEBUG = false

        val version: String
            get() = GetInfo.getVersion()

        /**
         * Load a Class by name
         *
         * @param className
         * @return
         */
        fun loadClass(className: String): Any? {

            val cl = ExtenClassLoader()

            // Make a new one of whatever type of Obj it is...
            var obj: Any? = null
            try {
                val c = cl.loadClass(className, true)
                obj = c!!.newInstance()
                return obj
            } catch (t: ClassFormatError) {
                Logger.logErr(t.toString())
                return null
            } catch (t: ClassNotFoundException) {
                Logger.logErr(t)
                return null
            } catch (t: ClassCastException) {
                Logger.logErr(t)
                return null
            } catch (t: InstantiationException) {
                Logger.logErr(t)
                return null
            } catch (t: IllegalAccessException) {
                Logger.logErr(t)
                return null
            }

        }

        /**
         * Loads the named resource from the class path.
         */
        @Throws(IOException::class)
        fun getBytesFromJar(name: String): ByteArray? {

            val classLoader = Thread.currentThread().contextClassLoader
            val source = classLoader.getResourceAsStream("io/starter/OpenXLS/templates/prototype.ser") ?: return null

            val sink = ByteArrayOutputStream()
            val buffer = ByteArray(1024)
            var count = 0

            while (count != -1) {
                sink.write(buffer, 0, count)
                count = source.read(buffer)
            }

            source.close()
            return sink.toByteArray()
        }

        /**
         * Returns the file system-specific path to a given resource in the classpath
         * for the VM.
         *
         * @param resource
         * @return
         */
        fun getFilePathForResource(resource: String): String? {
            val u = ResourceLoader().javaClass.getResource(resource)
            // 20070107 KSC: report error
            if (u == null) {
                Logger.logErr("ResourceLoader.getFilePathForResource: $resource not found.")
                return null
            }
            if (DEBUG)
                Logger.logInfo("ResourceLoader.getFilePathForResource() got:" + u!!.toString())
            var s: String? = u!!.getFile()

            if (DEBUG)
                Logger.logInfo("ResourceLoader.getFilePathForResource Decoding:" + s!!)
            s = ResourceLoader.Decode(s)
            if (DEBUG)
                Logger.logInfo("ResourceLoader.getFilePathForResource Decoded:" + s!!)

            val i = s!!.indexOf("!")
            if (i > -1) { // file is in a jar
                var zipstring = s.substring(0, i)

                // cut off the internal zip file part & the file:/
                var begin = zipstring.indexOf(":")
                begin += 1
                zipstring = zipstring.substring(begin)
                if (zipstring.indexOf(":") != -1) { // windoze box
                    if (zipstring.indexOf("/") == 0) {
                        zipstring = zipstring.substring(1)
                    }
                }
                if (DEBUG)
                    Logger.logInfo("Resourceloader.getFilePathForResource(): Successfully obtained $zipstring")
                return zipstring
            } else { // file is not in a jar
                if (DEBUG)
                    Logger.logErr("ResourceLoader.getFilePathForResource(): File is not in jar:$s")
                return s
            }
        }

        /**
         * write file f to jar referenced by jarandresource ( <jar><resource> ) and set
         * path/name to resource
         *
         * @param jarandResource
         * @param f
        </resource></jar> */
        fun addFileToJar(jarandResource: String, f: String) {
            val tmp = extractJarAndResourceName(jarandResource)
            try {
                val out = ZipOutputStream(FileOutputStream(tmp[0])) // open Archive for outptu
                val fin = ZipInputStream(FileInputStream(f))
                out.putNextEntry(ZipEntry(tmp[1]))
                val buf = ByteArray(fin.available())
                fin.read(buf)
                out.write(buf)
                out.flush()
                out.closeEntry()
                out.close()
            } catch (e: Exception) {
                Logger.logErr("addFileToJar: Jar: " + tmp[0] + " File: " + tmp[1] + " : " + e.toString())
            }

        }

        /**
         * returns truth of "file is a jar/archive file"
         *
         * @param f
         * @return
         */
        fun isJarFile(f: String): Boolean {
            var f = f
            f = f.toLowerCase()
            var i = f.indexOf(".war")
            if (i < 0)
                i = f.indexOf(".jar")
            if (i < 0)
                i = f.indexOf(".rar")
            if (i < 0)
                i = f.indexOf(".zip")
            return i > -1
        }

        /**
         * separate and return the jar portion and resource portion of a jar and
         * resource string: <jar (.war></jar>/.zip/.jar/.rar)><resource>
         *
         * @param jarAndResource
         * @return String[2]
        </resource> */
        fun extractJarAndResourceName(jarAndResource: String): Array<String> {
            var jarAndResource = jarAndResource
            jarAndResource = jarAndResource.toLowerCase()
            var i = jarAndResource.indexOf(".war")
            if (i < 0)
                i = jarAndResource.indexOf(".jar")
            if (i < 0)
                i = jarAndResource.indexOf(".rar")
            if (i < 0)
                i = jarAndResource.indexOf(".zip")
            return arrayOf(jarAndResource.substring(0, i + 4), jarAndResource.substring(i + 5))
        }

        /**
         * Get the path to a directory by locating the jar file in the classpath
         * containing the given resource name.
         *
         * @param resource
         * @return
         */
        fun getWorkingDirectoryFromJar(resource: String): String {
            var s: String
            // if jarloc property is set, use it to find working directory
            if (System.getProperty("io.starter.OpenXLS.jarloc") != null) {
                s = System.getProperty("io.starter.OpenXLS.jarloc") + "!"
            } else {
                val u = ResourceLoader::class.java!!.getResource(resource)
                s = u.getFile()
            }
            if (DEBUG)
                Logger.logInfo("Resource: $resource found in: $s")

            // cut off the internal zip file part & the file:/
            var begin = -1
            begin = s.indexOf("file:")
            if (begin < 0) {
                begin = s.indexOf(":")
                begin += 1
            } else
                begin += 5
            s = s.substring(begin)
            if (s.indexOf(":") != -1) { // windoze box
                if (s.indexOf("/") == 0)
                    s = s.substring(1)
            }
            if (DEBUG)
                Logger.logInfo("ResourceLoader() after stripping:$s")
            var i = s.indexOf("!")
            if (i > -1) {
                var zipstring = s.substring(0, i)
                i = zipstring.lastIndexOf("/")
                if (i == -1) {
                    i = zipstring.lastIndexOf("\\")
                }
                zipstring = zipstring.substring(0, i)
                if (DEBUG)
                    Logger.logInfo("ResourceLoader() returning zipstring Final Working Directory Setting: $zipstring")
                return zipstring
            } else {
                if (DEBUG)
                    Logger.logInfo("ResourceLoader() returning Final Working Directory Setting: $s")
                return s
            }
        }

        private val decodr = URLDecoder()

        /**
         * Decode a URL String, if supported by the JDK version in use this method will
         * utilize the
         *
         * @param s
         * @return
         */
        fun Decode(s: String?): String? {
            val tmpstr = arrayOf<String>(s, "ISO-8859-1")
            // first attempt using encoding...
            var ret = s
            ret = ResourceLoader.executeIfSupported(decodr, tmpstr, "decode") as String?
            if (ret == null)
                try {
                    ret = URLDecoder.decode(s!!, "ISO-8859-1")
                } catch (e: Exception) {
                    Logger.logErr("ResourceLoader.Decode resource failed: $e")
                }

            return ret
        }

        /**
         * Decode a URL String, if supported by the JDK version in use this method will
         * utilize the non-deprecated method of decoding.
         *
         * @param s,        string to decode
         * @param encoding, the encoding type to use
         * @return
         */
        fun Decode(s: String, encoding: String): String? {
            val tmpstr = arrayOf(s, "Encoding")
            // first attempt using encoding...
            var ret: String? = s
            ret = ResourceLoader.executeIfSupported(decodr, tmpstr, "decode") as String?
            if (ret == null)
                try {
                    ret = URLDecoder.decode(s)
                } catch (e: Exception) {
                    Logger.logErr("ResourceLoader.Decode resource failed: $e")
                }

            return ret
        }

        /**
         * Attempt to execute a Method on an Object
         *
         * @param ob       the Object which contains the method you want to execute
         * @param args     an array of arguments to the Method, null if none
         * @param methname the name of the Method you are executing
         * @return the return value of the method if any
         */
        fun executeIfSupported(ob: Any, args: Array<Any>?, methname: String): Any? {
            try {
                var retob: Any? = null
                val mt = ob.javaClass.getMethods()
                var t = 0
                while (t < mt.size) {
                    // Make JDK-specific call to method
                    // Logger.logInfo(mt[t].getName());
                    var numparms = 0
                    var numargs = 0
                    if (args != null)
                        numargs = args.size
                    if (mt[t].getParameterTypes() != null)
                        numparms = mt[t].getParameterTypes().size
                    val nm = mt[t].getName()
                    if (nm == methname && numparms == numargs) {
                        try {
                            val mx = mt[t]
                            retob = mx.invoke(ob, *args)
                            break
                        } catch (e: Exception) {
                            if (false)
                                Logger.logWarn("ResourceLoader.executeIfSupported() Method NOT supported: " + methname
                                        + " in " + ob.javaClass.getName() + " for arguments "
                                        + StringTool.arrayToString(args!!))
                            return null
                        }

                    }
                    t++
                }
                if (false)
                    if (t == mt.size)
                        Logger.logWarn("ResourceLoader.executeIfSupported() Method NOT found: " + methname + " in "
                                + ob.javaClass.getName() + " for arguments " + StringTool.arrayToString(args!!))
                return retob
            } catch (e: NoSuchMethodError) {
                return null
            }

        }

        /**
         * Execute a Method on an Object
         *
         * @param ob       the Object which contains the method you want to execute
         * @param args     an array of arguments to the Method, null if none
         * @param methname the name of the Method you are executing
         * @return the return value of the method if any
         */
        @Throws(Exception::class)
        fun execute(ob: Any, args: Array<Any>, methname: String): Any? {
            val pc = arrayOfNulls<Class<*>>(args.size)
            for (r in args.indices) {
                pc[r] = args[r].javaClass
            }
            var mt: Method? = null
            try {
                mt = ob.javaClass.getMethod(methname, *pc)
            } catch (e: NoSuchMethodException) {
                // deal with 'unwrapping' primitives
                return executeIfSupported(ob, args, methname)
            }

            // Logger.logInfo("ResourceLoader.execute() Invoking:" + mt.getName() +" on " +
            // ob.getClass().getName());
            try {
                return mt!!.invoke(ob, *args)
            } catch (e: Exception) {
                Logger.logErr("ResourceLoader.execute " + methname + " on " + ob.javaClass.getName() + " failed: "
                        + e.toString())
                e.printStackTrace()
                return null
            }

        }
    }
}