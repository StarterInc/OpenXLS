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
import java.nio.channels.FileChannel
import java.nio.charset.StandardCharsets

/**
 * File utilities.
 */
class JFileWriter {

    internal var path: java.lang.String = ""
    internal var filename: java.lang.String = ""
    internal var data: java.lang.String = ""
    internal var newLine = Character.LINE_SEPARATOR

    fun setPath(p: String) {
        path = p
    }

    fun setFileName(f: String) {
        filename = f
    }

    fun setData(d: String) {
        data = d
    }

    internal fun printErr(err: String) {
        Logger.logInfo("Error in JFileWriter: $err")
        Logger.logWarn("Error in JFileWriter: $err")
    }

    fun writeIt(): Boolean {
        try {
            path += filename
            val SR = StringReader(data)
            val outputFile = File(path)
            val out = FileWriter(outputFile)
            var c: Int
            if (outputFile.length() > 0) {
                return false
            }
            while ((c = SR.read()) != -1) out.write(c)
            out.flush()
            out.close()
        } catch (e: IOException) {
            Logger.logInfo("JFileWriter IO Error : $e")
        }

        return true
    }


    fun writeIt(data: String, filename: String, path: String): Boolean {
        var path = path
        try {
            path += filename
            val SR = StringReader(data)
            val outputFile = File(path)
            val out = FileWriter(outputFile)
            var c: Int
            if (outputFile.length() > 0) {
                return false
            }
            while ((c = SR.read()) != -1) {
                out.write(c)
            }
            out.flush()
            out.close()
        } catch (e: IOException) {
            Logger.logInfo("JFileWriter IO Error : $e")
        }

        return true
    }

    fun readFile(fname: String): String {
        val addTxt = StringBuffer()
        try {
            val d = BufferedReader(FileReader(fname))

            while (d.ready()) addTxt.append(d.readLine())
            d.close()
        } catch (e: Exception) {
            printErr("problem reading file: $e")
        }

        return addTxt.toString()
    }

    fun writeLine(file: String, line: String) {
        var s: String
        try {
            val f = File(file)
            // f.mkdirs();
            val out = FileWriter(f)
            val inStream = DataInputStream(StringBufferInputStream(line))
            while ((s = inStream.readLine()) != null) {
                out.write(s)
                out.write(newLine.toInt())
            }
            out.close()
        } catch (e: FileNotFoundException) {
            printErr(e.toString())
        } catch (e: Exception) {
            printErr(e.toString())
        }

    }

    fun writeLogToFile(fname: String, jta: javax.swing.JTextArea) {
        try {
            val n = OutFile(fname)
            val logText = jta.text
            n.writeBytes(logText)
            jta.text = ""
            n.close()
        } catch (e: FileNotFoundException) {
            printErr(e.toString())
        } catch (e: IOException) {
            printErr(e.toString())
        }

    }

    fun readLog(logFname: String): String {
        var addTxt = ""
        try {
            val n = InFile(logFname)
            while (n.available() != 0) {
                addTxt += n.readLine()
            }
        } catch (e: FileNotFoundException) {
            printErr(e.toString())
        } catch (e: IOException) {
            printErr(e.toString())
        }

        return addTxt += "\r\n"
    }

    companion object {


        /**
         * append text to the end of a text file
         */
        @Synchronized
        fun appendToFile(pth: String, text: String) {
            try {
                val bbuf = text.toByteArray(StandardCharsets.UTF_8)
                var outp = File(pth)

                if (!outp.exists()) {
                    outp.mkdirs()
                    outp.delete()
                    outp = java.io.File(pth)
                }

                val outputFile = RandomAccessFile(outp, "rw")
                outputFile.skipBytes(outputFile.length().toInt())
                var strt = 0
                if (outp.exists()) strt = outputFile.length().toInt()
                outputFile.write(bbuf, 0, bbuf.size)
                outputFile.close()
            } catch (e: Exception) {
                Logger.logInfo("JFileWriter.appendToFile() IO Error : $e")
            }

        }

        /**
         * write the inputstream contents to file
         *
         * @param is
         * @param file
         * @throws IOException
         */
        @Throws(IOException::class)
        fun writeToFile(`is`: InputStream, file: File) {
            val out = DataOutputStream(BufferedOutputStream(FileOutputStream(file)))

            // Transfer bytes from in to out
            val buf = ByteArray(1024)
            var len: Int
            while ((len = `is`.read(buf)) > 0) {
                out.write(buf, 0, len)
            }
            `is`.close()
            out.close()
        }

        @Throws(IOException::class)
        fun copyFile(infile: String, outfile: String) {
            val fx = File(infile)
            copyFile(fx, outfile)

            // this.writeLine(outfile, this.readFile(infile));
        }


        /**
         * Copy method, using FileChannel#transferTo
         * NOTE:  will overwrite existing files
         *
         * @param File source
         * @param File target
         * @throws IOException
         */
        @Throws(IOException::class)
        fun copyFile(source: File, target: String) {
            var fout = File(target)
            fout.mkdirs()
            fout.delete()
            fout = File(target)
            val `in` = FileInputStream(source).channel
            val out = FileOutputStream(target).channel
            `in`.transferTo(0, `in`.size(), out)
            `in`.close()
            out.close()
        }
    }
}