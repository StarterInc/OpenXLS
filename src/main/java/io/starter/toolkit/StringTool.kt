/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.toolkit

import java.io.*
import java.nio.CharBuffer
import java.util.StringTokenizer

/**
 * A collection of useful methods for manipulating Strings.
 */
class StringTool : Serializable {

    // escaped slashes
    internal var oneslash = 0x005C.toChar().toString()
    internal var twoslash = oneslash + oneslash

    companion object {

        // static final long serialVersionUID =
        // -5757918511951798619l;
        internal const val serialVersionUID = -2761264230959133529L

        /**
         * replace illegal XML characters with their html counterparts
         *
         *
         * ie: "&" is converted to "&amp;"
         *
         *
         * check out [the w3 list of XML
 * characters](http://www.w3.org/TR/REC-xml/)
         *
         * @param rep
         * @return
         */
        fun convertXMLChars(rep: String): String {
            return io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(rep)
                    .toString() // 20110815 KSC: this method is more complete
        }

        // test stuff
        @JvmStatic
        fun main(args: Array<String>) {

            val majorHTML = "<html><body>Testing <b>yes</b><ul><li>item1</li><li>item2</li><li>item3</li></ul><br/>newline<br/>newline2<BR/>newline3 yes no 124<>something <bR>newline4<b>bold</b></body></html>"

            Logger.logInfo(StringTool.stripHTML(majorHTML))

        }

        /**
         * strip out all most HTML tags
         *
         * @param rep
         * @return a string stripped of all html tags
         */
        fun stripHTML(rep: String): String {
            var rep = rep

            // first convert newlines
            rep = rep.replace("<[B,b][R,r]?>".toRegex(), "\r\n")
            rep = rep.replace("<[B,b][R,r]?/>".toRegex(), "\r\n")

            rep = rep.replace("<[L,l][I,i]?>".toRegex(), "\r\n\r\n")

            val ret = StringBuffer()
            val cx = rep.toCharArray()
            var skipping = false
            var t = 0
            while (t < cx.size) {
                val tt = cx[t]
                // begin match
                if (tt == '<') {
                    skipping = true
                    t++
                } else if (tt == '>') {
                    skipping = false
                }
                if (!skipping && tt != '>')
                    ret.append(cx[t])
                t++

            }
            return ret.toString()
        }

        /**
         * replace endoded text with normal text ie: "&amp;" is converted to "&"
         *
         *
         * check out [the w3 list of XML
 * characters](http://www.w3.org/TR/REC-xml/)
         *
         * @param rep
         * @return
         */
        fun convertHTML(rep: String): String {
            // if(true)return rep;
            rep.replace("&amp;".toRegex(), "&")
            rep.replace("&apos;".toRegex(), "'")
            rep.replace("&quot;".toRegex(), "\"")
            rep.replace("&lt;".toRegex(), "<")
            rep.replace("&gt;".toRegex(), ">")
            rep.replace("&copy;".toRegex(), "")
            return rep
        }

        /**
         * If the string matches any part of the pattern, strip the pattern from the
         * string.
         */
        fun stripMatch(pattern: String, matchstr: String): String {
            val upat = pattern.toUpperCase()
            val umat = matchstr.toUpperCase()
            var retval = ""
            if (umat.lastIndexOf(upat) > -1) {
                val pos = umat.lastIndexOf(upat) + upat.length
                retval = matchstr.substring(pos)
                Logger.logInfo("foundpos: $pos")
            }
            return retval
        }

        /**
         * get the variable name for a "getXXXX" a field name per JavaBean java
         * naming conventions.
         *
         *
         * ie: Converts "getFirstName" to "firstName"
         */
        fun getVarNameFromGetMethod(thismethod: String): String {
            val getidx = thismethod.indexOf("get")
            if (getidx < 0)
                return ""
            var retval = thismethod.substring(getidx + 3)
            retval = retval.substring(0, retval.length - 2)
            var upcase = retval.substring(0, 1)
            upcase = upcase.toUpperCase()
            retval = retval.substring(1)
            retval = upcase + retval
            return retval
        }

        /**
         * converts java member naming convention to underscored DB-style naming
         * convention
         *
         *
         * ie: take upperCamelCase and turn into upper_camel_case
         */
        fun convertJavaStyletoDBConvention(name: String): String {
            val chars = name.toCharArray()
            val buf = StringBuffer()
            for (i in chars.indices) {
                // if there is a single upper-case letter, then it's a
                // case-word
                if (Character.isUpperCase(chars[i])) {
                    if (i > 0 && i + 1 < chars.size) {
                        if (!Character.isUpperCase(chars[i]))
                            buf.append("_")
                    }
                    buf.append(chars[i])
                } else {
                    buf.append(Character.toUpperCase(chars[i]))
                }
            }
            return buf.toString()
        }

        /**
         * converts java member naming convention to underscored DB-style naming
         * convention
         *
         *
         * ie: take upperCamelCase and turn into upper_camel_case
         */
        fun convertJavaStyletoFriendlyConvention(name: String): String {
            if (name == "")
                return ""
            val buf = StringBuffer()
            val chars = name.toCharArray()
            buf.append(chars[0].toString().toUpperCase())
            var i = 1
            while (i < chars.size) {
                if (chars[i] == '_') {
                    chars[i++] = Character.toUpperCase(chars[i + 1])
                    buf.append(" ")
                    buf.append(chars[i])
                } else {
                    buf.append(chars[i].toString().toLowerCase())
                }
                i++
            }
            return buf.toString()
        }

        /**
         * convert an Array to a String representation of its objects
         *
         * @param name
         * @return
         */
        fun arrayToString(objs: Array<Any>): String {
            val ret = StringBuffer("[")
            for (x in objs.indices) {
                ret.append(objs[x].toString())
                ret.append(", ")
            }
            ret.setLength(ret.length - 1)
            ret.append("]")
            return ret.toString()
        }

        /**
         * Returns the given throwable's stack trace as a string.
         *
         * @param target the Throwable whose stack trace should be returned
         * @return the stack trace of the given Throwable as a String
         * @throws NullPointerException if `target` is null
         */
        fun stackTraceToString(target: Throwable): String {
            val trace = StringWriter()
            val writer = PrintWriter(trace)

            target.printStackTrace(writer)
            writer.flush()

            return trace.toString()
        }

        /**
         * converts strings to "proper" capitalization
         *
         *
         * ie: take "mr. fraNK sMITH" and turn into "Mr. Frank Smith"
         */
        fun proper(name: String): String {
            if (name == "")
                return ""
            val buf = StringBuffer()
            val chars = name.toCharArray()
            buf.append(chars[0].toString().toUpperCase())
            var i = 1
            while (i < chars.size) {
                if (chars[i] == ' ') {
                    buf.append(" ")
                    i++
                    if (chars[i] != ' ') {
                        chars[i] = Character.toUpperCase(chars[i])
                        buf.append(chars[i])
                    }
                } else {
                    buf.append(chars[i].toString().toLowerCase())
                }
                i++
            }
            return buf.toString()
        }

        /**
         * Each object in the array must support the .toString() method, as it will
         * be used to render the object to its string representation.
         */

        fun makeDelimitedList(list: Array<Any>, delimiter: String): String {
            val listBuf = StringBuffer()
            for (i in list.indices) {
                if (i != 0)
                    listBuf.append(delimiter)
                listBuf.append(list[i].toString())
            }
            return listBuf.toString()
        }

        /**
         * Returns an array of strings from a single string, similar to
         * String.split() in JavaScript
         */
        fun splitString(value: String, delimeter: String): Array<String> {

            val stoken = StringTokenizer(value, delimeter)
            val returnValue = arrayOfNulls<String>(stoken.countTokens())
            var i = 0
            while (stoken.hasMoreTokens()) {
                returnValue[i] = stoken.nextToken()
                i++
            }
            return returnValue
        }

        /**
         * get compressed UNICODE string from uncompressed string
         */
        fun getCompressedUnicode(input: ByteArray): String {
            val output = ByteArray(input.size / 2)
            var pos = 0
            var i = 0
            while (i < input.size) {
                output[pos++] = input[i]
                i++
                i++
            }
            return String(output)
        }

        fun getLowerCaseFirstLetter(thismember: String): String {
            var thismember = thismember
            var upcase = thismember.substring(0, 1)
            upcase = upcase.toLowerCase()
            thismember = thismember.substring(1)
            thismember = upcase + thismember
            return thismember
        }

        /**
         * generate a "setXXXX" string from a field name per Extentech java naming
         * conventions.
         *
         *
         * ie: Converts "firstName" to "setFirstName"
         */
        fun getSetMethodNameFromVar(thismember: String): String {
            return "set" + getUpperCaseFirstLetter(thismember)
        }

        fun getUpperCaseFirstLetter(thismember: String): String {
            var thismember = thismember
            var upcase = thismember.substring(0, 1)
            upcase = upcase.toUpperCase()
            thismember = thismember.substring(1)
            thismember = upcase + thismember
            return thismember
        }

        /**
         * generate a "getXXXX" string from a field name per Extentech java naming
         * conventions.
         *
         *
         * ie: Converts "firstName" to "getFirstName"
         */
        fun getGetMethodNameFromVar(thismember: String): String {
            return "get" + getUpperCaseFirstLetter(thismember)
        }

        /**
         * converts underscored DB-style naming convention to java member naming
         * convention
         */
        fun convertDBtoJavaStyleConvention(name: String): String {
            val scoreloc = name.indexOf("_")
            val buf = StringBuffer()
            if (scoreloc < 0)
                return name
            else {
                val chars = name.toCharArray()
                for (i in chars.indices) {
                    if (chars[i] == '_') {
                        chars[i + 1] = Character.toUpperCase(chars[i + 1])
                    } else {
                        buf.append(chars[i])
                    }
                }
            }
            return buf.toString()
        }

        /**
         * converts underscored DB-style naming convention to java member naming
         * convention
         */
        fun convertFilenameToJSPName(name: String): String {
            val buf = StringBuffer()
            val chars = name.toCharArray()
            for (i in chars.indices) {
                if (chars[i] == '_') {
                    chars[i + 1] = Character.toUpperCase(chars[i + 1])
                } else if (chars[i] == '-') {
                    chars[i + 1] = Character.toUpperCase(chars[i + 1])
                } else if (chars[i] == ' ') {
                    chars[i + 1] = Character.toUpperCase(chars[i + 1])
                } else {
                    buf.append(chars[i])
                }
            }
            return buf.toString()
        }

        /**
         * Basically a String tokenizer
         *
         * @param instr
         * @param token
         * @return
         */
        fun getTokensUsingDelim(instr: String, token: String): Array<String> {
            if (instr.indexOf(token) < 0) {
                val ret = arrayOfNulls<String>(1)
                ret[0] = instr
                return ret
            }
            val output = CompatibleVector()
            StringBuffer()
            var lastpos = 0
            var offset = 0
            val toklen = token.length
            var pos = instr.indexOf(token)
            // pos--;
            while (pos > -1) {
                if (lastpos > 0)
                // if the line starts with a token
                    offset = lastpos + toklen
                else
                    offset = 0
                val st = instr.substring(offset, pos)
                output.add(st)
                lastpos = pos
                pos = instr.indexOf(token, lastpos + 1)
            }
            if (lastpos < instr.length) {
                val st = instr.substring(lastpos + toklen)
                output.add(st)
            }
            val retval = arrayOfNulls<String>(output.size)
            for (i in output.indices)
                retval[i] = output[i] as String
            return retval
        }

        /**
         *
         */
        fun dbencode(holder: String): String {
            return replaceText(holder, "'", "''", 0)
        }

        /**
         * Lose the whitespace at the end of strings...
         *
         * @param holder The String that you want stripped.
         * @return Your stripped string.
         */

        fun allTrim(holder: String): String {
            var holder = holder
            holder = holder.trim { it <= ' ' }
            return rTrim(holder)
        }

        /**
         * strip trailing spaces
         */
        fun stripTrailingSpaces(s: String): String {
            var s = s
            while (s.endsWith(" "))
                s = s.substring(0, s.length - 1)
            return s
        }

        /**
         * Strips all occurences of a string from a given string.
         *
         * @param tostrip   The String that you want stripped.
         * @param stripchar The char you want stripped from the String.
         * @return Your stripped string.
         */

        fun strip(tostrip: String, stripstr: String): String {
            var tostrip = tostrip
            val stripped = StringBuffer(tostrip.length)
            while (tostrip.indexOf(stripstr) > -1) {
                stripped.append(tostrip, 0, tostrip.indexOf(stripstr))
                tostrip = tostrip
                        .substring(tostrip.indexOf(stripstr) + stripstr.length)
            }
            stripped.append(tostrip)
            return stripped.toString()
        }

        /**
         * Strips all occurences of a character from a given string.
         *
         * @param tostrip   The String that you want stripped.
         * @param stripchar The char you want stripped from the String.
         * @return Your stripped string.
         */

        fun strip(tostrip: String, stripchar: Char): String {
            val stripped = StringBuffer(tostrip.length)
            var i = 0
            var currentChar: Char
            while (i < tostrip.length) {
                currentChar = tostrip[i]
                if (currentChar == stripchar) {
                    i++
                } else {
                    stripped.append(currentChar)
                    i++
                }
            }
            return stripped.toString()
        }

        /**
         * Replaces an occurence of String B with String C within String A. This
         * method is case sensitive.
         *
         *
         * Example: String A = "I am a happy dog."; String A =
         * stringtool.replaceText(A, "happy", "sad", 0);
         *
         *
         * The result is A="I am a sad dog."
         *
         * @param originalText    Original text
         * @param replaceText     Text to replace.
         * @param replacementText Text to replace with.
         * @param offset          offset of replacement within original string.
         * @return Processed text.
         */
        fun replaceText(originalText: String, replaceText: String, replacementText: String, offset: Int, skipmatch: Boolean): String {
            if (!skipmatch)
                return replaceText(originalText, replaceText, replacementText, offset)

            val sb = StringBuffer()
            if (originalText.indexOf(replaceText) < 0) {
                return originalText
            } else {
                var nextidx = 0
                var lastidx = 0
                var pos = 0
                val textlen = replaceText.length
                val stringlen = originalText.length

                while (nextidx <= originalText.lastIndexOf(replaceText)) {
                    pos++
                    nextidx = originalText.indexOf(replaceText, lastidx)
                    sb.append(originalText, lastidx, nextidx)
                    if (pos > offset)
                        sb.append(replacementText)
                    else
                        sb.append(replaceText)
                    nextidx += textlen
                    if (textlen == 0)
                        break// case of ""
                    lastidx = nextidx
                }
                if (nextidx < stringlen) {
                    sb.append(originalText.substring(nextidx))
                }
                return sb.toString()
            }
        }

        /**
         * Replaces an occurence of String B with String C within String A. This
         * method is case sensitive.
         *
         *
         * Example: String A = "I am a happy dog."; String A =
         * stringtool.replaceText(A, "happy", "sad", 0);
         *
         *
         * The result is A="I am a sad dog."
         *
         * @param originalText    Original text
         * @param replaceText     Text to replace.
         * @param replacementText Text to replace with.
         * @param offset          offset of replacement within original string.
         * @return Processed text.
         */
        fun replaceText(originalText: String, replaceText: String, replacementText: String, offset: Int): String {

            var newlen = originalText.length - replaceText.length + replacementText.length

            if (newlen < 1)
                newlen = 0
            val sb = StringBuffer(newlen)
            if (originalText.indexOf(replaceText) < 0) {
                return originalText
            } else {
                if (replaceText != null && replaceText == replacementText)
                    return originalText // avoid infinite loops
                var nextidx = 0
                var lastidx = 0
                val textlen = replaceText.length
                val stringlen = originalText.length
                while (nextidx <= originalText.lastIndexOf(replaceText)) {
                    nextidx = originalText.indexOf(replaceText, lastidx)
                    sb.append(originalText, lastidx, nextidx + offset)
                    sb.append(replacementText)
                    nextidx += textlen
                    if (textlen == 0)
                        break // case of ""
                    lastidx = nextidx
                }
                if (nextidx < stringlen) {
                    sb.append(originalText.substring(nextidx))
                }
                return sb.toString()
            }
        }

        /**
         * Trims whitespace from the right side of strings.
         *
         * @param originalText Text to trim.
         * @return Trimmed text.
         */
        fun rTrim(originalText: String): String {
            var sb = StringBuffer(originalText)
            sb.reverse()
            val rstr = sb.toString()
            rstr.trim { it <= ' ' }
            sb = StringBuffer(rstr)
            sb.reverse()
            return sb.toString()
        }

        /**
         * This method will retrieve the first instance of text between any two
         * given patterns. This method is case sensitive.
         *
         *
         * Example:
         *
         *
         * String A = "I am a happy dog."; B = getTextBetweenDelims(A,"a",".");
         *
         *
         * B is now equal to "m a happy dog". C declines comment.
         *
         * @param originalText Text to process.
         * @param beginDelim   Delimeter for beginning of retrieved section.
         * @param endDelim     Delimeter for end of retrieved section.
         * @return Text between delims or "" if not found.
         */
        fun getTextBetweenNestedDelims(originalText: String, beginDelim: String, endDelim: String): String {
            val sb = StringBuffer(originalText.length)
            // Check to see that both delimiters exist in the string
            if (originalText.indexOf(beginDelim) < 0 || originalText.lastIndexOf(endDelim) < 0) {
                return ""
            } else {
                val begidx = originalText.indexOf(beginDelim) + beginDelim.length
                var endidx = originalText.lastIndexOf(endDelim)
                var holder = 0
                if (begidx < endidx) {
                    sb.append(originalText, begidx, endidx)
                } else {
                    while (begidx > endidx && endidx > -1) {
                        holder = endidx
                        endidx = originalText.lastIndexOf(endDelim, holder + 1)
                    }
                    if (begidx < endidx && endidx > -1)
                        sb.append(originalText, begidx, endidx)
                }
                return sb.toString()
            }
        }

        /**
         * This method will retrieve the first instance of text between any two
         * given patterns. This method is case sensitive.
         *
         *
         * Example:
         *
         *
         * String A = "I am a happy dog."; B = getTextBetweenDelims(A,"a",".");
         *
         *
         * B is now equal to "m a happy dog". C declines comment.
         *
         * @param originalText Text to process.
         * @param beginDelim   Delimeter for beginning of retrieved section.
         * @param endDelim     Delimeter for end of retrieved section.
         * @return Text between delims or "" if not found.
         */
        fun getTextBetweenDelims(originalText: String, beginDelim: String, endDelim: String): String {
            val sb = StringBuffer(originalText.length)
            // Check to see that both delimiters exist in the string
            if (originalText.indexOf(beginDelim) < 0 || originalText.indexOf(endDelim) < 0) {
                return ""
            } else {
                val begidx = originalText.indexOf(beginDelim) + beginDelim.length
                var endidx = originalText.indexOf(endDelim)
                var holder = 0
                if (begidx < endidx) {
                    sb.append(originalText, begidx, endidx)
                } else {
                    while (begidx > endidx && endidx > -1) {
                        holder = endidx
                        endidx = originalText.indexOf(endDelim, holder + 1)
                    }
                    if (begidx < endidx && endidx > -1)
                        sb.append(originalText, begidx, endidx)
                }
                return sb.toString()
            }
        }

        /**
         * This method will replace any instance of given text within another
         * string. This method is case sensitive.
         *
         *
         * Example:
         *
         *
         * String A = "I am a happy dog."; A =
         * replaceSection(A,"happy","hippie cat", "dog");
         *
         *
         * A is now equal to "I am a hippie cat.".
         *
         * @param originalText    Text to process.
         * @param replaceBegin    Beggining pattern of replaced section.
         * @param replacementText Text to replace with.
         * @param replaceEnd      End pattern of replaced section.
         * @return Processed text.
         */

        fun replaceSection(originalText: String, replaceBegin: String, replacementText: String, replaceEnd: String): String {
            val sb = StringBuffer(originalText.length)
            if (originalText.indexOf(replaceBegin) < 0 || originalText.indexOf(replaceEnd) < 0) {
                return originalText
            } else {
                val begidx = originalText.indexOf(replaceBegin)
                val endlen = replaceEnd.length
                var endidx = originalText.indexOf(replaceEnd) + endlen
                var holder = 0
                if (begidx < endidx) {
                    sb.append(originalText, 0, begidx)
                    sb.append(replacementText)
                    sb.append(originalText.substring(endidx))
                } else {
                    while (begidx > endidx && endidx > -1) {
                        holder = endidx
                        endidx = originalText.indexOf(replaceEnd, holder + 1)
                    }
                    if (begidx < endidx && endidx > -1) {
                        sb.append(originalText, 0, begidx)
                        sb.append(replacementText)
                        sb.append(originalText.substring(endidx + endlen))
                    }
                }
                return sb.toString()
            }
        }

        fun StripChars(theFilter: String, theString: String): String {
            val strOut = StringBuffer(theString.length)
            var curChar: Char
            for (i in 0 until theString.length) {
                curChar = theString[i]
                if (theFilter.indexOf(curChar.toInt()) < 0) { // if it's not in the filter,
                    // send it thru
                    strOut.append(curChar)
                }
            }
            return strOut.toString()
        }

        fun UseOnlyChars(theFilter: String, theString: String): String {
            val strOut = StringBuffer(theString.length)
            var curChar: Char
            for (i in 0 until theString.length) {
                curChar = theString[i]
                if (theFilter.indexOf(curChar.toInt()) > -1) { // if it's in the filter,
                    // send it thru
                    strOut.append(curChar)
                }
            }
            return strOut.toString()
        }

        fun replaceChars(theFilter: String, theString: String, replacement: String): String {
            val strOut = StringBuffer(theString.length)
            var curChar: Char
            for (i in 0 until theString.length) {
                curChar = theString[i]
                if (theFilter.indexOf(curChar.toInt()) < 0) { // if it's not in the filter,
                    // send it thru
                    strOut.append(curChar)
                } else {
                    strOut.append(replacement)
                }
            }
            return strOut.toString()
        }

        /**
         * replace a section of text based on pattern match throughout string.
         */
        fun replaceText(theString: String, theFilter: String, replacement: String): String {
            return replaceText(theString, theFilter, replacement, 0)
        }

        fun AllInRange(x: Int, y: Int, theString: String): Boolean {
            var curChar: Char
            for (i in 0 until theString.length) {
                curChar = theString[i]
                if (curChar.toInt() < x || curChar.toInt() > y) {
                    return false
                }
            }
            return true
        }

        /**
         * Replaces a specified token in a string with a value from the passed
         * through array this is done in matching order from String to Arrray.
         */
        fun replaceTokenFromArray(replace: String, token: String, vals: Array<String>): String {
            val sb = StringBuffer()
            val toke = StringTokenizer(replace, token, false)
            var i = 0
            // add the first element
            while (toke.hasMoreTokens()) {
                sb.append(toke.nextToken())
                if (i <= vals.size - 1) {
                    sb.append(vals[i])
                    ++i
                }
            }
            return sb.toString()
        }

        /**
         * returns file path fPath qualified with an ending slash
         *
         * @param fPath
         * @return fPath with an ending slash
         */
        fun qualifyFilePath(fPath: String): String {
            var fPath = fPath
            StringTool.replaceChars("\\", fPath, "/")
            fPath = fPath.trim { it <= ' ' }
            if (!fPath.endsWith("/"))
                fPath += "/"
            return fPath
        }

        /**
         * splits a filepath into directory and filename
         *
         * @param filePath
         * @return
         */
        fun splitFilepath(filePath: String): Array<String> {
            var filePath = filePath
            val path = arrayOfNulls<String>(2)
            filePath = StringTool.replaceText(filePath, "\\", "/")
            val lastpath = filePath.lastIndexOf("/")
            if (lastpath > -1) { // strip path and directory
                path[0] = filePath.substring(0, lastpath + 1) // get directory
                path[1] = filePath.substring(lastpath + 1) // strip directory from
                // filename
            } else
                path[1] = filePath
            return path
        }

        /**
         * strips the path portion from a filepath and returns the filename
         *
         * @param filePath
         * @return
         */
        fun stripPath(filePath: String): String {
            var filePath = filePath
            filePath = StringTool.replaceText(filePath, "\\", "/")
            val lastpath = filePath.lastIndexOf("/")
            return if (lastpath > -1) filePath.substring(lastpath + 1) else filePath // strip directory from
            // filename

        }

        /**
         * strips the path portion from a filepath and returns it
         *
         * @param filePath
         * @return
         */
        fun getPath(filePath: String): String {
            var filePath = filePath
            filePath = StringTool.replaceText(filePath, "\\", "/")
            val lastpath = filePath.lastIndexOf("/")
            return filePath.substring(0, lastpath + 1)
        }

        /**
         * replaces the extension of a filepath with an new extension
         *
         * @param filepath Source FilePaht
         * @param ext
         * @return
         */
        fun replaceExtension(filepath: String, ext: String): String {
            val i = filepath.lastIndexOf(".")
            var f = filepath
            if (i > 0)
                f = filepath.substring(0, i) + ext
            else
                f = filepath + ext
            return f
        }

        /**
         * given a string, return the maximum of the width in pixels in the given
         * the awt Font. <br></br>
         * NOTE: this method does not account for line feeds contained within
         * strings
         *
         * @param f awt Font
         * @param s String to compute
         * @return double approximate width in pixels
         */
        fun getApproximateStringWidth(f: java.awt.Font, s: String): Double {
            val fm = java.awt.Toolkit.getDefaultToolkit()
                    .getFontMetrics(f)
            /*
         * width/height in pixels = (w/h field) * DPI of the display
         * device / 72
         */
            val conversion = java.awt.Toolkit.getDefaultToolkit()
                    .screenResolution / 72.0
            return fm.stringWidth(s) * conversion // pixels * conversion
        }

        /**
         * given a string, return the maximum of the width in pixels in the given
         * the awt Font Observing Line Breaks. <br></br>
         *
         * @param f awt Font
         * @param s String to compute
         * @return double approximate width in pixels
         */
        fun getApproximateStringWidthLB(f: java.awt.Font, s: String): Double {
            val fm = java.awt.Toolkit.getDefaultToolkit()
                    .getFontMetrics(f)
            /*
         * width/height in pixels = (w/h field) * DPI of the display
         * device / 72
         */
            val conversion = java.awt.Toolkit.getDefaultToolkit()
                    .screenResolution / 72.0
            val ss = s.split("\n".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            var len = 0.0
            for (st in ss) {
                len = Math.max(len, fm.stringWidth(st) * conversion)
            }
            // return fm.stringWidth(s) * conversion;
            return len
        }

        /**
         * return the approximate witdth in width in pixels of the given character
         *
         * @param f awt Font
         * @param c character
         * @return double approximate width in pixels
         */
        fun getApproximateCharWidth(f: java.awt.Font, c: Char?): Double {
            val fm = java.awt.Toolkit.getDefaultToolkit()
                    .getFontMetrics(f)
            /*
         * width/height in pixels = (w/h field) * DPI of the display
         * device / 72
         */
            val conversion = java.awt.Toolkit.getDefaultToolkit()
                    .screenResolution / 72.0
            return fm.charWidth(c!!) * conversion
        }

        /**
         * return the approximate height it takes to display the given string in the
         * given font in the given width
         *
         * @param f java.awt.Font
         * @param s String
         * @param w width in points
         * @return
         */
        fun getApproximateHeight(f: java.awt.Font, s: String, w: Double): Double {
            var s = s
            var len = StringTool.getApproximateStringWidth(f, s)
            while (len > w) {
                var lastSpace = -1
                var j = s.lastIndexOf("\n") + 1
                len = 0.0
                while (len < w && j < s.length) {
                    len += StringTool.getApproximateCharWidth(f, s[j])
                    if (s[j] == ' ')
                        lastSpace = j
                    j++
                }
                if (len < w) {
                    break // got it
                }
                if (lastSpace == -1) { // no spaces to break apart
                    if (s.indexOf(' ') == -1)
                        break
                    lastSpace = s.lastIndexOf(' ') // break at
                }
                s = s.substring(0, lastSpace) + "\n" + s.substring(lastSpace + 1)
            }
            val nl = s.split("\n".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray().size
            val fm = java.awt.Toolkit.getDefaultToolkit()
                    .getFontMetrics(f)
            val lm = f
                    .getLineMetrics(s, fm.fontRenderContext)
            // this calc appears to match Excel's ...
            val l = lm.leading
            // float h= lm.getHeight();
            // io.starter.toolkit.Logger.log("Font: " +
            // f.toString());
            // io.starter.toolkit.Logger.log("l-i:" +
            // fm.getLeading() + " l:" + l + " h-i:" + fm.getHeight() +
            // " h:" + h + " a-i:" + fm.getAscent() + " a:" +
            // lm.getAscent() + " d-i:" + fm.getDescent() + " d:" +
            // lm.getDescent());
            val h = fm.height.toFloat() // KSC: revert for now ... - l/3; // i don't
            // know why but this seems to match Excel's
            // the closest
            return Math.ceil((h * nl).toDouble())// +1)); // KSC: added + 1 for testing
        }

        /**
         * converts an excel-style custom format to String.format custom format i.e.
         * %flags-width-precision-conversion
         *
         * NOTE: the pattern should be a single item in the excel-style format i.e.one of the terms in positive;negative;zero;text
         * without semicolon.
         * <br></br>NOTE: the java-specific pattern returned will have the negative formatting (sign, parenthesis ...) and so, when used,
         * the double value must be it's absolute value i.e String.format(pattern, Math.abs(d))
         * <br></br>NOTE: date formats are handled separately.  this only applies to number and currency formats
         * <br></br>Excel-style:
         * 0 (zero) 	Digit placeholder. This code pads the value with zeros to fill the format.
         * # 	Digit placeholder. This code does not display extra zeros.
         * ? 	Digit placeholder. This code leaves a space for insignificant zeros but does not display them.
         * . (period) 	Decimal number.
         * % 	Percentage. Microsoft Excel multiplies by 100 and adds the % character.
         * , (comma) 	Thousands separator. A comma followed by a placeholder scales the number by a thousand.
         * E+ E- e+ e- 	Scientific notation.
         * Text Code 	Description
         * $ - + / ( ) : space 	These characters are displayed in the number. To display any other character, enclose the character in quotation marks or precede it with a backslash.
         * \character 	This code displays the character you specify.
         * "text" 	This code displays text.
         * This code repeats the next character in the format to fill the column width.		Note Only one asterisk per section of a format is allowed.
         * _ (underscore) 	This code skips the width of the next character.
         * This code is commonly used as "_)" (without the quotation marks)
         * to leave space for a closing parenthesis in a positive number format
         * when the negative number format includes parentheses.
         * This allows the values to line up at the decimal point.
         *
         * @param pattern    String format pattern in Excel format
         * @param isNegative true if the source is a negative number
         * @return
         * @ Text placeholder.
         */
        fun convertPatternFromExcelToStringFormatter(pattern: String, isNegative: Boolean): String {
            var pattern = pattern
            val curPattern = pattern
            var jpattern = "" // return pattern
            var w = 0
            var precision = 0
            var flags = ""
            var conversion = 'f' // default
            var inConversion = false
            var inPrecision = false
            var removeSign = false // true if value is negative and pattern
            // calls for parens or color change or ...
            // i.e. don't display the negative sign
            /*
         * TODO: \ uXXX is Locale-specific to display? works
         * manually ...
         * TODO; finish fractional formats: ?/?
         */
            var i = 0
            while (i < curPattern.length) {
                var c = curPattern[i].toInt()
                when (c) {
                    '0' -> {
                        w++
                        if (!inConversion) {
                            jpattern += "%"
                            inConversion = true
                        }
                        if (inPrecision && conversion != 'E')
                            precision++
                    }
                    '?' // don't really know what to do with this one!
                    -> {
                    }
                    '#' -> if (!inConversion) {
                        jpattern += "%"
                        inConversion = true
                    }
                    ',' -> flags += ","
                    '.' -> inPrecision = true
                    'E', 'e' -> {
                        if (!inConversion) {
                            jpattern += "%"
                            inConversion = true
                        }
                        conversion = 'E'
                        i++ // format is e+, E+, e- or E-
                    }
                    '[' // either color code or local-specific formatting
                    -> {
                        var j = ++i
                        var k = j
                        while (i < curPattern.length) { // skip colors for now
                            c = curPattern[i].toInt()
                            if (c == '-'.toInt())
                            // got end of an extended char sequence - skip
                            // rest (Locale code ...)
                                k = i
                            if (c == ']'.toInt())
                                break
                            i++
                        }
                        if (inConversion) {
                            inConversion = false
                            inPrecision = false
                            jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                    + conversion)
                        }
                        if (k == j)
                        // then it was a color string
                            removeSign = true
                        else
                        // it was a locale-specific string ...
                            jpattern += curPattern.substring(++j, k)
                    }
                    '"' // start of delimited text
                    -> {
                        if (inConversion) {
                            inConversion = false
                            inPrecision = false
                            jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                    + conversion)
                        }
                        i++
                        while (i < curPattern.length) {
                            c = curPattern[i].toInt()
                            if (c == '"'.toInt())
                                break
                            jpattern += c.toChar()
                            i++

                        }
                    }
                    // ignore
                    '@' // text placeholder
                    -> jpattern += "%s"
                    '*' // repeats the next char to fill -- IGNORE!!!
                    -> {
                    }
                    '(' // enclose negative #'s in parens
                        , ')' ->
                        // flags+="(";
                        if (isNegative) {
                            if (inConversion) {
                                inConversion = false
                                inPrecision = false
                                jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                        + conversion)
                            }
                            jpattern += c.toChar()
                            removeSign = true
                        }
                    '_' // skips the width of the next char - usually _) - to
                    -> {
                        // leave space for a closing parenthesis in a positive
                        // number format when the negative number format
                        // includes parentheses. This allows the values to line
                        // up at the decimal point.
                        if (inConversion) {
                            inConversion = false
                            inPrecision = false
                            jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                    + conversion)
                        }
                        i++ // skip next char -- true in all cases???
                    }
                    '%' -> {
                        if (inConversion) {
                            inConversion = false
                            inPrecision = false
                            jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                    + conversion)
                        }
                        jpattern += "%%"
                    }
                    '\\' -> {
                        if (inConversion) {
                            inConversion = false
                            inPrecision = false
                            jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                    + conversion)
                        }
                        val z: Int
                        if (i + 1 < curPattern.length && curPattern[i + 1] == 'u')
                            z = i + 6
                        else
                            z = i + 1
                        while (i < z && i < curPattern.length) {
                            jpattern += curPattern[i]
                            i++
                        }
                    }
                    else // %, $, - space -- keep
                    -> {
                        if (inConversion) {
                            inConversion = false
                            inPrecision = false
                            jpattern += (flags + (if (w > 0) w else "") + "." + precision
                                    + conversion)
                        }
                        jpattern += c.toChar()
                    }
                }// TODO: handle such as: ###0.00######### --- what's the
                // format spec for that?????
                // if (inPrecision) precision++;
                i++
            }
            if (inConversion) {
                jpattern += flags + (if (w > 0) w else "") + "." + precision + conversion
            }
            if (isNegative && !removeSign)
                jpattern = "-$jpattern"
            // System.out.print("Original Pattern " + pattern + " new "
            // + jpattern);
            // patterns[z]= jpattern;
            pattern = jpattern
            return pattern
        }

        fun convertDatePatternFromExcelToStringFormatter(pattern: String): String {
            var jpattern = "" // return pattern
            var dString = "" // d string -- ddd ==> EEE and dddd ==> EEEE
            var mString = "" // m string -- either month (M, MM, MMM, MMMM) or
            // minute
            var prev = 0
            var i = 0
            while (i < pattern.length) {
                val c = pattern[i].toInt()
                if (c != 'd'.toInt() && dString != "") {
                    if (dString.length <= 2)
                        jpattern += dString
                    else if (dString.length == 3)
                        jpattern += "EEE"
                    else if (dString.length == 4)
                        jpattern += "EEEE"
                    dString = ""
                } else if (c != 'm'.toInt() && mString != "") {
                    if (c == ':'.toInt() || prev == 'h'.toInt()) { // it's time
                        jpattern += mString
                        prev = c
                    } else
                        jpattern += mString.toUpperCase()
                    mString = ""
                }

                when (c) {
                    'y' -> jpattern += c.toChar()
                    'h' -> {
                        jpattern += 'H'.toString() // h in java is 1-24 excel h= 0-23
                        prev = 'h'.toInt()
                    }
                    '\\' // found case of erroneous use of backslash, as in:
                        ,
                        // mm\-dd\-yy ignore!
                    '[' // no java equivalent of [h] [m] or [ss] == elapsed time
                        , ']' -> {
                    }
                    's' -> jpattern += c.toChar()
                    'A' -> if (pattern.substring(i, i + 5) == "AM/PM") {
                        jpattern += "a"
                        i += 5
                        for (z in jpattern.length - 2 downTo 0) {
                            if (jpattern[z] == 'H') {
                                jpattern = (jpattern.substring(0, z) + 'h'.toString()
                                        + jpattern.substring(z + 1))
                            }
                        }
                    }
                    'd' -> dString += c.toChar()
                    'm' -> mString += c.toChar()
                    else -> {
                        if (c != ':'.toInt() && c != 'm'.toInt())
                            prev = c
                        jpattern += c.toChar()
                    }
                }
                i++
            }
            if (mString != "") {
                if (prev == 'h'.toInt())
                // it's time
                    jpattern += mString
                else
                    jpattern += mString.toUpperCase() // remaining month string
            } else if (dString != "") {
                if (dString.length <= 2)
                    jpattern += dString
                else if (dString.length == 3)
                    jpattern += "EEE"
                else if (dString.length == 4)
                    jpattern += "EEEE"
                dString = ""
            }
            return jpattern
        }

        /**
         * extract info, if any, from bracketed expressions within Excel custom number formats
         *
         * @param pattern String Excel number format
         * @return String returned number format without the bracketed expression
         */
        fun convertPatternExtractBracketedExpression(pattern: String): String {
            var pattern = pattern
            val s = pattern.split("\\[".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            if (s.size > 1) {
                pattern = ""
                for (i in s.indices) {
                    val zz = s[i].indexOf("]")
                    if (zz != -1) {
                        var term = ""
                        if (s[i].get(0) == '$')
                            term = s[i].substring(1, zz) // skip first $
                        else
                            term = s[i].substring(0, zz)
                        if (term.indexOf("-") != -1)
                        // extract character TODO:
                        // locale specifics
                            pattern += term.substring(0, term.indexOf("-"))
                        else
                            pattern += term
                    }
                    pattern += s[i].substring(zz + 1)
                }
            }
            return pattern
        }

        /**
         * qualifies a pattern string to make valid for applying the pattern
         *
         * @param pattern
         * @return
         */
        fun qualifyPatternString(pattern: String): String {
            var pattern = pattern
            pattern = StringTool.strip(pattern, "*")
            pattern = StringTool.strip(pattern, "_(") // width placeholder
            pattern = StringTool.strip(pattern, "_)") // width placeholder
            pattern = StringTool.strip(pattern, "_")
            pattern = pattern.replace("\"".toRegex(), "")
            pattern = StringTool.strip(pattern, "?")
            // there are more bracketed expressions to deal with
            // see
            // http://office.microsoft.com/en-us/excel-help/creating-international-number-formats-HA001034635.aspx?redir=0
            // pattern = StringTool.strip(pattern, "[Red]"); // [Black]
            // [h] [hhh] [=1] [=2]
            // pattern = StringTool.strip(pattern, "Red]");
            // TODO: implement locale-specific entries: [$-409] [$-404]
            // ... ********************
            // pattern= pattern.replaceAll("\\[.+?\\]", "");
            /*
         * if (s.length > 1) {
         * io.starter.toolkit.Logger.log(s[0]);
         * java.util.regex.Pattern p =
         * java.util.regex.Pattern.compile("\\[(.*?)\\]");
         * java.util.regex.Matcher m = p.matcher(pattern);
         *
         * while(m.find()) {
         * io.starter.toolkit.Logger.log(m.group(1));
         * }
         * }
         */
            val s = pattern.split("\\[".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            if (s.size > 1) {
                pattern = ""
                for (i in s.indices) {
                    val zz = s[i].indexOf("]")
                    if (zz != -1) {
                        val term = s[i].substring(1, zz) // skip first $
                        if (term.indexOf("-") != -1) { // extract character TODO:
                            // locale specifics
                            pattern += term.substring(0, term.indexOf("-"))
                        }
                    }
                    pattern += s[i].substring(zz + 1)
                }
            }
            return pattern
        }

        /**
         * Reads from a `Reader` into a `String`.
         * Blocking reads will be issued to the reader and the results will be
         * concatenated into a string, which will be returned once the reader
         * reports end-of-input.
         *
         * @param reader the `Reader` from which to read
         * @return a string containing all characters read from the input
         */
        @Throws(IOException::class)
        fun readString(reader: Reader): String {
            val builder = StringBuilder()
            val buffer = CharBuffer.allocate(512)

            while (-1 != reader.read(buffer)) {
                buffer.flip()
                builder.append(buffer)
                buffer.clear()
            }

            return builder.toString()
        }
    }
}