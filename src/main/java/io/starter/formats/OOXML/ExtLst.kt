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

import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.HashMap
import java.util.Stack

// TODO: FINISH
class ExtLst : OOXMLElement {
    private var attrs: HashMap<String, String>? = null
    private var nameSpace: String? = null

    override//TODO: FINISH
            /*
		StringBuffer ooxml= new StringBuffer();
    	ooxml.append("<" + this.ns + ":ext");
    	// attributes
    	Iterator i= attrs.keySet().iterator();
    	while (i.hasNext()) {
    		String key= (String) i.next();
    		String val= (String) attrs.get(key);
    		ooxml.append(" " + key + "=\"" + val + "\"");
    	}
    	ooxml.append("/>");
    	return ooxml.toString();
    	*/ val ooxml: String
        get() = ""

    constructor() {
        attrs = HashMap()
        attrs!!["cx"] = "0"
        attrs!!["cy"] = "0"
    }

    constructor(attrs: HashMap<String, String>, ns: String) {
        this.attrs = attrs
        this.nameSpace = ns
    }

    constructor(e: ExtLst) {
        this.attrs = e.attrs
        this.nameSpace = e.nameSpace
    }

    /**
     * set the namespace for ext element
     *
     * @param ns
     */
    fun setNamespace(ns: String) {
        this.nameSpace = ns
    }

    override fun cloneElement(): OOXMLElement {
        return ExtLst(this)
    }

    companion object {

        private val serialVersionUID = -4122012942547055359L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            val attrs = HashMap<String, String>()
            var ns: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "extLst") {        // get attributes
                            ns = xpp.prefix
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "extLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("extLst.parseOOXML: $e")
            }

            return Ext(attrs, ns)
        }
    }
}
	
