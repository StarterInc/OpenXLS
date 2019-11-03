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

/**
 * EffectPropsGroup Effect Properties either effectDag or effectLst
 */
//TODO: FINISH CHILD ELEMENTS for both effectDag and effectLst
class EffectPropsGroup : OOXMLElement {

    private var effectDag: EffectDag? = null
    private var effectLst: EffectLst? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            if (effectDag != null) ooxml.append(effectDag!!.ooxml)
            if (effectLst != null) ooxml.append(effectLst!!.ooxml)
            return ooxml.toString()
        }

    constructor(ed: EffectDag, el: EffectLst) {
        this.effectDag = ed
        this.effectLst = el
    }

    constructor(e: EffectPropsGroup) {
        this.effectDag = e.effectDag
        this.effectLst = e.effectLst
    }

    override fun cloneElement(): OOXMLElement {
        return EffectPropsGroup(this)
    }

    companion object {

        private val serialVersionUID = 8250236905326475833L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            var ed: EffectDag? = null
            var el: EffectLst? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "effectDag") {
                            lastTag.push(tnm)
                            ed = EffectDag.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        } else if (tnm == "effectLst") {
                            lastTag.push(tnm)
                            el = EffectLst.parseOOXML(xpp, lastTag)
                            lastTag.pop()
                            break
                        }
                    } else if (eventType == XmlPullParser.END_TAG) { // shouldn't get here
                        lastTag.pop()
                        break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("EffectPropsGroup.parseOOXML: $e")
            }

            return EffectPropsGroup(ed, el)
        }
    }
}

/**
 * effectDag (Effect Container)
 * This element specifies a list of effects. Effects are applied in the order specified by the container type (sibling or
 * tree).
 *
 *
 * parent: many
 * children: MANY (EFFECT)
 */ // TODO: FINISH CHILD ELEMENTS
internal class EffectDag : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    //    	if (CHILD!=null) { ooxml.append(CHILD.getOOXML());
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:effectDag")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append(">")
            ooxml.append("</a:effectDag>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(e: EffectDag) {
        this.attrs = e.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return EffectDag(this)
    }

    companion object {

        private val serialVersionUID = 4786440439664356745L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): EffectDag {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "effectDag") {
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "CHILDELEMENT") {
                            lastTag.push(tnm)
                            //layout = (layout) layout.parseOOXML(xpp, lastTag).clone();

                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "effectDag") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("effectDag.parseOOXML: $e")
            }

            return EffectDag(attrs)
        }
    }
}

/**
 * effectLst (Effect Container)
 * This element specifies a list of effects. Effects in an effectLst are applied in the default order by the rendering
 * engine. The following diagrams illustrate the order in which effects are to be applied, both for shapes and for
 * group shapes.
 *
 *
 * parent: many
 * children: MANY (EFFECT)
 */ // TODO: FINISH CHILD ELEMENTS
internal class EffectLst : OOXMLElement {

    override//StringBuffer ooxml= new StringBuffer();
    //    	ooxml.append("<a:effectLst");
    //    	if (CHILD!=null) { ooxml.append(CHILD.getOOXML());
    //    	ooxml.append("</a:effectLst>");
    // TODO: FINISH CHILD ELEMENTS
    //    	return ooxml.toString();
    val ooxml: String
        get() = "<a:effectLst/>"

    //	public effectLst() { 	}
    constructor() {}

    constructor(e: EffectLst) {}

    override fun cloneElement(): OOXMLElement {
        return EffectLst(this)
    }

    companion object {

        private val serialVersionUID = -6164888373165090983L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): EffectLst {
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "CHILDELEMENT") {
                            //lastTag.push(tnm);
                            //layout = (layout) layout.parseOOXML(xpp, lastTag).clone();
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "effectLst") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("effectLst.parseOOXML: $e")
            }

            return EffectLst()
        }
    }
}

