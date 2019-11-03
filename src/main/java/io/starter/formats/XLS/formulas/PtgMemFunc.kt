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
package io.starter.formats.XLS.formulas

import io.starter.formats.XLS.Boundsheet
import io.starter.formats.XLS.ExpressionParser
import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.formats.XLS.WorkBook
import io.starter.toolkit.ByteTools
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.Stack

/**
 * PtgMemFunc refers to a reference subexpression that doesn't evaluate
 * to a constant reference.  This is still somewhat unclear to me how it functions,
 * or why it exists for that matter.
 *
 *
 * This token encapsulates a reference subexpression that results in a non-constant cell address,
 * cell range address, or cell range list. For
 *
 *
 * A little update on it.  Apparently this is used in situations where a record only contains one ptg, but needs to refer to a whole stack of them.
 * An example is in the Name record.  For a built in name that has both row & col repeat regions, the name expression is a ptgmemfunc, but contains
 * 2 ptgArea3d's.
 * Also used in Ai's (Chart Series Range Refs) when non-contiguous range refs are required
 *
 *
 * PtgMemFunc basically represents a complex range and is used where only one PtgRef-type ptg is expected
 *
 *
 * NOTE: that this represents a NON-CONSTANT expression while PtgMemArea represents a CONSTANT expression
 *
 * <pre>
 * OFFSET		NAME		SIZE		CONTENTS
 * -------------------------------------------------------
 * 0			cce			2			The length of the reference subexpression
</pre> */
class PtgMemFunc : GenericPtg() {

    /**
     * @return Returns the lastPtg.
     */
    var subExpression: Stack<*>? = null
        internal set //

    internal var ptgs: Array<Ptg>? = null    // 20090905 KSC: can be PtgRef3d, PtgArea3d, PtgName  ...

    override val isOperand: Boolean
        get() = true

    override val isReference: Boolean
        get() = true    // 20100202 KSC

    internal var cce: Int = 0
    internal var subexp: ByteArray

    internal var refsheets: ArrayList<String>? = ArrayList()

    /**
     * update the record internally for ptgmemfunc
     */
    override val record: ByteArray
        get() {
            var len = 0
            for (i in subExpression!!.indices) {
                val p = subExpression!![i] as Ptg
                len += p.record.size
            }
            val b = ByteArray(len + 3)
            val leng = ByteTools.shortToLEBytes(len.toShort())
            b[0] = 0x29
            b[1] = leng[0]
            b[2] = leng[1]
            var offset = 3
            for (i in subExpression!!.indices) {
                val p = subExpression!![i] as Ptg
                System.arraycopy(p.record, 0, b, offset, p.record.size)
                offset += p.record.size
            }
            record = b
            return record
        }

    override val length: Int
        get() = cce + 3

    internal var calc_id = 1

    override// not parsed yet
    val value: Any?
        get() {
            if (ptgs == null)
                parseSubexpression()
            try {
                val dub: DoubleArray?
                try {
                    dub = PtgCalculator.getDoubleValueArray(ptgs!!)
                } catch (e: CalculationException) {
                    return null
                }

                var result = 0.0
                for (i in dub!!.indices) {
                    result += dub[i]
                }
                return result
            } catch (e: FunctionNotSupportedException) {
                Logger.logWarn("Function Unsupported error in PtgMemFunction: $e")
                return null
            }

        }

    internal var colrefs: Array<PtgRef>? = null

    internal var comps: Array<PtgRef>? = null

    /**
     * Returns all of the cells of this range as PtgRef's.
     * This includes empty cells, values, formulas, etc.
     * Note the setting of parent-rec requires finding the cell
     * the PtgRef refer's to.  If that is null then the PtgRef
     * will exist, just with a null value.  This could cause issues when
     * programatically populating cells.
     */
    override// not parsed yet
    val components: Array<Ptg>?
        get() {
            if (ptgs == null)
                parseSubexpression()
            return ptgs
        }

    /**
     * @return Returns the firstPtg.
     */
    // not parsed yet
    val firstloc: Ptg?
        get() {
            if (ptgs == null)
                parseSubexpression()
            return if (ptgs != null) ptgs!![0] else null
        }

    /**
     * Ptgs upkeep their mapping in reference tracker, however, some ptgs
     * are components of other Ptgs, such as individual ptg cells in a PtgArea.  These
     * should not be stored in the RT.
     */
    var useReferenceTracker = true

    /**
     * given a complex range, parse and set this PtgMemFunc's associated ptgs
     *
     * @param String complexrange  String representing a complex range
     */
    // some possibilities
    // a:b,c:d
    // a:b c:d,e:f g:h
    // a, b, c, d
    // Q34:Q36:Q35:Q37 Q36:Q38
    // name1:name2:name3:name4
    /*
     * results of parsing complex ranges from Excel:
     *  @V@27, ), @V$27, ), " "		=(($V$27) ($V$27))
		Q27:Q29, Q28:Q30, ","		'=SUM((Q27:Q29,Q28:Q30))
		Q27:Q29, Q28:Q30, " "		'=SUM(Q27:Q29 Q28:Q30)
		Q27:Q29, Q28:Q30, ":"		'=SUM(Q27:Q29:Q28:Q30)
		Q34:Q36, Q35:Q37, ":", Q36:Q38, " ", Q37:Q39, ","	'=SUM((Q34:Q36:Q35:Q37 Q36:Q38,Q37:Q39))
		Q34:Q36, Q35:Q37, ":", ), Q36:Q38, " ", Q37:Q39, ","	'=SUM(((Q34:Q36:Q35:Q37) Q36:Q38,Q37:Q39))
		Q34:Q36, Q35:Q37, Q36:Q38, " ", ), ":", Q37:Q39, ","	'=SUM((Q34:Q36:(Q35:Q37 Q36:Q38),Q37:Q39))
		Q34:Q36, Q35:Q37, ":", Q36:Q38, Q37:Q39, ",", ), " "	'=SUM((Q34:Q36:Q35:Q37 (Q36:Q38,Q37:Q39)))
		Q34:Q36, Q35:Q37, ":", Q36:Q38, " ", ), Q37:Q39, ","	'=SUM(((Q34:Q36:Q35:Q37 Q36:Q38),Q37:Q39))
		Q34:Q36, Q35:Q37, Q36:Q38, " ", Q37:Q39, ",", ), ":"	'=SUM((Q34:Q36:(Q35:Q37 Q36:Q38,Q37:Q39)))
		Q40:Q42, Q41:Q43, Q42:Q44, Q43:Q45, ":", " ", ","		'=SUM((Q40:Q42,Q41:Q43 Q42:Q44:Q43:Q45))
		Q40:Q42, Q41:Q43, ",", ), Q42:Q44, Q43:Q45, ":", " "	'=SUM(((Q40:Q42,Q41:Q43) Q42:Q44:Q43:Q45))
		Q40:Q42, Q41:Q43, " ", ), Q43:Q44, ":", Q45, ":", ","	'=SUM((Q40:Q42,(Q41:Q43 Q42):Q43:Q44:Q45))
		Q40:Q42, Q41:Q43, Q42:Q44, Q43:Q45, ":", ), " ", ","	'=SUM((Q40:Q42,Q41:Q43 (Q42:Q44:Q43:Q45)))
		Q40:Q42, Q41:Q43, Q42:Q44, " ", ",", ), Q43:Q45, ":"	'=SUM(((Q40:Q42,Q41:Q43 Q42:Q44):Q43:Q45))
		Q40:Q42, Q41:Q43, Q42:Q44, Q43:Q45, ":", " ", ), ","	'=SUM((Q40:Q42,(Q41:Q43 Q42:Q44:Q43:Q45)))
		Q46, Q46, " "											'=-Q46 Q46
		Q47, ), Q47, ":"										'=-(Q47):Q47
		Q48, Q48, ), ":"										'=-Q48:(Q48)
     */
    override// 1st 3 bytes= id + cce (length of following data)
//?
// memfuncs are assumed to be wrapped in parens, no need to specify
// KSC: TESTING: revert settng subsexpression here for now as tests fail ((;
//this.subexpression= new Stack();
// structure:
// ref, ref, op, [op?] [, ref, op ...]
// it's an operator
//this.subexpression.add(pu);
//this.subexpression.add(pi);
//this.subexpression.add(pr);
//this.subexpression.add(pp);
// in the rare case of PtgMemFuncs which contain embedded formulas, Ptgs are already created (see parseFmla)
//this.subexpression.add(o);
// TODO: handle in quote!!!
//important for charting/ptgmemfuncs in series/categories - will error on open otherwise
//this.subexpression.add(p);
// ptgId
// KSC: don't re-parse as already have all the ptgs ... also, rw/col bytes for Excel2007 exceed maximums so conversion can't be 100%
// KSC: TESTING: revert for now tests fail ((;
/*			cce = ByteTools.readShort(record[1], record[2]);
			subexp = new byte[cce];
	        System.arraycopy(record, 3, subexp, 0, cce);/**/ var location:String
get
 set(complexrange) {
var complexrange = complexrange
var newData = ByteArray(3)
var sheetName = ""
var bk:WorkBook? = null
val sheets = ArrayList<String>()
try
{
bk = this.parentRec!!.workBook
for (i in 0 until bk!!.sheetVect.size)
{
sheets.add(bk!!.getWorkSheetByNumber(i).sheetName)
}
sheetName = this.parentRec!!.sheet!!.sheetName + "!"
}
catch (e:Exception) {}


if (complexrange.startsWith("(") && complexrange.endsWith(")"))
complexrange = complexrange.substring(0, complexrange.length - 1)
val refs = parseFmla(complexrange)
try
{
val ref:String
while (refs.size != 0)
{
while (refs.size > 0)
{
if (refs.get(0) is Char)
{
val cOp = refs.get(0) as Char
if (cOp!!.charValue() == ',')
{
val pu = PtgUnion()
cce += pu.record.size
newData = ByteTools.append(pu.record, newData)
}
else if (cOp!!.charValue() == ' ')
{
val pi = PtgIsect()
cce += pi.record.size
newData = ByteTools.append(pi.record, newData)
}
else if (cOp!!.charValue() == ':')
{
val pr = PtgRange()
cce += pr.record.size
newData = ByteTools.append(pr.record, newData)
}
else if (cOp!!.charValue() == ')')
{
val pp = PtgParen()
cce += pp.record.size
newData = ByteTools.append(pp.record, newData)
}
}
else
{
val o = refs.get(0)
if (o is Ptg)
{
cce += (o as Ptg).record.size
newData = ByteTools.append((o as Ptg).record, newData)
}
else
{
ref = o as String
val isName = (this.parentRec!!.workBook!!.getName(ref) != null)
var p:Ptg? = null
if (isName)
{
p = PtgName()
p!!.parentRec = this.parentRec
(p as PtgName).setName(ref)
cce += p!!.record.size
newData = ByteTools.append(p!!.record, newData)
}
else if (ref.indexOf(":") > 0)
{
if (ref.indexOf("!") == -1) ref = sheetName + ref
p = PtgArea3d()
p!!.parentRec = this.parentRec
p!!.location = ref
cce += p!!.record.size
newData = ByteTools.append(p!!.record, newData)
}
else
{
if (ref.indexOf("!") == -1) ref = sheetName + ref
p = PtgRef3d()
p!!.parentRec = this.parentRec
p!!.location = ref
(p as PtgRef3d).setPtgType(PtgRef.REFERENCE)
cce += p!!.record.size
newData = ByteTools.append(p!!.record, newData)
}
}
}
refs.removeAt(0)
}
}
}
catch (e:Exception) {
throw IllegalArgumentException("PtgMemFunc Error Parsing Location " + complexrange + ":" + e.toString())
}

val ix = ByteTools.shortToLEBytes(cce.toShort())
System.arraycopy(ix, 0, newData, 1, 2)
newData[0] = 41
record = newData
try
{
populateVals()
}
catch (e:Exception) {
Logger.logErr("PtgMemFunc setLocation failed for: " + complexrange + " " + e.toString())
}

}

public override fun init(b:ByteArray) {
opcode = b[0]
record = b
try
{
this.populateVals()
}
catch (e:Exception) {
Logger.logErr("PtgMemFunc init: " + e.toString())
}

}

@Throws(Exception::class)
internal fun populateVals() {
cce = ByteTools.readShort(record[1].toInt(), record[2].toInt()).toInt()
subexp = ByteArray(cce)
System.arraycopy(record, 3, subexp, 0, cce)
 // subexpression stack in form of:  REFERENCE, REFERENCE, OP [,REFERENCE, OP] ...
        // op can be one of:  PtgUnion [,] PtgIsect [ ] or PtgRange [:]
        subExpression = ExpressionParser.parseExpression(subexp, this.parentRec)
 // try parsing/calculating on-demand rather than upon init
        //parseSubexpression();
    }

/**
 * return the ptg components for a certain column within a ptgArea()
 *
 * @param colNum
 * @return all Ptg's within colNum
 */
     fun getColComponents(colNum:Int):Array<Ptg> {
if (colrefs != null)
 // cache
            return colrefs
val v = FastAddVector()
val allComponents = this.components
for (i in allComponents!!.indices)
{
val p = allComponents!![i] as PtgRef
 //			 TODO: check rc sanity here
            val x = p.intLocation
if (x!![1] == colNum) v.add(p)
}
colrefs = arrayOfNulls<PtgRef>(v.size)
v.toTypedArray()
return colrefs
}


/**
 * parses subexpression into ptgs array + traps referenced sheets
 */
    private fun parseSubexpression() {
 // calculate subexpression to obtain ptgs
        val o = FormulaCalculator.calculateFormula(this.subExpression!!)
val components = ArrayList<Ptg>()
if (o != null && o is Array<Ptg>)
{
 // Firstly: take subexpression and remove reference-tracked elements; calcualted elements are ref-tracked below
            for (i in subExpression!!.indices)
{
try
{
(subExpression!!.get(i) as PtgRef).removeFromRefTracker()
}
catch (e:Exception) {}

}
ptgs = o as Array<Ptg>?
for (i in ptgs!!.indices)
{
try
{
if (!refsheets!!.contains((ptgs!![i] as PtgRef).sheetName))
refsheets!!.add((ptgs!![i] as PtgRef).sheetName)
(ptgs!![i] as PtgRef).addToRefTracker()
if ((ptgs!![i] is PtgArea) and !(ptgs!![i] is PtgAreaErr3d))
{
val p = ptgs!![i].components
for (j in p.indices)
components.add(p[j])
}
else
components.add(ptgs!![i])
}
catch (e:Exception) {
Logger.logErr("PtgMemFunc init: " + e.toString())
}

}
}
else
{    // often a single reference surrounded by parens
for (i in subExpression!!.indices)
{
try
{
val pr = subExpression!!.get(i) as PtgRef
if (!refsheets!!.contains(pr.sheetName))
refsheets!!.add(pr.sheetName)
if (pr is PtgArea)
{
val pa = pr.components
for (j in pa!!.indices)
components.add(pa!![j])
}
else
components.add(pr)
}
catch (e:Exception) {}

}
}
ptgs = arrayOfNulls<Ptg>(components.size)
components.toTypedArray<Ptg>()
}

/**
 * takes a range string which may contain operators:
 * union [,] isect [ ] range [:} or paren
 * plus range elements and/or named range
 * parse and order each element correctly
 * may be called recurrsively
 * NOTE: may also be VERY complex, of type OFFSET(x,y,0):OFFSET(z,w,0)
 * ALSO INDEX and INDIRECT ...
 *
 * @param complexrange
 * @return ordered stack containing parsed range elements
 */
    private fun parseFmla(complexrange:String):Stack<Comparable<*>> {
val ops = Stack<Comparable<*>>()
var lastOp = 0
var finishRange = false
var refs = Stack<Comparable<*>>()
var ref = ""
var inquote = false
var range:String? = null    // holds partial range
var i = 0
while (i < complexrange.length)
{
val c = complexrange.get(i)
if (c == '\'')
{
inquote = !inquote
ref += c
}
else if (!inquote)
{
if (c == ',' || c == ' ' || c == ')' || (c == ':' && finishRange))
{    // it's an operand
if (c == ' ' && lastOp == ' '.toInt())
{
i++
continue
}    // skip 2nd space op (Isect)
if (finishRange)
{ // add ref to rest of range
refs.push(range!! + ref)
if (!refs.isEmpty() && !ops.isEmpty())
refs = handleOpearatorPreference(refs, ops)
while (!ops.isEmpty())
refs.push(ops.pop())
range = null
ref = ""
finishRange = false
ops.push(c)
}
else if (refs.isEmpty())
{    // no operands yet - put in 1st
if (ref != "")
refs.push(ref)
ref = ""
ops.push(c)
}
else
{    // have all we need to process
if (ref != "")
refs.push(ref)
while (!ops.isEmpty())
refs.push(ops.pop())
ref = ""    // handle case of two spaces ... unfortunately
ops.push(c)
}
lastOp = c.toInt()
}
else if (c == ':')
{
if (this.parentRec!!.workBook!!.getName(ref) == null)
{ // it's a regular range
 // check if the ref is a sheet name in a 3d ref
                        if (ref != "")
{
range = ref + c
finishRange = true
}
else
{ // happens in cases such as (opopop):ref:ref
ops.push(c)
}
ref = ""
}
else
{    // it's a named range
refs.push(ref)
ref = ""
ops.push(c)
finishRange = false    // it's not a regular range
}
}
else if (c == '(')
{
var endparen = FormulaParser.getMatchOperator(complexrange, i, '(', ')')
if (endparen == -1)
endparen = complexrange.length - 1
else if (ref != "")
{
 // rare case of a PtgMemFunc containing a formula:
                        val f = ref + "(" + complexrange.substring(i + 1, endparen + 1)
ref = ""
refs = mergeStacks(refs, FormulaParser.getPtgsFromFormulaString(this.parentRec, f, true))
i = endparen
if (!ops.isEmpty())
refs = handleOpearatorPreference(refs, ops)
while (!ops.isEmpty())
refs.push(ops.pop())
i++
continue
}
refs = mergeStacks(refs, parseFmla(complexrange.substring(i + 1, endparen + 1)))
i = endparen
if (!ops.isEmpty())
refs = handleOpearatorPreference(refs, ops)
while (!ops.isEmpty())
refs.push(ops.pop())
}
else
ref += c
}
else
ref += c
i++
}
 // get any remaining
        if (finishRange)
{ // add ref to rest of range
 // range op has more precedence than others ...
            if ((!ops.isEmpty() && (ops.peek() as Char).charValue() == ':' && !refs.isEmpty() &&
refs.peek() is Char))
{
while (refs.peek() is Char)
{
if ((refs.peek() as Char).charValue() != ':')
ops.add(0, refs.pop())
else
break
}
}
if (ref != "")
refs.push(range!! + ref)
else
{
refs.push(range!!.substring(0, range!!.length - 1))
ops.push(':')
}
}
else
{
if (ref != "")
refs.push(ref)
}
while (!ops.isEmpty())
refs.push(ops.pop())
return refs
}


/**
 * when parenthesed sub-functions
 *
 * @param first
 * @param last
 * @return
 */
    private fun mergeStacks(first:Stack<Comparable<*>>, last:Stack<Comparable<*>>):Stack<Comparable<*>> {
first.addAll(last)
return first
}

/**
 * traverse through expression to retrieve set of ranges
 * either discontiguous union (,), intersected ( ) or regular range (:)
 */
    public override fun toString():String {
return FormulaParser.getExpressionString(subExpression!!).substring(1)    // avoid "="
}

/**
 * return the boundsheet associated with this complex range
 * <br></br>NOTE: since complex ranges may contain more than one sheet, this is incomplete for those instanaces
 *
 * @param b
 * @return
 */
     fun getSheets(b:WorkBook):Array<Boundsheet>? {
if (ptgs == null)
parseSubexpression()
if (this.refsheets != null || this.refsheets!!.size != 0)
{
try
{
val sheets = arrayOfNulls<Boundsheet>(this.refsheets!!.size)
for (i in sheets.indices)
sheets[i] = b.getWorkSheetByName(refsheets!!.get(i))
return sheets
}
catch (e:Exception) {
 // TODO: report error?
            }

}
return null

}

public override fun close() {
if (ptgs != null)
{
for (i in ptgs!!.indices)
{
if (ptgs!![i] is PtgRef)
ptgs!![i].close()
else
ptgs!![i].close()
}
}
ptgs = null
super.close()
}

companion object {

 val serialVersionUID = 666555444333222L


/**
 * handle precedence of complex range operators:  : before , before ' '
 *
 * @param sourceStack
 * @param destStack
 */
    private fun handleOpearatorPreference(refs:Stack<Comparable<*>>, ops:Stack<Comparable<*>>):Stack<Comparable<*>> {
val lastOp = (ops.pop() as Char).charValue()
if (refs.peek() is Char)
{
val curOp = (refs.pop() as Char).charValue()
val group1 = rankPrecedence(lastOp)
val group2 = rankPrecedence(curOp)
if (group2 >= group1)
{
ops.push(lastOp)
refs.push(curOp)
}
else
{
ops.push(curOp)
refs.push(lastOp)
}

}
else
ops.push(lastOp)
return refs
}

/**
 * rank a Ptg Operator's precedence (lower
 *
 * @param curOp
 * @return
 */
    internal fun rankPrecedence(curOp:Char):Int {
if (curOp.toInt() == 0) return -1
if (curOp == ')')
return 6
if (curOp == ':')
return 5
if (curOp == ',' || curOp == ' ')
 // same level????
            return 4
return 0    // ' '
}
}
 /*	protected void finalize() {
		this.close();
	}*/
}





