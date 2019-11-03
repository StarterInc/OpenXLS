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

import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.*
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool

import java.util.Locale
import java.util.Stack
import java.util.Vector

//import com.sun.org.apache.xerces.internal.impl.xs.identity.Selector.Matcher;


/**
 * Formula Parser.
 *
 *
 * Translates Excel-compatible Strings into Biff8/OpenXLS Compatible Formulas and vice-versa.
 *
 * @see Formula
 */
object FormulaParser {


    internal var DEBUGLEVEL = -1
    // KSC: handling unary operators necessitated, in the end, a complete rewrite ...

    /**
     * getPtgsFromFormulaString
     * returns ordered stack of Ptgs parsed from formula string fmla
     *
     * @param    Formula form	formula record
     * @param    String fmla		string rep. of current state of formula (either original "=F(Y)" or a recurrsed state e.g. "Y")
     * @returns Stack    ordered Ptgs that represent formula expression
     */
    fun getPtgsFromFormulaString(form: XLSRecord, fmla: String): Stack<*> {
        return getPtgsFromFormulaString(form, fmla, true)
    }

    /**
     * getPtgsFromFormulaString is the main entry point for parsing a string and creating a formula.
     * The formula passed in at this point can either be an existing formula with an expression, or a
     * templated formula with no expression. The string gets parsed and entered
     * as the expression for the formula.
     *
     * @param    Formula form	formula record
     * @param    String fmla		string rep. of current state of formula (either original "=F(Y)" or a recurrsed state e.g. "Y")
     * @param    boolean bIsCompleteExpression	truth of "formula fmla represents a complete formula i.e. we are not currently in a recurrsed state"
     * @returns Stack    ordered Ptgs that represent formula expression
     */
    /**
     * should handle all sorts of variations of formula strings such as:
     * =PV(C17,1+-(1*1)-9, 0, 1)
     * =100*0.5
     * =(B2-B3)*B4
     * =SUM(IF(A1:A10=B1:B10, 1, 0))
     * =IF(B4<=10,"10", if(b4<=100, "15", "20"))
     * ="STRING"&IF(A<>"",A,"N/A")&" - &IF(C<>"",C,"N/A")&" Result "
     *
     *
     * in basic essence, handles signatures such as
     * a op f(b op c, d, uop e ...) op g
     *
     *
     * where op is any binary operator, uop is a unary operator f is a formula
     * ...
     */
    internal fun getPtgsFromFormulaString(form: XLSRecord, fmla: String, bMergeWithLast: Boolean): Stack<*> {
        var fmla = fmla
        var bMergeWithLast = bMergeWithLast
        val operands = arrayOfNulls<Any>(2)

        fmla = fmla.trim { it <= ' ' }
        if (fmla.startsWith("=")) fmla = fmla.substring(1)
        fmla = fmla.trim { it <= ' ' }


        if (fmla == "") {    // 20081120 KSC: Handle Missing Argument
            val s = Stack()
            s.add(PtgMissArg())
            return s
        }

        if (fmla.startsWith("{")) { // must process array formula first, PtgArray expects full function string
            val pa = PtgArray()
            pa.parentRec = form
            val endarray = getMatchOperator(fmla, 0, '{', '}')
            pa.setVal(fmla.substring(0, endarray + 1))
            fmla = fmla.substring(endarray + 1)
            val s = Stack()
            s.add(pa)
            operands[0] = s
            bMergeWithLast = false
        }
        // TODO: complex ranges??
        var inQuote = false
        var inRange = false
        var inOp = false
        var prefix = ""
        val ops = Stack()
        var op = ""
        var i = 0
        while (i < fmla.length) {
            val c = fmla[i]
            if (c == '"' || c == '\'') { // get to ending quote
                if (inQuote) {
                    inQuote = c != prefix.trim { it <= ' ' }.get(0)    // if start quote == end quote, inQuote is false
                } else
                    inQuote = true
                if (inQuote) {
                    if (inOp) {
                        inOp = false
                        if (op != "") ops.add(0, op)
                        op = ""
                        if (operands[0] != null && operands[1] != null)
                            addOperands(form, /*functionStack, */operands, ops)
                    }
                }
                prefix += c
            } else if (inQuote) {
                prefix += c
            } else if (c == ':') {
                if (i > 0)
                    inRange = true
                prefix += c
            } else if (c == '(') {        // found a formula?? check out
                // if the parenthesis is part of a complex range, keep going i.e. keep entire expression together
                if (inRange) {
                    prefix += c
                    i++
                    continue
                }
                if (inOp) {
                    inOp = false
                    if (op != "") ops.add(0, op)
                    op = ""
                    if (!ops.isEmpty() && operands[0] != null && operands[1] != null) {
                        addOperands(form, /*functionStack, */operands, ops)
                    }
                }
                var funcName = ""
                for (k in prefix.length - 1 downTo 0) {
                    if (Character.isLetterOrDigit(prefix[k]) || Character.toString(prefix[k]) == ".") {
                        funcName = prefix[k] + funcName
                        prefix = prefix.substring(0, k)
                    } else
                        break
                }
                // prefix= anything before function name
                if (prefix.trim { it <= ' ' } != "") {
                    if (operands[0] == null)
                        operands[0] = prefix.trim { it <= ' ' }
                    else
                        operands[1] = prefix.trim { it <= ' ' }
                    prefix = ""
                }
                // function name should = part just before parents
                if (funcName != "") {
                    var funcPtg: Ptg? = null
                    funcPtg = getFuncPtg(funcName, form)

                    // do we have a valid function Ptg?
                    if (funcPtg != null) {    // yes, then handle function paramters i.e. evertyhing between the parentheses
                        val endparen = getMatchOperator(fmla, i, '(', ')')
                        if (endparen < fmla.length - 1) {
                            if (fmla[endparen + 1] == ':') {    // it's a VERY complex range :)
                                inRange = true
                                prefix = funcName + fmla.substring(i, endparen + 1)    // keep function name together ...
                                i = endparen
                                i++
                                continue
                            }
                        }
                        // things like: xyz + f(x)
                        if (!ops.isEmpty() && operands[0] != null && operands[1] != null) {
                            addOperands(form, /*functionStack, */operands, ops)
                            inOp = false
                        }    // have [xyz, xyz, OP] + [abc, def, ghi]
                        // parse function
                        val s = parseFunctionPtg(form, funcName, fmla.substring(i + 1, endparen), funcPtg)
                        if (operands[0] == null)
                            operands[0] = s
                        else
                            operands[1] = s
                        //functionStack.addAll(parseFunctionPtg(form, funcName, fmla.substring(i+1, endparen), funcPtg));
                        i = endparen    // inc. pointer to past processing point
                    } else
                    // else, we have *something* in front of the parentheses ...
                        throw FunctionNotSupportedException("$funcName is not a supported function")
                } else {        // enclosing parens
                    // complexities occur for complex ranges and enclosing parens ...
                    val endparen = getMatchOperator(fmla, i, '(', ')')
                    if (endparen == -1 || endparen < fmla.length - 1 && fmla[endparen + 1] == ':') {    // it's a VERY complex range :)
                        inRange = true
                        prefix = "("
                        i++
                        continue
                    }
                    val f = fmla.substring(i + 1, endparen)    // the statement less the parenthesis
                    i = endparen    // skip parens ...
                    // see if the enclosed expression is a complex range - must parse as 1 unit, rather than parsing particular ptgs
                    if (FormulaParser.isComplexRange("($f)")) {
                        val s = Stack()
                        s.push(parseSinglePtg(form, "($f)", false))
                        s.push(parseSinglePtg(form, ")", false))
                        if (operands[0] == null)
                            operands[0] = s
                        else
                            operands[1] = s
                    } else {    // embedded functions, keep parsing
                        val s = getPtgsFromFormulaString(form, f, true)    // flag as a complete expression
                        s.push(PtgParen())    // add ending parens to stack
                        if (operands[0] == null)
                            operands[0] = s
                        else
                            operands[1] = s
                        if (!ops.isEmpty()) {
                            addOperands(form, /*functionStack, */operands, ops)
                        }
                    }
                }
            } else {    // see if we have found an operataor
                if (!Character.isJavaIdentifierPart(c) && c != ' ' && c != '%') {
                    //  if (inRange && !Character.isJavaIdentifierPart(c) && c!=',' && c!=' ')
                    if (inRange) {
                        if (c != ',' && c != ' ' && c != ')' && c != '!') {
                            inRange = false
                            if (prefix.trim { it <= ' ' } != "") {
                                if (operands[0] == null)
                                    operands[0] = prefix.trim { it <= ' ' }
                                else
                                    operands[1] = prefix.trim { it <= ' ' }
                                prefix = ""
                            }
                        } else {
                            prefix += c
                            i++
                            continue
                        }
                    }
                    if (c != '!' && c != '#' && c != '.') {    // ignore !
                        // FOUND AN OPERATOR - ready to add operands yet?
                        inOp = true
                        if (prefix.trim { it <= ' ' } != "") {
                            if (operands[0] == null)
                                operands[0] = prefix.trim { it <= ' ' }
                            else
                                operands[1] = prefix.trim { it <= ' ' }
                            prefix = ""
                        }
                        if (!ops.isEmpty() && operands[0] != null && operands[1] != null) {
                            addOperands(form, /*functionStack, */operands, ops)
                        }
                        if (op != "") {
                            if (!(c == '=' || c == '>')) {
                                ops.add(0, op)
                                op = ""
                            }
                        }
                        op += c    // >,<,-,/, ,+
                        i++
                        continue
                    }
                } else if (inOp) {
                    inOp = false
                    if (!ops.isEmpty() && operands[0] != null && operands[1] != null) {
                        addOperands(form, /*functionStack, */operands, ops)
                    }
                    if (op != "")
                        ops.add(0, op)
                    op = ""
                }
                prefix += c
            }
            i++

        }
        if (prefix.trim { it <= ' ' } != "") {    // get any remaining elements
            if (operands[0] == null)
                operands[0] = prefix.trim { it <= ' ' }
            else
                operands[1] = prefix.trim { it <= ' ' }
            prefix = ""
        }

        if (op != "")
            ops.add(0, op)
        addOperands(form, operands, ops)
        //if (((Stack)operands[0]).isEmpty())
        //return functionStack;
        return operands[0] as Stack<*>

    }

    /**
     * Given two operands (objects) and operators (can be up to 2 if there is a unary operator present)
     * organize and add to functionStack in reverse polish notation i.e.
     * OPERAND, OPERAND, OP [...]
     *
     * @param form
     * @param functionStack
     * @param operands
     * @param ops
     */
    private fun addOperands(form: XLSRecord, /*Stack functionStack, */operands: Array<Any>, ops: Stack<*>) {
        var pOp: Ptg? = null
        if (!ops.isEmpty())
            pOp = parseSinglePtg(form, ops.pop() as String, operands[1] == null)

        var s = Stack()
        s.addAll(handleOperatorPrecedence(form, /*functionStack, */operands, pOp))
        operands[0] = s
        //functionStack.clear();
        if (!ops.isEmpty()) {
            pOp = parseSinglePtg(form, ops.pop() as String, true)
            s = Stack()
            s.addAll(handleOperatorPrecedence(form, /*functionStack, */operands, pOp))
            operands[0] = s
        }

    }

    private fun handleOperatorPrecedence(form: XLSRecord, /*Stack functionStack, */operands: Array<Any>, pOp: Ptg?): Stack<*> {
        var pOp = pOp
        var functionStack = Stack()
        if (operands[0] is Stack<*>) {
            functionStack = operands[0] as Stack<*>
            operands[0] = null
        }
        if (!functionStack.isEmpty() && pOp != null) {
            val lastOp = functionStack.peek() as Ptg
            if (lastOp != null && lastOp.isOperator) {
                if (lastOp.isOperator) {
                    functionStack.pop()    //= lastOp
                    val group1 = rankPrecedence(pOp)
                    val group2 = rankPrecedence(lastOp)
                    if (group2 >= group1) {
                        functionStack.push(lastOp)
                    } else {    // current op has higher priority
                        if (operands[0] != null) {
                            if (operands[0] is String)
                                functionStack.push(parseSinglePtg(form, operands[0] as String, functionStack.isEmpty()))
                            else
                                functionStack.addAll(operands[0] as Stack<*>)
                        }
                        if (operands[1] != null) {
                            if (operands[1] is String)
                                functionStack.push(parseSinglePtg(form, operands[1] as String, functionStack.isEmpty()))
                            else
                                functionStack.addAll(operands[1] as Stack<*>)
                        }
                        operands[0] = null
                        operands[1] = null
                        functionStack.push(pOp)
                        pOp = lastOp
                    }
                }
            }
        }
        if (operands[0] != null) {
            if (operands[0] is String)
                functionStack.push(parseSinglePtg(form, operands[0] as String, functionStack.isEmpty()))
            else
                functionStack.addAll(operands[0] as Stack<*>)
        }
        if (operands[1] != null) {
            if (operands[1] is String)
                functionStack.push(parseSinglePtg(form, operands[1] as String, functionStack.isEmpty()))
            else
                functionStack.addAll(operands[1] as Stack<*>)
        }
        operands[0] = null
        operands[1] = null
        if (pOp != null)
            functionStack.push(pOp)
        return functionStack
    }


    /**
     * merge last stacks to ensure operator order is correct
     *
     * @param functionStack
     * @return
     */
    private fun mergeStacks(prevStack: Stack<*>, curStack: Stack<*>, bIsCompleteExpression: Boolean): Stack<*> {
        if (prevStack.isEmpty())
            return curStack

        var lastOp = prevStack.peek() as Ptg
        val curOp = if (curStack.isEmpty()) null else curStack.peek() as Ptg
        val group1 = rankPrecedence(lastOp)
        val group2 = rankPrecedence(curOp)
        if (group1 >= 0 && (group1 < group2 || group2 == -1)) {
            lastOp = prevStack.pop() as Ptg
            curStack.push(lastOp)
            /*  				while (curOp.getIsOperator()) {
  					lastOp= curOp;
  					curStack.push(curOp);
  					if (!prevStack.isEmpty()) {
  						curOp = (Ptg) prevStack.pop();
  						// handle precedence
  				  		group1=rankPrecedence(curOp);
  				  		group2=rankPrecedence(lastOp);
  				  		if (!(group1>=0 && (group1 < group2 || group2==-1)))
  				  			break;
  					}else
  						return curStack;
  				}
*/
        }
        prevStack.addAll(curStack)
        return prevStack
    }


    /**
     * parse and add to Stack a valid Excel function represented by funcPtg and fmla string
     * called from getPtgsFromFormulaString
     *
     * @param form    formula record
     * @param fmla    function parameters in the form of (x, y, z)
     * @param funcPtg function data for the formula represented by fmla
     * @para func        function name f
     * @return Stack    ordered parsed Stack of Ptgs
     */
    private fun parseFunctionPtg(form: XLSRecord, func: String, fmla: String, funcPtg: Ptg): Stack<*> {
        var fmla = fmla
        val returnStack = Stack()
        fmla = fmla.trim { it <= ' ' }
        var nParens = 0
        var enclosing = false
        // change:  only remove 1 set of parens:
        if (fmla.length > 0 && fmla[0] == '(') {
            if (getMatchOperator(fmla, 0, '(', ')') == fmla.length - 1) {    // then strip enclosing parens
                nParens++
                enclosing = true
            }
            fmla = fmla.trim { it <= ' ' }
        }

        // NOTE: all memfuncs/complex ranges are enclosed by parentheses
        // IF enclosed by parens, DO NOT split apart into operands:
        var funcLen = 1
        if (enclosing) {
            returnStack.addAll(FormulaParser.getPtgsFromFormulaString(form, fmla, true))
        } else {
            val cv = splitFunctionOperands(fmla)
            funcLen = cv.size
            // loop through the operands to the function and recurse
            for (y in cv.indices) {
                val s = cv.elementAt(y) as String        // flag as a complete expression
                returnStack.addAll(FormulaParser.getPtgsFromFormulaString(form, s, true))
            }
        }

        // Handle PtgFuncVar-specifics such as number of parameters and add-in PtgNameX record
        if (funcPtg is PtgFuncVar) {
            if (funcPtg.`val` == FunctionConstants.xlfADDIN) {
                // if an add-in, must add PtgNameX to stack
                val pn = PtgNameX()
                pn.parentRec = form
                pn.setName(func)
                returnStack.add(0, pn)    // add to bottom of stack
                funcLen++
                funcPtg.parentRec = form // nec. to resolve external name
            }
            funcPtg.setNumParams(funcLen.toByte())
        }
        returnStack.push(funcPtg)
        return returnStack
    }

    /**
     * combine two stacks of Ptgs, popping the operator of the sourceStack and
     * adding it to the end of the destination stack to
     * ensure it is in the correct order in the destination stack
     *
     * @param sourceStack
     * @param destStack
     * /
     * private static Stack addPtgStacks(Stack sourceStack, Stack destStack) {
     * Ptg opPtg = (Ptg) sourceStack.pop();
     * Ptg lastOp= (destStack.isEmpty()?null:((Ptg)destStack.peek()));
     *
     * // handle precedence:  unaries before ^(power) before *, / before +, - before &(concat), before comparisons (=, <>, <=, >=, <, >)
     * int group1=rankPrecedence(opPtg);
     * int group2=rankPrecedence(lastOp);
     *
     * if (group1>=0 && (group1 < group2 || group2==-1)) {
     * while (opPtg.getIsOperator()) {
     * lastOp= opPtg;
     * destStack.push(opPtg);
     * if (!sourceStack.isEmpty()) {
     * opPtg = (Ptg) sourceStack.pop();
     * // handle precedence
     * group1=rankPrecedence(opPtg);
     * group2=rankPrecedence(lastOp);
     * if (!(group1>=0 && (group1 < group2 || group2==-1)))
     * break;
     * }else
     * return destStack;
     * }
     * }
     * sourceStack.push(opPtg);
     *
     * // after sorting out operators, assemble two stacks into one
     * Stack nwstack = destStack;
     * destStack = new Stack();
     * destStack.addAll(sourceStack);
     * destStack.addAll(nwstack);
     *
     * return destStack;
     * }
     */

    /**
     * rank a Ptg Operator's precedence (lower
     *
     * @param curOp
     * @return
     */
    internal fun rankPrecedence(curOp: Ptg?): Int {
        if (curOp == null) return -1
        //		if (curOp==null || !curOp.getIsOperator()) return -1;
        if (curOp is PtgUMinus || curOp is PtgUPlus)
            return 7
        if (curOp is PtgPercent)
            return 6
        if (curOp is PtgPower)
            return 5
        if (curOp is PtgMlt || curOp is PtgDiv)
            return 4
        else if (curOp is PtgAdd || curOp is PtgSub)
            return 3
        else if (curOp is PtgConcat)
            return 2
        else if (curOp is PtgEQ || curOp is PtgNE ||
                curOp is PtgLE || curOp is PtgLT ||
                curOp is PtgGE || curOp is PtgGT)
            return 1
        //  		else if (curOp instanceof PtgParen)
        //  			return 0;
        return -1
    }

    /*
     * getMatchOperator takes a string and starting operator location.
     * It then parses the string and determines which closing operator
     * matches the opening parens specified by startParenLoc.  Returns
     * -1 if it cannot find a match.
     */
    fun getMatchOperator(input: String, startParenLoc: Int, matchOpenChar: Char, matchCloseChar: Char): Int {
        // 20081112 KSC: do a different way as it wasn't working for all cases
        var openCnt = 0
        var i = startParenLoc
        while (i < input.length) {
            if (input[i] == '"' || input[i] == '\'') {// handle quoted strings within input (quoted strings may of course contain match chars ...
                val endquote = input[i]
                while (++i < input.length) {
                    if (input[i] == endquote) {
                        break
                    }
                }
            }
            if (i == input.length) return i - 1
            if (input[i] == matchOpenChar)
                openCnt++
            else if (input[i] == matchCloseChar) {
                openCnt--
                if (openCnt == 0)
                    return i
            }
            i++
        }

        // no parens for you!
        return -1
    }

    /**
     * Looks up a function string and returns a funcPtg if it is found
     *
     * @param    func    function string without parents i.e. SUM or DB
     * @returns Ptg        valid funcPtg or null if not found
     */
    // 20090210 KSC: add form so can set parent record for PtgFunc and PtgFuncVar - nec for self-referential formulas such as COLUMN
    private fun getFuncPtg(func: String, form: XLSRecord): Ptg? {
        var funcPtg: Ptg? = null
        //    if (true) {
        if (Locale.JAPAN == Locale.getDefault()) {
            for (y in FunctionConstants.jRecArr.indices) {
                if (func.equals(FunctionConstants.jRecArr[y][0], ignoreCase = true)) {
                    val FID = Integer.parseInt(FunctionConstants.jRecArr[y][1])
                    val Ftype = Integer.parseInt(FunctionConstants.jRecArr[y][2])
                    if (Ftype == FunctionConstants.FTYPE_PTGFUNC) {
                        funcPtg = PtgFunc(FID, form)
                    } else if (Ftype == FunctionConstants.FTYPE_PTGFUNCVAR) {
                        funcPtg = PtgFuncVar(FID, 0, form)
                    } else if (Ftype == FunctionConstants.FTYPE_PTGFUNCVAR_ADDIN) {
                        funcPtg = PtgFuncVar(FunctionConstants.xlfADDIN, 0, form)
                    }
                    return funcPtg
                }
            }
        }
        for (y in FunctionConstants.recArr.indices) {
            if (func.equals(FunctionConstants.recArr[y][0], ignoreCase = true)) {
                val FID = Integer.parseInt(FunctionConstants.recArr[y][1])
                val Ftype = Integer.parseInt(FunctionConstants.recArr[y][2])
                if (Ftype == FunctionConstants.FTYPE_PTGFUNC) {
                    funcPtg = PtgFunc(FID, form)
                } else if (Ftype == FunctionConstants.FTYPE_PTGFUNCVAR) {
                    funcPtg = PtgFuncVar(FID, 0, form)
                } else if (Ftype == FunctionConstants.FTYPE_PTGFUNCVAR_ADDIN) {
                    funcPtg = PtgFuncVar(FunctionConstants.xlfADDIN, 0, form)
                }
                return funcPtg
            }
        }
        for (y in FunctionConstants.unimplRecArr.indices) {
            if (func.equals(FunctionConstants.unimplRecArr[y][0], ignoreCase = true)) {
                val FID = Integer.parseInt(FunctionConstants.unimplRecArr[y][1])
                val Ftype = Integer.parseInt(FunctionConstants.unimplRecArr[y][2])
                if (Ftype == FunctionConstants.FTYPE_PTGFUNC) {
                    funcPtg = PtgFunc(FID, form)
                } else if (Ftype == FunctionConstants.FTYPE_PTGFUNCVAR) {
                    funcPtg = PtgFuncVar(FID, 0, form)
                } else if (Ftype == FunctionConstants.FTYPE_PTGFUNCVAR_ADDIN) {
                    funcPtg = PtgFuncVar(FunctionConstants.xlfADDIN, 0, form)
                }
                return funcPtg
            }
        }
        return funcPtg

    }


    /**
     * take a string guaranteed to be a single Ptg (operator, reference, string, etc) and convert to correct Ptg
     *
     * @param form
     * @param fmla
     * @param bIsUnary -- operator is a unary version
     * @return
     */
    private fun parseSinglePtg(form: XLSRecord, fmla: String, bIsUnary: Boolean): Ptg? {
        val bk = form.workBook    // nec. to determine if parsed element is a valid name handle name

        var `val` = fmla
        var name = convertString(`val`, bk)
        if (name == "+" && bIsUnary)
            name = "u+"
        else if (name == "-" && bIsUnary)
            name = "u-"

        var pthing: Ptg? = null
        try {
            pthing = XLSRecordFactory.getPtgRecord(name)
            if (pthing == null && name == "PtgName")
            // TODO: MUST evaluate which type of PtgName is correct: understand usage!
                if (form.opcode == XLSConstants.FORMULA || form.opcode == XLSConstants.ARRAY)
                    pthing = PtgName(0x43)    // assume this token to be of type Value (i.e PtgNameV) instead of Reference (PtgNameR)
                else
                // DV needs ref-type name
                    pthing = PtgName(0x23)    // PtgNameR
        } catch (e: InvalidRecordException) {
            Logger.logInfo("parsing formula string.  Invalid Ptg: $name error: $e")
        }

        // if it is an operator we don't need to do anything with it!
        if (pthing != null) {
            pthing.parentRec = form

            if (!pthing.isOperator) {
                if (pthing.isReference) {
                    // createPtgRefFromString will handle any type of string reference
                    // will return a PtgRefErr if cannot parse location
                    pthing = PtgRef.createPtgRefFromString(`val`, form)
                } else if (pthing is PtgStr) {
                    val pstr = pthing as PtgStr?
                    `val` = StringTool.strip(`val`, '\"')
                    pstr!!.`val` = `val`
                } else if (pthing is PtgNumber) {
                    val pnum = pthing as PtgNumber?
                    if (`val`.indexOf("%") == -1)
                        pnum!!.`val` = Double(`val`)
                    else
                        pnum!!.setVal(`val`)
                } else if (pthing is PtgInt) {
                    val pint = pthing as PtgInt?
                    pint!!.`val` = Integer.valueOf(`val`).toInt()
                } else if (pthing is PtgBool) {
                    val pbool = pthing as PtgBool?
                    pbool!!.setVal(java.lang.Boolean.valueOf(`val`).booleanValue())
                } else if (pthing is PtgArray) {
                    val parr = pthing as PtgArray?
                    parr!!.setVal(`val`)
                } else if (pthing is PtgName) { // SHOULD really return PtgErr("#NAME!") as it's a missing Name instead of adding a new name
                    val pname = pthing as PtgName?
                    pname!!.setName(`val`)
                    pname.addToRefTracker()
                } else if (pthing is PtgNameX) {
                    val pnameX = pthing as PtgNameX?
                    pnameX!!.setName(`val`)
                } else if (pthing is PtgMissArg) {
                    pthing.init(byteArrayOf(22))
                } else if (pthing is PtgErr) {
                    pthing = PtgErr(PtgErr.convertStringToLookupByte(`val`))
                } else if (pthing is PtgAtr) {
                    pthing = PtgAtr(0x40.toByte())    // assume space
                }
            }
        } else {
            val pname = PtgMissArg()
        }
        return pthing
    }

    private fun findPtg(fmla: String, bUnary: Boolean): String {
        val s = StringTool.allTrim(fmla)

        if (s.startsWith("\"") || s.startsWith("'")) {
            //return s;	// it's a string
            return s
        }

        for (i in XLSRecordFactory.ptgOps.indices) {
            var ptgOpStr = XLSRecordFactory.ptgOps[i][0]

            val x = s.indexOf(ptgOpStr)
            if (x == 0) { // found instance of an operator
                // if encounter a parenthesis, must determine if it is an expression limit OR
                // if it is part of a complex range, in which case the expression must be kept together
                if (ptgOpStr == "(") {
                    return s    // parens means a whole complex range==>PtgMemFunc
                }
                if (bUnary) {
                    // unary ops +, - ... have a diff't Ptg than regular vers of the operator
                    if (ptgOpStr == "-" && x == 1 && ptgOpStr.length > 1) break //negative number, NOT a unary -
                    for (j in XLSRecordFactory.ptgPrefixOperators.indices) {
                        if (ptgOpStr.startsWith(XLSRecordFactory.ptgPrefixOperators[j][0].toString())) {
                            ptgOpStr = XLSRecordFactory.ptgPrefixOperators[j][1].toString()
                        }
                    }
                }
                return XLSRecordFactory.ptgOps[i][1]
            }
        }
        return s
    }

    /**
     * parseFinalLevel is where strings get converted into ptg's
     * This method can handle multiple ptg's within a string, but cannot handle
     * recursion.  If you are having recursion problems look into getPtgsFromFormulaString above.
     *
     *
     * This method should be called from the final level of parsing.
     * There should not be additional sub-expressions at this point.
     * Example (1+2) or (3,4,5)
     * NOT ((1<2),3,4) or TAN(23);
     *
     *
     *
     *
     * TODO: HANDLE these references:
     *
     *
     * =SUM(table[[#This Row];['#Head3]:[Calced]])
     *
     *
     * table[['#Head3]:[Calced]]
     *
     *
     * I assume this means a table of data, the Head3 table? and the calced column?
     */
    private fun parseFinalLevel(form: XLSRecord, fmla: String, bIsComplete: Boolean): Stack<*> {
        var returnStack = Stack()
        val parseThings = CompatibleVector()

        // break it up into components first
        var elements = Vector()
        elements = splitString(fmla, bIsComplete)

        val bk = form.workBook    // nec. to determine if parsed element is a valid name handle name

        // convert each element into Ptg's
        // each element at this point should be a named operand, or an unidentified operator
        for (x in elements.indices) {
            var `val` = elements.elementAt(x) as String
            val name = convertString(`val`, bk)

            var pthing: Ptg? = null
            try {
                pthing = XLSRecordFactory.getPtgRecord(name)
                if (pthing == null && name == "PtgName")
                // TODO: MUST evaluate which type of PtgName is correct: understand usage!
                    if (form.opcode == XLSConstants.FORMULA || form.opcode == XLSConstants.ARRAY)
                        pthing = PtgName(0x43)    // assume this token to be of type Value (i.e PtgNameV) instead of Reference (PtgNameR)
                    else
                    // DV needs ref-type name
                        pthing = PtgName(0x23)    // PtgNameR
            } catch (e: InvalidRecordException) {
                Logger.logInfo("parsing formula string.  Invalid Ptg: $name error: $e")
            }

            // if it is an operator we don't need to do anything with it!
            if (pthing != null) {
                pthing.parentRec = form

                if (!pthing.isOperator) {
                    if (pthing.isReference) {
                        // createPtgRefFromString will handle any type of string reference
                        // will return a PtgRefErr if cannot parse location
                        pthing = PtgRef.createPtgRefFromString(`val`, form)
                    } else if (pthing is PtgStr) {
                        val pstr = pthing as PtgStr?
                        `val` = StringTool.strip(`val`, '\"')
                        pstr!!.`val` = `val`
                    } else if (pthing is PtgNumber) {
                        val pnum = pthing as PtgNumber?
                        if (`val`.indexOf("%") == -1)
                            pnum!!.`val` = Double(`val`)
                        else
                            pnum!!.setVal(`val`)
                    } else if (pthing is PtgInt) {
                        val pint = pthing as PtgInt?
                        pint!!.`val` = Integer.valueOf(`val`).toInt()
                    } else if (pthing is PtgBool) {
                        val pbool = pthing as PtgBool?
                        pbool!!.setVal(java.lang.Boolean.valueOf(`val`).booleanValue())
                    } else if (pthing is PtgArray) {
                        val parr = pthing as PtgArray?
                        parr!!.setVal(`val`)
                    } else if (pthing is PtgName) { // SHOULD really return PtgErr("#NAME!") as it's a missing Name instead of adding a new name
                        val pname = pthing as PtgName?
                        pname!!.setName(`val`)
                    } else if (pthing is PtgNameX) {
                        val pnameX = pthing as PtgNameX?
                        pnameX!!.setName(`val`)
                    } else if (pthing is PtgMissArg) {
                        pthing.init(byteArrayOf(22))
                    } else if (pthing is PtgErr) {
                        pthing = PtgErr(PtgErr.convertStringToLookupByte(`val`))
                    }
                }
                parseThings.add(pthing)
            } else {
                val pname = PtgMissArg()
            }

        }
        //reorder in polish notation and add to stack.
        // 20081128 KSC: Do later as reordering will depend upon position of this segment in formula returnStack = reorderStack(parseThings);  see getPtgsFromFormulaString
        returnStack = convertToStack(parseThings)
        return returnStack
    }

    /*
    private static Stack reorderStack(Stack sourceStack, boolean bIsComplete) {
    	Stack returnStack = new Stack();
    	Stack pOperators= new Stack();
		for (int x = 0;x<sourceStack.size(); x++){
			Ptg pthing  = (Ptg)sourceStack.get(x);
			if (pthing.getIsOperator()){	// Must account for precedence of operators
				if (bIsComplete && !returnStack.isEmpty()) {
					// handle infrequent cases of two operators in a row i.e. one op and one unary op
					// e.g. 1--1
					while (pOperators.size() > 0) {
						returnStack.push(pOperators.pop());
					}
  					if (((Ptg)returnStack.peek()).getIsOperator()) {
						int precedence=  rankPrecedence(pthing);
						Ptg p= (Ptg) returnStack.pop();
						int prevprecedence= rankPrecedence(p);
				  		if (precedence>prevprecedence) {
							pOperators.push(p);	 // switch less precedence operator with greater
				  		}
				  		else
							returnStack.push(p); // back to normal
					}
				} else if (!pOperators.isEmpty()) {	// added for instances such as (2)^-(2);
					//code below prevents the switching of the operators ^-
					Ptg p= (Ptg) pOperators.pop();
					pOperators.push(pthing);	// save operators
					pthing= p;
				}
				pOperators.push(pthing);	// save operators
			}else{	// it's not an operand; put on stack and pop all operators thus far
				returnStack.push(pthing);
				while (pOperators.size() > 0) {
					returnStack.push(pOperators.pop());
				}
			}
		}
		while (pOperators.size() > 0) {
			returnStack.push(pOperators.pop());
		}
		return returnStack;
    }
    */


    /**
     * convert list of parseThings without reordering
     *
     * @param parseThings
     * @return
     */
    private fun convertToStack(parseThings: CompatibleVector): Stack<*> {
        val returnStack = Stack()
        val pOperators = Stack()
        for (x in parseThings.indices) {
            val pthing = parseThings.elementAt(x) as Ptg
            returnStack.push(pthing)
        }
        return returnStack
    }

    /*
     * Parses an internal string for a function, splitting out elements.
     * for instance,(1<2), 3, tan(5); should return
     * [(1<2)][3][tan(5)].  Currently just working off of commas, but this may change...
     *
     * One of the keys here is to not split on a comma from an internal function, for instance,
     * "IF((1<2),MOD(45,6),0) should not split between 45 & 6!  Note the badLocs vector that handles this.
     */
    private fun splitFunctionOperands(formStr: String): CompatibleVector {
        val locs = CompatibleVector()
        // if there are no commas then we don't have to do all of this...
        if (formStr == "") return locs    // KSC: Handle no parameters by returning a null vector

        // first handle quoted strings 20081111 KSC
        var loop = true
        var pos = 0
        val badLocs = CompatibleVector()
        while (loop) {
            var c = '"'
            var start = formStr.indexOf(c.toInt(), pos)
            if (start == -1) {        // process single quotes as well
                c = '\''
                start = formStr.indexOf(c.toInt(), pos)
            }
            if (start != -1) {
                var end = formStr.indexOf(c.toInt(), start + 1)
                end += 1  //include trailing quote
                // check for being a part of a reference ...
                if (end < formStr.length && formStr[end] == '!') {// then it's part of a reference
                    end++
                    while (end < formStr.length && loop) {
                        c = formStr[end]
                        if (!(Character.isLetterOrDigit(c) || c == ':' || c == '$') || c == '-' || c == '+')
                            loop = false
                        else
                            end++
                    }
                }
                for (y in start until end) {
                    //make sure it is not a segment of a previous operand, like <> and >;
                    badLocs.add(Integer.valueOf(y))
                }
                if (end == 0) { // means it didn't find an end quote
                    end = formStr.length - 1
                    loop = false
                } else {
                    pos = end
                    loop = true
                }
            } else {
                loop = false
            }
        }


        if (formStr.indexOf(",") == -1) {
            locs.add(formStr)
        } else {
            // Handle each parameter (delimited by ,)
            // fill the badLocs vector with string locations we should disregard for comma proccesing
            run {
                var i = 0
                while (i < formStr.length) {
                    val openparen = formStr.indexOf("(", i)
                    if (openparen != -1) {
                        if (!badLocs.contains(Integer.valueOf(openparen))) {
                            var closeparen = getMatchOperator(formStr, openparen, '(', ')')
                            if (closeparen == -1) closeparen = formStr.length
                            i = openparen
                            while (i < closeparen) {
                                val `in` = Integer.valueOf(i)
                                badLocs.add(`in`)
                                i++
                            }
                        } else {    // open paren nested in quoted string
                            i = openparen + 1
                        }
                    } else
                    // 20081112 KSC
                        break
                    i++
                }
            }
            // lets do the same for the array items
            run {
                var i = 0
                while (i < formStr.length) {
                    val openparen = formStr.indexOf("{", i)
                    if (openparen != -1) {
                        if (!badLocs.contains(Integer.valueOf(openparen))) {
                            var closeparen = getMatchOperator(formStr, openparen, '{', '}')
                            if (closeparen == -1) closeparen = formStr.length
                            i = openparen
                            while (i < closeparen) {
                                val `in` = Integer.valueOf(i)
                                badLocs.add(`in`)
                                i++
                            }
                        } else {    // open paren nested in quoted string
                            i = openparen + 1
                        }
                    } else
                    // 20081112 KSC
                        break
                    i++
                }
            }
            // now check bad locations:
            var placeholder = 0
            var holder = 0
            while (holder != -1) {
                val i = formStr.indexOf(",", holder)
                if (i != -1) {
                    val ing = Integer.valueOf(i)
                    if (!badLocs.contains(ing)) {
                        val s = formStr.substring(placeholder, i)
                        locs.add(s)
                        placeholder = i + 1
                    }
                    holder = i + 1
                } else {
                    val s = formStr.substring(placeholder)
                    locs.add(s)
                    return locs
                }
            }
        }
        return locs
    }

    /**
     * parse a given string into known Ptg operators
     *
     * @param s
     * @return
     */
    private fun parsePtgOperators(s: String, bUnary: Boolean): CompatibleVector {
        var s = s
        var bUnary = bUnary
        val ret = CompatibleVector()
        s = StringTool.allTrim(s)

        for (i in XLSRecordFactory.ptgOps.indices) {
            var ptgOpStr = XLSRecordFactory.ptgOps[i][0]
            if (s.startsWith("\"") || s.startsWith("'")) {
                var end = s.substring(1).indexOf(s[0].toInt())
                end += 1  //include trailing quote
                // TEST IF The quoted item is a sheet name
                if (end < s.length && s[end] == '!') {// then it's part of a reference
                    end++
                    var loop = true
                    while (end < s.length && loop) {    // if the quoted string is a sheet ref, get rest of reference
                        val c = s[end]
                        if (c == '#' && s.endsWith("#REF!")) {
                            end += 5
                            loop = false
                        } else if (!(Character.isLetterOrDigit(c) || c == ':' || c == '$') || c == '-' || c == '+') {
                            loop = false
                        } else
                            end++
                    }
                }
                ret.add(s.substring(0, end + 1))
                s = s.substring(end + 1)
                bUnary = false
                if (s != "")
                    ret.addAll(parsePtgOperators(s, bUnary))
                break
            }

            var x = s.indexOf(ptgOpStr)
            if (x > -1) {    // found instance of an operator
                // if encounter a parenthesis, must determine if it is an expression limit OR
                // if it is part of a complex range, in which case the expression must be kept together
                if (ptgOpStr == "(") {
                    val end = getMatchOperator(s, x, '(', ')')
                    ret.add(s)    // add entire
                    break
                    /*					String ss= s.substring(x, end+1);
					ret.add(ss.substring(x));	// add entire
					if (FormulaParser.isComplexRange(ss)) {
//						ret.add(ss.substring(x+1));	// skip beginning paren as screws up later parsing
						ret.add(ss.substring(x));
						s= s.substring(end+1);
						bUnary= false;
						if (!s.isEmpty())
							ret.addAll(parsePtgOperators(s, bUnary));
						break;
					}
				} else if (ptgOpStr.equals(")")) {
					try {
						String ss= s.substring(x+1);
						char nextChar= ss.charAt(0);
						if (nextChar==' ') {	// see if there is another operand after the space
							ss= ss.trim();
							if (ss.length()>0 && ss.matches("[^(a-zA-Z].*")) {
								nextChar= ss.charAt(0);
							}
						}
						// complex ranges can contain parentheses in combo with these operators: :, (
						if (nextChar==' ' || nextChar==',' || nextChar==':' || nextChar==')')
							continue;	// keep complex range expression together
					} catch (Exception e) { ; }
*/
                }
                if (ptgOpStr == ")")
                // parens are there to keep expression together
                    continue
                if (x > 0) {// process prefix, if any - unary since it's the first operand
                    // exception here-- error range in the form of "Sheet!#REF! (eg) needs to be kept whole
                    if (!(XLSRecordFactory.ptgLookup[i][1] == "PtgErr" && s[x - 1] == '!')) {
                        ret.addAll(parsePtgOperators(s.substring(0, x), bUnary))
                        bUnary = false
                    } else {    // keep entire error reference together
                        ptgOpStr = s
                    }
                }
                x = x + ptgOpStr.length
                if (bUnary) {
                    // unary ops +, - ... have a diff't Ptg than regular vers of the operator
                    if (ptgOpStr == "-" && x == 1 && ptgOpStr.length > 1) break //negative number, NOT a unary -
                    for (j in XLSRecordFactory.ptgPrefixOperators.indices) {
                        if (ptgOpStr.startsWith(XLSRecordFactory.ptgPrefixOperators[j][0].toString())) {
                            ptgOpStr = XLSRecordFactory.ptgPrefixOperators[j][1].toString()
                        }
                    }
                }
                ret.add(ptgOpStr)
                if (x < s.length)
                // process suffix, if any
                    ret.addAll(parsePtgOperators(s.substring(x), true))
                break
            }
        }
        if (ret.isEmpty())
            ret.add(s)
        return ret
    }

    /*
     * Parses a string and returns an array based on contents
     * Assumed to be 1 "final-level" operand i.e a range, complex range, a op b[...]
     */
    private fun splitString(formStr: String, bIsComplete: Boolean): CompatibleVector {
        var formStr = formStr
        var bIsComplete = bIsComplete
        // Use a vector, and the collections methods to sort in natural order
        val locs = CompatibleVector()
        val retVect = CompatibleVector()

        // check for escaped string literals & add positions to vector if needed
        formStr = StringTool.allTrim(formStr)
        if (formStr == "") {
            retVect.add(formStr)
            return retVect
        }
        if (true) {
            retVect.addAll(parsePtgOperators(formStr, bIsComplete))        // cleanString if not an array formula????  s= cleanString(s);
            bIsComplete = false
            return retVect
        }

        // 20081207 KSC: redo completely to handle complex formula strings e.g. strings containing quoted commas, parens ...
        // first, pre-process to parse quoted strings, parentheses and array formulas
        var isArray = false
        var loop = true
        var s = ""
        var inRange = false
        var prevc: Char = 0.toChar()
        var i = 0
        while (i < formStr.length) {
            var c = formStr[i]
            if (c == '"' || c == '\'') {
                /*				if (!s.equals("")) {
					locs.add(s);
					s= "";
				}
*/
                var end = formStr.indexOf(c.toInt(), i + 1)
                end += 1  //include trailing quote
                // TEST IF The quoted item is a sheet name
                if (end < formStr.length && formStr[end] == '!') {// then it's part of a reference
                    end++
                    loop = true
                    while (end < formStr.length && loop) {    // if the quoted string is a sheet ref, get rest of reference
                        c = formStr[end]
                        if (c == '#' && formStr.endsWith("#REF!")) {
                            end += 5
                            loop = false
                        } else if (!(Character.isLetterOrDigit(c) || c == ':' || c == '$') || c == '-' || c == '+') {
                            loop = false
                        } else
                            end++
                    }
                }
                locs.add(s + formStr.substring(i, end))
                s += formStr.substring(i, end)
                i = end - 1
            } else if (c == '(') {    // may be a complex range if s==""
                if (s != "" && !inRange) {
                    //					char prevc= s.charAt(s.length()-1);
                    if (!(prevc == ' ' || prevc == ':' || prevc == ',' || prevc == '(')) {
                        locs.add(s)
                        s = ""
                    } else {    // DO NOT split apart complex ranges - they parse to PtgMemFuncs
                        //Logger.logInfo("FormulaParser.splitString:  PtgMemFunc" + formStr);
                        s += c
                        inRange = true
                    }
                }
            } else if (c == ':') {
                if (prevc == ')' && locs.size > 0)
                // complex range in style of: F(x):Y(x)
                    s = locs[locs.size - 1] as String + '('.toString() + s
                inRange = true
                s += c
            } else if (c == '{') {
                if (s != "") {
                    locs.add(s)
                    s = ""
                }
                var end = formStr.indexOf("}", i + 1)
                end += 1  //include trailing }
                locs.add(formStr.substring(i, end))
                i = end - 1
            } else
                s += c
            if (c != ' ')
                prevc = c
            i++
        }
        if (s != "") {
            locs.add(s)
            s = ""
        }

        // loop through the possible operator ptg's and get locations & length of them
        for (j in locs.indices) {
            s = locs[j] as String
            if (s.startsWith("\"") || s.startsWith("'"))
                retVect.add(s)    // quoted strings
            else {
                if (s.startsWith("{"))
                // it's an array formula
                    isArray = true // Do what?? else, cleanString??
                retVect.addAll(parsePtgOperators(s, bIsComplete))        // cleanString if not an array formula????  s= cleanString(s);
            }
            bIsComplete = false    // already parsed part of the formula string so cannot be unary :)
        }
        return retVect
    }

    /**
     * helper method that turns operator & operand strings into the Ptg\ equivalent
     * if there is no equivalant it leaves the string alone.
     */
    private fun convertString(ptg: String, bk: WorkBook?): String {
        // first check for operators
        for (i in XLSRecordFactory.ptgLookup.indices) {
            if (ptg.equals(XLSRecordFactory.ptgLookup[i][0], ignoreCase = true)) {
                return ptg
            }
        }

        // KSC: Added for missing arguments ("")
        if (StringTool.allTrim(ptg) == "")
        //return "PtgMissArg";
        //if (ptg.equals(""))
            return "PtgAtr"    // a space

        // Now we need to figure out what type of operand it is
        // see if it is a string, should be encased by ""
        if (ptg.substring(0, 1).equals("\"", ignoreCase = true)) {
            return "PtgStr"
        }
        // is it an array?
        if (ptg.substring(0, 1).equals("{", ignoreCase = true)) {
            return "PtgArray"
        }
        // see if it is an integer
        if (ptg.indexOf(".") == -1) {
            try {
                val i = Integer.valueOf(ptg)
                return if (i.toInt() >= 0 && i.toInt() <= 65535)
                // PtgInts are UNSIGNED + <=65535
                    "PtgInt"
                else
                    "PtgNumber"
            } catch (e: NumberFormatException) {
            }

        }
        if (ptg.indexOf("%") == ptg.length - 1) { // see if it's a percentage
            try {
                val d = Double(ptg.substring(0, ptg.indexOf("%")))
                return "PtgNumber"
            } catch (e: NumberFormatException) {
            }

        }
        // see if it is a Number
        try {
            val d = Double(ptg)
            return "PtgNumber"
        } catch (e: NumberFormatException) {
        }


        // at this point it is probably some sort of ptgref
        if (ptg.indexOf(":") != -1 || ptg.indexOf(',') != -1 || ptg.indexOf("!") != -1) {
            // ptgarea or ptgarea3d or ptgmemfunc
            return "PtgArea" // in ParseFinalLevel, PtgRef.createPtgRefFromString will handle all types of string refs
        }


        // maybe it is a garbage string, or a reference to a name (unsupported right now....)
        // check if the last character is a number. If not, it sure isn't a reference, no?
        // NO!  Can have named ranges with numbers at the end -- better to try to parse it
        try {
            if (bk!!.getName(ptg) == null) {// it's not a named range
                ExcelTools.getRowColFromString(ptg)    // if passes it's a PtgRef
                return "PtgRef"
            } else
                return "PtgName"
        } catch (e: IllegalArgumentException) {
            return "PtgName"
        }

    }

    /*
     * helper method that cleans out unneccesary parts of the formula string.
     */
    private fun cleanString(dirtystring: String): String {
        var cleanstring = StringTool.allTrim(dirtystring)
        cleanstring = StringTool.strip(cleanstring, "(")
        cleanstring = StringTool.strip(cleanstring, ",")
        return cleanstring
    }

    internal fun getPtgsFromFormulaString(fmla: String): Stack<*>? {
        return null

    }

    /**
     * parse a formula in string form and create a formula record from it
     * caluclate the new formula based on boolean setting calculate
     *
     * @param form      Formula rec
     * @param fmla      String formula either =EXPRESSION or {=EXPRESSION} for array formulas
     * @param calculate boolean truth of "calculate formula after setting"
     * @return Formula rec
     */
    fun setFormula(form: Formula, fmla: String, rc: IntArray): Formula {
        var rc = rc
        if (fmla[0] != '{') {
            try {
                val newptgs = FormulaParser.getPtgsFromFormulaString(form, fmla)
                FormulaParser.adjustParameterIds(newptgs)     // 20100614 KSC: adjust function parameter id's, if necessary, for Value, Array or Reference type
                form.expression = newptgs
            } catch (e: FunctionNotSupportedException) {  // 200902 KSC: still add record if function is not found (using N/A in place of said function)
                Logger.logErr("Adding new Formula at " + form.sheet + "!" + ExcelTools.formatLocation(rc) + " failed: " + e.toString() + ".")
                val newptgs = Stack()
                newptgs.push(PtgErr(PtgErr.ERROR_NA))
                form.expression = newptgs
            }

        } else { // Handle Array Formulas
            val pe = PtgExp()
            pe.parentRec = form
            // rowcol reference is from PARENT PtgExp not (necessarily) this formula's cell address
            // [BugTracker 2683 + OOXML Array Formulas]
            val o = form.sheet!!.getArrayFormulaParent(rc)
            if (o != null)
            // there is a parent array formula; use it's rowcol
                rc = o as IntArray
            else {            // no parent yet- add
                val addr = ExcelTools.formatLocation(rc)
                form.sheet!!.addParentArrayRef(addr, addr)
            }
            pe.init(rc[0], rc[1])
            val e = Stack()
            e.push(pe)
            FormulaParser.adjustParameterIds(e)     // adjust function parameter id's, if necessary, for Value, Array or Reference type
            form.expression = e    // add PtgExp to Formula Stack
            val a = Array()    // Create new Array Record
            a.setSheet(form.sheet)
            a.workBook = form.workBook
            a.init(fmla, rc[0], rc[1])    // init Array record from Formula String
            form.addInternalRecord(a)    // link array record to parent formula
        }

        /* is this calc necessary?
    	Object val = null;
    	try{
    	    if (form.getWorkBook().getCalcMode() == WorkBook.CALCULATE_ALWAYS)
    	    	val = form.calculateFormula();
    	}catch (Exception e){
      		Logger.logWarn("Unsupported Function: " + e + ".  OpenXLS calculation will be unavailable: " + fmla);	//20081118 KSC: display a little more info ...
    	}
        if(DEBUGLEVEL > 0)Logger.logInfo("FormulaParser.setFormula() string:" +fmla + " value: " + val);
        */

        return form
    }


    fun getFormulaString(form: Formula): String {
        val expression = form.expression
        return FormulaParser.getExpressionString(expression!!)
    }

    fun getExpressionString(expression: Stack<*>): String {
        var retval = StringBuffer()
        val sz = expression.size
        var stck = arrayOfNulls<Ptg>(sz)
        stck = expression.toTypedArray() as Array<Ptg>
        val newstck = Stack()
        for (t in 0 until sz) { // flip the stack
            newstck.add(0, stck[t])
        }
        val vals = Stack()
        while (!newstck.isEmpty()) {
            handlePtg(newstck, vals)
        }
        var s = ""
        while (!vals.isEmpty()) {
            val topP = vals.pop() as Ptg
            s = topP.textString + s
        }
        retval = StringBuffer("=$s")
        return retval.toString()
    }

    /**
     * adjustParameterIds pre-processes the function expression stack,
     * analyzing function parameters to ensure that function operands
     * contain the correct PtgID (which indicates Value, Reference or Array)
     * <br></br>
     * Reference-type ptg's contain an ID which indicates the type
     * required by the calling function: Value, Reference or Array.
     * <br></br>
     * When formula strings are parsed into formula expressions,
     * ptg reference-type parameters are assigned a default id,
     * but this id may not be correct for all functions.
     */
    fun adjustParameterIds(expression: Stack<*>) {
        val retval = StringBuffer()
        val sz = expression.size
        var stck = arrayOfNulls<Ptg>(sz)
        stck = expression.toTypedArray() as Array<Ptg>
        val newstck = Stack()
        for (t in 0 until sz) { // flip the stack
            newstck.add(0, stck[t])
        }
        val params = Stack()
        // we only care about PtgFuncVar and PtgFunc's but need to process the expression
        // stack thoroughly to get the correct parameters
        // Process function stack, gathering parameters.  When we have all the parameters
        // for a PtgFunc or a PtgFuncVar, adjust any PtgRef types, if necessary
        while (!newstck.isEmpty()) {
            val p = newstck.pop() as Ptg
            val x = 0
            var t = 0// cargs = p.getNumParams();
            if (p.isControl) {
                // do the parens thing here...
                if (p.opcode.toInt() == 0x15) { // its a parens... and there is a val
                    if (t > 0) {
                        // 20060128 - KSC: handle parens
                        val vx = arrayOfNulls<Ptg>(1)    // parens are unary ops so only 1 var allowed
                        vx[0] = params.pop() as Ptg
                        p.setVars(vx)
                        params.push(p)    // put paren (with var) back on stack
                    } else { // this paren wraps other parens...
                        params.push(p)
                    }
                }
            } else if (p.isOperator || p.isFunction) {
                if (p.isBinaryOperator) t = 2
                if (p.isUnaryOperator) t = 1
                if (p.isStandAloneOperator) t = 0
                if (p.opcode.toInt() == 0x22 || p.opcode.toInt() == 0x42 || p.opcode.toInt() == 0x62) {
                    t = p.numParams
                }// it's a ptgfuncvar
                if (p.opcode.toInt() == 0x21 || p.opcode.toInt() == 0x41 || p.opcode.toInt() == 0x61) {
                    t = p.numParams
                }// it's a ptgFunc

                if (t > params.size) {
                    t = params.size
                }
                val vx = arrayOfNulls<Ptg>(t)
                while (t > 0)
                    vx[--t] = params.pop() as Ptg// get'em
                p.setVars(vx)  // set'em
                // here is where we adjust the ptg's of the func or funcvar parameters
                if (p.opcode.toInt() == 0x22 || p.opcode.toInt() == 0x42 || p.opcode.toInt() == 0x62)
                /* it's a ptgfuncvar*/
                    (p as PtgFuncVar).adjustParameterIds()
                else if (p.opcode.toInt() == 0x21 || p.opcode.toInt() == 0x41 || p.opcode.toInt() == 0x61)
                /* it's a ptgFunc */
                    (p as PtgFunc).adjustParameterIds()
                params.push(p)// push it back on the stack
            } else if (p.isOperand) {
                params.push(p)
            }
        }
    }

    /**
     * set up the Formula chain
     *
     *
     * A5
     * 10
     *
     *
     * 3
     * EXP(
     * +
     * SUM(
     * =SUM(A5*10+EXP(3))
     * =SUM(EXP(A5*103))
     *
     *
     * A3      add to vals
     * E5      add to vals
     * +       pop last2 vals add to vals
     * (       check last -- if oper add oper, else add last, add to vals
     *
     *
     * A1:A2   at this point we should see: (A3+E5) push to vals
     * SUM(    check last -- if oper add oper, else add last, add to vals
     * +       pop last2 add to vals
     * SUM(    check last -- if oper add oper, else add last, add to vals
     * =SUM((A3+E5)+SUM(A1:A2))
     * =SUM((A3+E5)+SUM(A1:A2))
     *
     *
     * A1      add to vals
     * A2      add to vals
     * +       pop last2 vals add to vals?
     * (       check last -- if oper add oper, else add last, add to vals
     * A3      add to vals
     * A4      add to vals
     * +       check last -- if paren pop last2 vals add to vals?
     * (       check last -- if oper add oper, else add last, add to vals
     * /       pop last2 vals add to vals
     *
     *
     * =(A1+A2)/(A3+A4)
     * =(A1+A2)/(A3+A4)
     *
     *
     *
     *
     * PtgFunc taking 3 vals again can only have 1
     * We know IF has 3 vals
     * Do we need logic which knows how many vars a Ptg takes? (we sure do... -nick) Might help.
     *
     *
     * ---> WRITE CODE TO SWITCH ON NUMBER OF PARAMS.  Should Fix.
     *
     *
     *
     *
     * =IF(SUM(CONCATENATE(SUM((EXP(C2,D4,4))*2SUM(A5,C7,A2)A1:A5))=1),SUM(3),SUM(22))
     * =IF(CONCATENATE(C2,(D4+EXP(4))*2,SUM(A5,C7,A2),SUM(A1:A5))=1,3,22)
     */
    internal fun handlePtg(newstck: Stack<*>, vals: Stack<*>) {
        val p = newstck.pop() as Ptg
        val x = 0
        var t = 0// cargs = p.getNumParams();
        if (p.isOperator || p.isControl || p.isFunction) {
            t = vals.size   //this is faulty logic.  We don't care what is there, the operator should tell us.
            // do the parens thing here...
            if (p.isControl /* !vals.isEmpty()*/) {
                if (p.opcode.toInt() == 0x15) { // its a parens... and there is a val
                    if (t > 0) {
                        // 20060128 - KSC: handle parens
                        val vx = arrayOfNulls<Ptg>(1)    // parens are unary ops so only 1 var allowed
                        vx[0] = vals.pop() as Ptg
                        p.setVars(vx)
                        vals.push(p)    // put paren (with var) back on stack
                    } else { // this paren wraps other parens...
                        vals.push(p)
                    }
                    return
                }
            }
            if (t > 0) {
                // make sure we have the correct amount popped back in..
                if (p.isBinaryOperator) t = 2
                if (p.isUnaryOperator) t = 1
                if (p.isStandAloneOperator) t = 0
                if (p.opcode.toInt() == 0x22 || p.opcode.toInt() == 0x42 || p.opcode.toInt() == 0x62) {
                    t = p.numParams
                }// it's a ptgfuncvar!
                if (p.opcode.toInt() == 0x21 || p.opcode.toInt() == 0x41 || p.opcode.toInt() == 0x61) {
                    t = p.numParams
                }// it's a ptgFunc

                if (t > vals.size) {

                    // is this a real error? throw an exception?
                    if (DEBUGLEVEL > 0)
                        Logger.logWarn("FormulaParser.handlePtg: number of parameters " + t + " is greater than available " + vals.size)
                    t = vals.size
                }
                val vx = arrayOfNulls<Ptg>(t)
                while (t > 0)
                    vx[--t] = vals.pop() as Ptg// get'em
                p.setVars(vx)  // set'em
            }
            vals.push(p)// push it back on the stack
        } else if (p.isOperand) {
            vals.push(p)
        } else if (p is PtgAtr) {
            // this is probably just a space at this point, don't output error message
        } else {
            if (DEBUGLEVEL > -1)
                Logger.logInfo("FormulaParser Error - Ptg Type: " + p.opcode + " " + p.string)
        }
    }

    /**
     * create a new formula record at row column rc using formula string formStr
     *
     * @param formStr String
     * @param st
     * @param rc      int[]
     * @throws Exception
     * @return new Formula record
     */
    @Throws(Exception::class)
    fun getFormulaFromString(formStr: String, st: Boundsheet?, rc: IntArray): Formula {
        var f = Formula()
        if (st != null) {
            f.setSheet(st)
            f.workBook = st.workBook
        }
        f.data = ByteArray(6)    // necessary for setRowCol
        f.setRowCol(rc)        // do before calculateFormula as array formulas use rowcol 20090817 KSC: [BugTracker 2683 + OOXML Array Formulas]
        f = FormulaParser.setFormula(f, formStr, rc)

        return f
    }

    /**
     * create a new formula record at row column rc using formula string formStr
     *
     * @param formStr String
     * @param rc      int[]
     * @throws Exception
     * @return new Formula record
     */
    @Throws(/* 20070212 KSC: FunctionNotSupported*/Exception::class)
    fun setFormulaString(formStr: String, rc: IntArray): Formula {
        var f = Formula()
        f = FormulaParser.setFormula(f, formStr, rc)
        return f
    }


    /**
     * returns true of string s is in the form of a basic reference e.g. A1
     *
     * @param s
     * @return
     */
    fun isRef(s: String?): Boolean {
        if (s == null) return false
        val simpleOne = "(([ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+){1})"
        return s.matches(simpleOne.toRegex())
    }

    /**
     * returns true if the stirng in question is in the form of a range
     *
     * @param s
     * @return
     */
    fun isRange(s: String?): Boolean {
        if (s == null) return false
        val one = "(([(]*[ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+[)]*){1})"
        val aRange = "$one(:$one)?"
        val rangeop = "([ ]*[: ,][ ]*)"
        val rangeMatchString = "$aRange$rangeop$aRange($rangeop$aRange)*"
        val simpleOne = "(([ ]*[']?([a-zA-Z0-9 ]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+){1})"
        val simpleRangeMatchString = "($simpleOne[ ]*[:][ ]*$simpleOne)"
        return s.matches(rangeMatchString.toRegex())
    }

    /**
     * returns true if the string represents a complex range (i.e. one containing multiple range values separated by one or more of: , : or space
     *
     * @param s
     * @return
     */
    fun isComplexRange(s: String): Boolean {
        val one = "(([(]*[ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+[)]*){1})"
        val aRange = "$one(:$one)?"
        val rangeop = "([ ]*[: ,][ ]*)"
        val rangeMatchString = "$aRange$rangeop$aRange($rangeop$aRange)*"
        val simpleOne = "(([ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+){1})"
        val simpleRangeMatchString = "($simpleOne[ ]*[:][ ]*$simpleOne)"
        return isRange(s) && !s.matches(simpleRangeMatchString.toRegex())
    }

}