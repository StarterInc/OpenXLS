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

/**
 * Function Constants for all Formula Types (PtgFunc, PtgFuncVar - regular and PtgFuncVar - AddIns)
 * Modifications:
 * all xlfXXX constants were originally in FunctionHandler
 * getFunctionString was orignally in PtgFuncVar
 * FUNCTION_STRINGS were originally in FunctionHandler
 * getNumVars was originally in PtgFunc
 *
 * @see
 */
object FunctionConstants {
    /* HOW TO USE:
     *
     * 1- If implementing a formula, MAKE SURE to put it in recArr AND getFunctionString list.  If it is a ptgFunc, also input
     *     the number of args in getNumArgs.  If it's an Add-in, add xlfXX constant to the end of the Excel function numbers list.
     *
     *     All function ID's/number MUST exist in the xlfXXX constants list
     *
     *
     */
    var FTYPE_PTGFUNC = 0
    var FTYPE_PTGFUNCVAR = 1
    var FTYPE_PTGFUNCVAR_ADDIN = 2

    /****************************************
     * *
     * Excel function numbers              *
     * *
     */

    val XLF_COUNT = 0
    val XLF_IS = 1
    val XLF_IS_NA = 2
    val XLF_IS_ERROR = 3
    val XLF_SUM = 4
    val XLF_AVERAGE = 5
    val XLF_MIN = 6
    val XLF_MAX = 7
    val XLF_ROW = 8
    val xlfColumn = 9
    val xlfNa = 10
    val xlfNpv = 11
    val xlfStdev = 12
    val xlfDollar = 13
    val xlfFixed = 14
    val xlfSin = 15
    val xlfCos = 16
    val xlfTan = 17
    val xlfAtan = 18
    val xlfPi = 19
    val xlfSqrt = 20
    val xlfExp = 21
    val xlfLn = 22
    val xlfLog10 = 23
    val xlfAbs = 24
    val xlfInt = 25
    val xlfSign = 26
    val xlfRound = 27
    val xlfLookup = 28
    val xlfIndex = 29
    val xlfRept = 30
    val xlfMid = 31
    val xlfLen = 32
    val xlfValue = 33
    val xlfTrue = 34
    val xlfFalse = 35
    val xlfAnd = 36
    val xlfOr = 37
    val xlfNot = 38
    val xlfMod = 39
    val xlfDcount = 40
    val xlfDsum = 41
    val xlfDaverage = 42
    val xlfDmin = 43
    val xlfDmax = 44
    val xlfDstdev = 45
    val xlfVar = 46
    val xlfDvar = 47
    val xlfText = 48
    val xlfLinest = 49
    val xlfTrend = 50
    val xlfLogest = 51
    val xlfGrowth = 52
    val xlfGoto = 53
    val xlfHalt = 54
    val xlfPv = 56
    val xlfFv = 57
    val xlfNper = 58
    val xlfPmt = 59
    val xlfRate = 60
    val xlfMirr = 61
    val xlfIrr = 62
    val xlfRand = 63
    val xlfMatch = 64
    val xlfDate = 65
    val xlfTime = 66
    val xlfDay = 67
    val xlfMonth = 68
    val xlfYear = 69
    val xlfWeekday = 70
    val xlfHour = 71
    val xlfMinute = 72
    val xlfSecond = 73
    val xlfNow = 74
    val xlfAreas = 75
    val xlfRows = 76
    val xlfColumns = 77
    val xlfOffset = 78
    val xlfAbsref = 79
    val xlfRelref = 80
    val xlfArgument = 81
    val xlfSearch = 82
    val xlfTranspose = 83
    val xlfError = 84
    val xlfStep = 85
    val xlfType = 86
    val xlfEcho = 87
    val xlfSetName = 88
    val xlfCaller = 89
    val xlfDeref = 90
    val xlfWindows = 91
    val xlfSeries = 92
    val xlfDocuments = 93
    val xlfActiveCell = 94
    val xlfSelection = 95
    val xlfResult = 96
    val xlfAtan2 = 97
    val xlfAsin = 98
    val xlfAcos = 99
    val xlfChoose = 100
    val xlfHlookup = 101
    val xlfVlookup = 102
    val xlfLinks = 103
    val xlfInput = 104
    val xlfIsref = 105
    val xlfGetFormula = 106
    val xlfGetName = 107
    val xlfSetValue = 108
    val xlfLog = 109
    val xlfExec = 110
    val xlfChar = 111
    val xlfLower = 112
    val xlfUpper = 113
    val xlfProper = 114
    val xlfLeft = 115
    val xlfRight = 116
    val xlfExact = 117
    val xlfTrim = 118
    val xlfReplace = 119
    val xlfSubstitute = 120
    val xlfCode = 121
    val xlfNames = 122
    val xlfDirectory = 123
    val xlfFind = 124
    val xlfCell = 125
    val xlfIserr = 126
    val xlfIstext = 127
    val xlfIsnumber = 128
    val xlfIsblank = 129
    val xlfT = 130
    val xlfN = 131
    val xlfFopen = 132
    val xlfFclose = 133
    val xlfFsize = 134
    val xlfFreadln = 135
    val xlfFread = 136
    val xlfFwriteln = 137
    val xlfFwrite = 138
    val xlfFpos = 139
    val xlfDatevalue = 140
    val xlfTimevalue = 141
    val xlfSln = 142
    val xlfSyd = 143
    val xlfDdb = 144
    val xlfGetDef = 145
    val xlfReftext = 146
    val xlfTextref = 147
    val XLF_INDIRECT = 148
    val xlfRegister = 149
    val xlfCall = 150
    val xlfAddBar = 151
    val xlfAddMenu = 152
    val xlfAddCommand = 153
    val xlfEnableCommand = 154
    val xlfCheckCommand = 155
    val xlfRenameCommand = 156
    val xlfShowBar = 157
    val xlfDeleteMenu = 158
    val xlfDeleteCommand = 159
    val xlfGetChartItem = 160
    val xlfDialogBox = 161
    val xlfClean = 162
    val xlfMdeterm = 163
    val xlfMinverse = 164
    val xlfMmult = 165
    val xlfFiles = 166
    val xlfIpmt = 167
    val xlfPpmt = 168
    val xlfCounta = 169
    val xlfCancelKey = 170
    val xlfInitiate = 175
    val xlfRequest = 176
    val xlfPoke = 177
    val xlfExecute = 178
    val xlfTerminate = 179
    val xlfRestart = 180
    val xlfHelp = 181
    val xlfGetBar = 182
    val xlfProduct = 183
    val xlfFact = 184
    val xlfGetCell = 185
    val xlfGetWorkspace = 186
    val xlfGetWindow = 187
    val xlfGetDocument = 188
    val xlfDproduct = 189
    val xlfIsnontext = 190
    val xlfGetNote = 191
    val xlfNote = 192
    val xlfStdevp = 193
    val xlfVarp = 194
    val xlfDstdevp = 195
    val xlfDvarp = 196
    val xlfTrunc = 197
    val xlfIslogical = 198
    val xlfDcounta = 199
    val xlfDeleteBar = 200
    val xlfUnregister = 201
    val xlfUsdollar = 204
    val xlfFindb = 205
    val xlfSearchb = 206
    val xlfReplaceb = 207
    val xlfLeftb = 208
    val xlfRightb = 209
    val xlfMidb = 210
    val xlfLenb = 211
    val xlfRoundup = 212
    val xlfRounddown = 213
    val xlfAsc = 214
    val xlfDbcs = 215
    val xlfRank = 216
    val xlfAddress = 219
    val xlfDays360 = 220
    val xlfToday = 221
    val xlfVdb = 222
    val xlfMedian = 227
    val xlfSumproduct = 228
    val xlfSinh = 229
    val xlfCosh = 230
    val xlfTanh = 231
    val xlfAsinh = 232
    val xlfAcosh = 233
    val xlfAtanh = 234
    val xlfDget = 235
    val xlfCreateObject = 236
    val xlfVolatile = 237
    val xlfLastError = 238
    val xlfCustomUndo = 239
    val xlfCustomRepeat = 240
    val xlfFormulaConvert = 241
    val xlfGetLinkInfo = 242
    val xlfTextBox = 243
    val xlfInfo = 244
    val xlfGroup = 245
    val xlfGetObject = 246
    val xlfDb = 247
    val xlfPause = 248
    val xlfResume = 251
    val xlfFrequency = 252
    val xlfAddToolbar = 253
    val xlfDeleteToolbar = 254
    val xlfADDIN = 255    // KSC: Added; Excel function ID for add-ins
    val xlfResetToolbar = 256
    val xlfEvaluate = 257
    val xlfGetToolbar = 258
    val xlfGetTool = 259
    val xlfSpellingCheck = 260
    val xlfErrorType = 261
    val xlfAppTitle = 262
    val xlfWindowTitle = 263
    val xlfSaveToolbar = 264
    val xlfEnableTool = 265
    val xlfPressTool = 266
    val xlfRegisterId = 267
    val xlfGetWorkbook = 268
    val xlfAvedev = 269
    val xlfBetadist = 270
    val xlfGammaln = 271
    val xlfBetainv = 272
    val xlfBinomdist = 273
    val xlfChidist = 274
    val xlfChiinv = 275
    val xlfCombin = 276
    val xlfConfidence = 277
    val xlfCritbinom = 278
    val xlfEven = 279
    val xlfExpondist = 280
    val xlfFdist = 281
    val xlfFinv = 282
    val xlfFisher = 283
    val xlfFisherinv = 284
    val xlfFloor = 285
    val xlfGammadist = 286
    val xlfGammainv = 287
    val xlfCeiling = 288
    val xlfHypgeomdist = 289
    val xlfLognormdist = 290
    val xlfLoginv = 291
    val xlfNegbinomdist = 292
    val xlfNormdist = 293
    val xlfNormsdist = 294
    val xlfNorminv = 295
    val xlfNormsinv = 296
    val xlfStandardize = 297
    val xlfOdd = 298
    val xlfPermut = 299
    val xlfPoisson = 300
    val xlfTdist = 301
    val xlfWeibull = 302
    val xlfSumxmy2 = 303
    val xlfSumx2my2 = 304
    val xlfSumx2py2 = 305
    val xlfChitest = 306
    val xlfCorrel = 307
    val xlfCovar = 308
    val xlfForecast = 309
    val xlfFtest = 310
    val xlfIntercept = 311
    val xlfPearson = 312
    val xlfRsq = 313
    val xlfSteyx = 314
    val xlfSlope = 315
    val xlfTtest = 316
    val xlfProb = 317
    val xlfDevsq = 318
    val xlfGeomean = 319
    val xlfHarmean = 320
    val xlfSumsq = 321
    val xlfKurt = 322
    val xlfSkew = 323
    val xlfZtest = 324
    val xlfLarge = 325
    val xlfSmall = 326
    val xlfQuartile = 327
    val xlfPercentile = 328
    val xlfPercentrank = 329
    val xlfMode = 330
    val xlfTrimmean = 331
    val xlfTinv = 332
    val xlfMovieCommand = 334
    val xlfGetMovie = 335
    val xlfConcatenate = 336
    val xlfPower = 337
    val xlfPivotAddData = 338
    val xlfGetPivotTable = 339
    val xlfGetPivotField = 340
    val xlfGetPivotItem = 341
    val xlfRadians = 342
    val xlfDegrees = 343
    val xlfSubtotal = 344
    val XLF_SUM_IF = 345
    val xlfCountif = 346
    val xlfCountblank = 347
    val xlfScenarioGet = 348
    val xlfOptionsListsGet = 349
    val xlfIspmt = 350
    val xlfDatedif = 351
    val xlfDatestring = 352
    val xlfNumberstring = 353
    val xlfRoman = 354
    val xlfOpenDialog = 355
    val xlfSaveDialog = 356
    val xlfViewGet = 357
    val xlfGetPivotData = 358
    val xlfHyperlink = 359
    val xlfPhonetic = 360
    val xlfAverageA = 361
    val xlfMaxA = 362
    val xlfMinA = 363
    val xlfStDevPA = 364
    val xlfVarPA = 365
    val xlfStDevA = 366
    val xlfVarA = 367
    // KSC: ADD-IN formulas - use any index; name must be present in FunctionConstants.addIns
    // Financial Formulas
    val xlfAccrintm = 368
    val xlfAccrint = 369
    val xlfCoupDayBS = 370
    val xlfCoupDays = 371
    val xlfCumIPmt = 372
    val xlfCumPrinc = 373
    val xlfCoupNCD = 374
    val xlfCoupDaysNC = 375
    val xlfCoupPCD = 376
    val xlfCoupNUM = 377
    val xlfDollarDE = 378
    val xlfDollarFR = 379
    val xlfEffect = 380
    val xlfINTRATE = 381
    val xlfXIRR = 382
    val xlfXNPV = 383
    val xlfYIELD = 384
    val xlfPRICE = 385
    val xlfPRICEDISC = 386
    val xlfPRICEMAT = 387
    val xlfDURATION = 388
    val xlfMDURATION = 389
    val xlfTBillEq = 390
    val xlfTBillPrice = 391
    val xlfTBillYield = 392
    val xlfYieldDisc = 393
    val xlfYieldMat = 394
    val xlfFVSchedule = 395
    val xlfAmorlinc = 396
    val xlfAmordegrc = 397
    val xlfOddFPrice = 398
    val xlfOddLPrice = 399
    val xlfOddFYield = 400
    val xlfOddLYield = 401
    val xlfNOMINAL = 402
    val xlfDISC = 403
    val xlfRECEIVED = 404
    // Engineering Formulas
    val xlfBIN2DEC = 405
    val xlfBIN2HEX = 406
    val xlfBIN2OCT = 407
    val xlfDEC2BIN = 408
    val xlfDEC2HEX = 409
    val xlfDEC2OCT = 410
    val xlfHEX2BIN = 411
    val xlfHEX2DEC = 412
    val xlfHEX2OCT = 413
    val xlfOCT2BIN = 414
    val xlfOCT2DEC = 415
    val xlfOCT2HEX = 416
    val xlfCOMPLEX = 417
    val xlfGESTEP = 418
    val xlfDELTA = 419
    val xlfIMAGINARY = 420
    val xlfIMABS = 421
    val xlfIMDIV = 422
    val xlfIMCONJUGATE = 423
    val xlfIMCOS = 424
    val xlfIMSIN = 425
    val xlfIMREAL = 426
    val xlfIMEXP = 427
    val xlfIMSUB = 428
    val xlfIMSUM = 429
    val xlfIMPRODUCT = 430
    val xlfIMLN = 431
    val xlfIMLOG10 = 432
    val xlfIMLOG2 = 433
    val xlfIMPOWER = 434
    val xlfIMSQRT = 435
    val xlfIMARGUMENT = 436
    val xlfCONVERT = 437
    val xlfERF = 460
    val xlfERFC = 461
    // Math Add-in Formulas
    val xlfDOUBLEFACT = 438
    val xlfGCD = 439
    val xlfLCM = 440
    val xlfMROUND = 441
    val xlfMULTINOMIAL = 442
    val xlfQUOTIENT = 443
    val xlfRANDBETWEEN = 444
    val xlfSERIESSUM = 445
    val xlfSQRTPI = 446
    val xlfSUMIFS = 456
    // Information Add-in Formulas
    val xlfISEVEN = 447
    val xlfISODD = 448
    // Date/Time Add-in Formulas
    val xlfNETWORKDAYS = 449
    val xlfEDATE = 450
    val xlfEOMONTH = 451
    val xlfWEEKNUM = 452
    val xlfWORKDAY = 453
    val xlfYEARFRAC = 459
    // Statistical
    val xlfAVERAGEIF = 454
    val xlfAVERAGEIFS = 457
    val xlfCOUNTIFS = 458
    // Logical
    val xlfIFERROR = 455
    val MAXXLF = 462


    /**
     * Japanese Excel contains some different values and string output than US English Excel.
     *
     *
     * This recArr is checked if locale = japan... if null value is returned then the main list is checked
     */
    var jRecArr = arrayOf(arrayOf("YEN", xlfDollar.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("DOLLAR", xlfUsdollar.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("JIS", xlfDbcs.toString(), FTYPE_PTGFUNC.toString()))

    /**
     * Unimplemented records.  This exists to allow writing of functions that are unsupported for calculation
     */
    var unimplRecArr = arrayOf(arrayOf("ASC", xlfAsc.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DBCS", xlfDbcs.toString(), FTYPE_PTGFUNC.toString()), arrayOf("MDETERM", xlfMdeterm.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SEARCHB", xlfSearchb.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("TRANSPOSE", xlfTranspose.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("BETAINV", xlfBetainv.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("BETADIST", xlfBetadist.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("TIMEVALUE", xlfTimevalue.toString(), FTYPE_PTGFUNC.toString()), arrayOf("MINVERSE", xlfMinverse.toString(), FTYPE_PTGFUNC.toString()), arrayOf("MDETERM", xlfMdeterm.toString(), FTYPE_PTGFUNC.toString()), arrayOf("GETPIVOTDATA", xlfGetPivotData.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("HYPERLINK", xlfHyperlink.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("PHONETIC", xlfPhonetic.toString(), FTYPE_PTGFUNC.toString()), arrayOf("PERCENTILE", xlfPercentile.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("TRUNC", xlfTrunc.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("PERCENTRANK", xlfPercentrank.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("RIGHTB", xlfRightb.toString(), FTYPE_PTGFUNC.toString()), arrayOf("REPLACEB", xlfReplaceb.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FINDB", xlfFindb.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MIDB", xlfMidb.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ROWS", xlfRows.toString(), FTYPE_PTGFUNC.toString()), arrayOf("COLUMNS", xlfColumns.toString(), FTYPE_PTGFUNC.toString()), arrayOf("OFFSET", xlfOffset.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("ISTEXT", xlfIstext.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LOOKUP", xlfLookup.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("EXPONDIST", xlfExpondist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FDIST", xlfFdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FINV", xlfFinv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FTEST", xlfFtest.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FISHER", xlfFisher.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FISHERINV", xlfFisherinv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("STANDARDIZE", xlfStandardize.toString(), FTYPE_PTGFUNC.toString()), arrayOf("PERMUT", xlfPermut.toString(), FTYPE_PTGFUNC.toString()), arrayOf("POISSON", xlfPoisson.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SUMXMY2", xlfSumxmy2.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SUMX2MY2", xlfSumx2my2.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SUMX2PY2", xlfSumx2py2.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ERFC", xlfERFC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("CONFIDENCE", xlfConfidence.toString(), FTYPE_PTGFUNC.toString()), arrayOf("CRITBINOM", xlfCritbinom.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DEVSQ", xlfDevsq.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SERIESSUM", xlfSERIESSUM.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("SUBTOTAL", xlfSubtotal.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SUMSQ", xlfSumsq.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("CHIDIST", xlfChidist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("CHIINV", xlfChiinv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("CHITEST", xlfChitest.toString(), FTYPE_PTGFUNC.toString()), arrayOf("GAMMADIST", xlfGammadist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("GAMMAINV", xlfGammainv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("GAMMALN", xlfGammaln.toString(), FTYPE_PTGFUNC.toString()), arrayOf("GEOMEAN", xlfGeomean.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("GROWTH", xlfGrowth.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("HARMEAN", xlfHarmean.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("HYPGEOMDIST", xlfHypgeomdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("KURT", xlfKurt.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("LOGEST", xlfLogest.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("LOGINV", xlfLoginv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LOGNORMDIST", xlfLognormdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NEGBINOMDIST", xlfNegbinomdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("PROB", xlfProb.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SKEW", xlfSkew.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("STDEVPA", xlfStDevPA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("STDEVP", xlfStdevp.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("STDEVA", xlfStDevA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("TDIST", xlfTdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("TINV", xlfTinv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("TTEST", xlfTtest.toString(), FTYPE_PTGFUNC.toString()), arrayOf("VARA", xlfVarA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("VARPA", xlfVarPA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("WEIBULL", xlfWeibull.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ZTEST", xlfZtest.toString(), FTYPE_PTGFUNCVAR.toString()))

    // Contains function name, id and type of ALL Formulas (PtgFuncs, PtgFuncVars and Add-in PtgFuncVars)

    // fetch the pattern match from: http://office.microsoft.com/client/helpcategory.aspx?CategoryID=CH100645029990&lcid=1033&NS=EXCEL&Version=12&CTT=4
    var recArr = arrayOf(arrayOf("Pi", xlfPi.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Round", xlfRound.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Rept", xlfRept.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Mid", xlfMid.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Mod", xlfMod.toString(), FTYPE_PTGFUNC.toString()), arrayOf("MMult", xlfMmult.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Rand", xlfRand.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Date", xlfDate.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Time", xlfTime.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Day", xlfDay.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Now", xlfNow.toString(), FTYPE_PTGFUNC.toString()), arrayOf("TAN", xlfTan.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Atan2", xlfAtan2.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Replace", xlfReplace.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Exact", xlfExact.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Trim", xlfTrim.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Text", xlfText.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Roundup", xlfRoundup.toString(), FTYPE_PTGFUNC.toString()), arrayOf("RoundDown", xlfRounddown.toString(), FTYPE_PTGFUNC.toString()), arrayOf("today", xlfToday.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Combin", xlfCombin.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Floor", xlfFloor.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Ceiling", xlfCeiling.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Power", xlfPower.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Hour", xlfHour.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Minute", xlfMinute.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Month", xlfMonth.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Year", xlfYear.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Second", xlfSecond.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Quartile", xlfQuartile.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Frequency", xlfFrequency.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Linest", xlfLinest.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Correl", xlfCorrel.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Slope", xlfSlope.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Intercept", xlfIntercept.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Pearson", xlfPearson.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Rsq", xlfRsq.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Steyx", xlfSteyx.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Forecast", xlfForecast.toString(), FTYPE_PTGFUNC.toString()), arrayOf("Covar", xlfCovar.toString(), FTYPE_PTGFUNC.toString()), arrayOf("IsNumber", xlfIsnumber.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DAVERAGE", xlfDaverage.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DCOUNT", xlfDcount.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DCOUNTA", xlfDcounta.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DGET", xlfDget.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DMIN", xlfDmin.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DMAX", xlfDmax.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DPRODUCT", xlfDproduct.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DSTDEVP", xlfDstdevp.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DSTDEV", xlfDstdev.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DSUM", xlfDsum.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DVAR", xlfDvar.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DVARP", xlfDvarp.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SQRT", xlfSqrt.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NA", xlfNa.toString(), FTYPE_PTGFUNC.toString()), arrayOf("EXP", xlfExp.toString(), FTYPE_PTGFUNC.toString()), arrayOf("MIRR", xlfMirr.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SLN", xlfSln.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SYD", xlfSyd.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ISPMT", xlfIspmt.toString(), FTYPE_PTGFUNC.toString()), arrayOf("UPPER", xlfUpper.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LOWER", xlfLower.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LEN", xlfLen.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ISLOGICAL", xlfIslogical.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ISERROR", XLF_IS_ERROR.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ISNONTEXT", xlfIsnontext.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ISBLANK", xlfIsblank.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ISREF", xlfIsref.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SIN", xlfSin.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SINH", xlfSinh.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ASIN", xlfAsin.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ASINH", xlfAsinh.toString(), FTYPE_PTGFUNC.toString()), arrayOf("COS", xlfCos.toString(), FTYPE_PTGFUNC.toString()), arrayOf("COSH", xlfCosh.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ACOS", xlfAcos.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ACOSH", xlfAcosh.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ATAN", xlfAtan.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ATANH", xlfAtanh.toString(), FTYPE_PTGFUNC.toString()), arrayOf("INT", xlfInt.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ABS", xlfAbs.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NOT", xlfNot.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DEGREES", xlfDegrees.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SIGN", xlfSign.toString(), FTYPE_PTGFUNC.toString()), arrayOf("EVEN", xlfEven.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ODD", xlfOdd.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LN", xlfLn.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FACT", xlfFact.toString(), FTYPE_PTGFUNC.toString()), arrayOf("RADIANS", xlfRadians.toString(), FTYPE_PTGFUNC.toString()), arrayOf("PROPER", xlfProper.toString(), FTYPE_PTGFUNC.toString()), arrayOf("CHAR", xlfChar.toString(), FTYPE_PTGFUNC.toString()), arrayOf("ERROR.TYPE", xlfErrorType.toString(), FTYPE_PTGFUNC.toString()), arrayOf("T", xlfT.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LOG10", xlfLog10.toString(), FTYPE_PTGFUNC.toString()), arrayOf("VALUE", xlfValue.toString(), FTYPE_PTGFUNC.toString()), arrayOf("CODE", xlfCode.toString(), FTYPE_PTGFUNC.toString()), arrayOf("N", xlfN.toString(), FTYPE_PTGFUNC.toString()), arrayOf("DATEVALUE", xlfDatevalue.toString(), FTYPE_PTGFUNC.toString()), arrayOf("SMALL", xlfSmall.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LARGE", xlfLarge.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NORMDIST", xlfNormdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NORMSDIST", xlfNormsdist.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NORMSINV", xlfNormsinv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("NORMINV", xlfNorminv.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LENB", xlfLenb.toString(), FTYPE_PTGFUNC.toString()), arrayOf("INFO", xlfInfo.toString(), FTYPE_PTGFUNC.toString()), arrayOf("LEFTB", xlfLeftb.toString(), FTYPE_PTGFUNC.toString()), arrayOf("TRUE", xlfTrue.toString(), FTYPE_PTGFUNC.toString()), arrayOf("FALSE", xlfFalse.toString(), FTYPE_PTGFUNC.toString()),
            // PtgFuncVars
            arrayOf("COUNT", XLF_COUNT.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("COUNTA", xlfCounta.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("COUNTIF", xlfCountif.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("COUNTBLANK", xlfCountblank.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("IF", XLF_IS.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("ISNA", XLF_IS_NA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("ISERR", xlfIserr.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SUM", XLF_SUM.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SUMIF", XLF_SUM_IF.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("AVERAGE", XLF_AVERAGE.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MINA", xlfMinA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MIN", XLF_MIN.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MAXA", xlfMaxA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MAX", XLF_MAX.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("ROW", XLF_ROW.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("COLUMN", xlfColumn.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("NPV", xlfNpv.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("PMT", xlfPmt.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("DB", xlfDb.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("FIND", xlfFind.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("DAYS360", xlfDays360.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("LEFT", xlfLeft.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("LOG", xlfLog.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MEDIAN", xlfMedian.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MODE", xlfMode.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("RANK", xlfRank.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("RIGHT", xlfRight.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("STDEV", xlfStdev.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("VAR", xlfVar.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("VARP", xlfVarp.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("TANH", xlfTanh.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("VLOOKUP", xlfVlookup.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("HLOOKUP", xlfHlookup.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("CONCATENATE", xlfConcatenate.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("INDEX", xlfIndex.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("MATCH", xlfMatch.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("FIXED", xlfFixed.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("AND", xlfAnd.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("OR", xlfOr.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("CHOOSE", xlfChoose.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("ADDRESS", xlfAddress.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("ROMAN", xlfRoman.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("DOLLAR", xlfDollar.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("USDOLLAR", xlfUsdollar.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("AVEDEV", xlfAvedev.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SUBSTITUTE", xlfSubstitute.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("PRODUCT", xlfProduct.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SEARCH", xlfSearch.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("AVERAGEA", xlfAverageA.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("TREND", xlfTrend.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("SUMPRODUCT", xlfSumproduct.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("INDIRECT", XLF_INDIRECT.toString(), FTYPE_PTGFUNCVAR.toString()),
            // Add-in Formulas
            // Financial Formulas
            arrayOf("ACCRINTM", xlfAccrintm.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ACCRINT", xlfAccrint.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUPDAYBS", xlfCoupDayBS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUPDAYS", xlfCoupDays.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("PV", xlfPv.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("FV", xlfFv.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IPMT", xlfIpmt.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("CUMIPMT", xlfCumIPmt.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("CUMPRINC", xlfCumPrinc.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUPNCD", xlfCoupNCD.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUPDAYSNC", xlfCoupDaysNC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUPPCD", xlfCoupPCD.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUPNUM", xlfCoupNUM.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DOLLARDE", xlfDollarDE.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DOLLARFR", xlfDollarFR.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("EFFECT", xlfEffect.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("INTRATE", xlfINTRATE.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IRR", xlfIrr.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("XIRR", xlfXIRR.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("XNPV", xlfXNPV.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("RATE", xlfRate.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("YIELD", xlfYIELD.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("PRICE", xlfPRICE.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("PRICEDISC", xlfPRICEDISC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DISC", xlfDISC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("PRICEMAT", xlfPRICEMAT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DURATION", xlfDURATION.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("MDURATION", xlfMDURATION.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("NPER", xlfNper.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("TBILLEQ", xlfTBillEq.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("TBILLPRICE", xlfTBillPrice.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("TBILLYIELD", xlfTBillYield.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("YIELDDISC", xlfYieldDisc.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("YIELDMAT", xlfYieldMat.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("PPMT", xlfPpmt.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("FVSCHEDULE", xlfFVSchedule.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("AMORLINC", xlfAmorlinc.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("AMORDEGRC", xlfAmordegrc.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ODDFPRICE", xlfOddFPrice.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ODDLPRICE", xlfOddLPrice.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ODDFYIELD", xlfOddFYield.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ODDLYIELD", xlfOddLYield.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("NOMINAL", xlfNOMINAL.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("VDB", xlfVdb.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DDB", xlfDdb.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("RECEIVED", xlfRECEIVED.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()),
            // Engineering Formulas
            arrayOf("BIN2DEC", xlfBIN2DEC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("BIN2HEX", xlfBIN2HEX.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("BIN2OCT", xlfBIN2OCT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DEC2BIN", xlfDEC2BIN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DEC2HEX", xlfDEC2HEX.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DEC2OCT", xlfDEC2OCT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("HEX2BIN", xlfHEX2BIN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("HEX2DEC", xlfHEX2DEC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("HEX2OCT", xlfHEX2OCT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("OCT2BIN", xlfOCT2BIN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("OCT2DEC", xlfOCT2DEC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("OCT2HEX", xlfOCT2HEX.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COMPLEX", xlfCOMPLEX.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("GESTEP", xlfGESTEP.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("DELTA", xlfDELTA.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMAGINARY", xlfIMAGINARY.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMREAL", xlfIMREAL.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMARGUMENT", xlfIMARGUMENT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMABS", xlfIMABS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMDIV", xlfIMDIV.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMCONJUGATE", xlfIMCONJUGATE.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMCOS", xlfIMCOS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMSIN", xlfIMSIN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMEXP", xlfIMEXP.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMPOWER", xlfIMPOWER.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMSQRT", xlfIMSQRT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMSUB", xlfIMSUB.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMSUM", xlfIMSUM.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMPRODUCT", xlfIMPRODUCT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMLN", xlfIMLN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMLOG10", xlfIMLOG10.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("IMLOG2", xlfIMLOG2.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("CONVERT", xlfCONVERT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ERF", xlfERF.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()),
            // Math Add-In Formulas
            arrayOf("FACTDOUBLE", xlfDOUBLEFACT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("GCD", xlfGCD.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("LCM", xlfLCM.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("MROUND", xlfMROUND.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("MULTINOMIAL", xlfMULTINOMIAL.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("QUOTIENT", xlfQUOTIENT.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("RANDBETWEEN", xlfRANDBETWEEN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("SERIESSUM", xlfSERIESSUM.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("SQRTPI", xlfSQRTPI.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("SUMIFS", xlfSUMIFS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()),
            // Information Add-Ins
            arrayOf("ISEVEN", xlfISEVEN.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("ISODD", xlfISODD.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()),
            // Date/Time Add-in Formulas
            arrayOf("NETWORKDAYS", xlfNETWORKDAYS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("EDATE", xlfEDATE.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("EOMONTH", xlfEOMONTH.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("WEEKNUM", xlfWEEKNUM.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("WEEKDAY", xlfWeekday.toString(), FTYPE_PTGFUNCVAR.toString()), arrayOf("WORKDAY", xlfWORKDAY.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("YEARFRAC", xlfYEARFRAC.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()),
            // Statistical
            arrayOf("AVERAGEIF", xlfAVERAGEIF.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("AVERAGEIFS", xlfAVERAGEIFS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()), arrayOf("COUNTIFS", xlfCOUNTIFS.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()),
            // Logical
            arrayOf("IFERROR", xlfIFERROR.toString(), FTYPE_PTGFUNCVAR_ADDIN.toString()))

    /**
     * Handles differences
     * in japanese locale xls
     *
     * @param iftb
     * @return
     */
    fun getJFunctionString(iftb: Short): String? {
        when (iftb) {
            xlfDollar -> return "YEN("
            xlfUsdollar -> return "DOLLAR("
            xlfDbcs -> return "JIS("
        }
        return null
    }


    fun getFunctionString(iftb: Short): String {
        when (iftb) {
            xlfADDIN -> return ""
            XLF_COUNT -> return "COUNT("
            XLF_IS -> return "IF("
            XLF_IS_NA -> return "ISNA("
            XLF_IS_ERROR -> return "ISERROR("
            XLF_SUM -> return "SUM("
            XLF_AVERAGE -> return "AVERAGE("
            XLF_MIN -> return "MIN("
            XLF_MAX -> return "MAX("
            XLF_ROW -> return "ROW("
            xlfColumn -> return "COLUMN("
            xlfNa -> return "NA("
            xlfNpv -> return "NPV("
            xlfStdev -> return "STDEV("
            xlfDollar -> return "DOLLAR("
            xlfFixed -> return "FIXED("
            xlfSin -> return "SIN("
            xlfCos -> return "COS("
            xlfTan -> return "TAN("
            xlfAtan -> return "ATAN("
            xlfPi -> return "PI("
            xlfSqrt -> return "SQRT("
            xlfExp -> return "EXP("
            xlfLn -> return "LN("
            xlfLog10 -> return "LOG10("
            xlfAbs -> return "ABS("
            xlfInt -> return "INT("
            xlfSign -> return "SIGN("
            xlfRound -> return "ROUND("
            xlfLookup -> return "LOOKUP("
            xlfIndex -> return "INDEX("
            xlfRept -> return "REPT("
            xlfMid -> return "MID("
            xlfLen -> return "LEN("
            xlfValue -> return "VALUE("
            xlfTrue -> return "TRUE("
            xlfFalse -> return "FALSE("
            xlfAnd -> return "AND("
            xlfOr -> return "OR("
            xlfNot -> return "NOT("
            xlfMod -> return "MOD("
            xlfDaverage -> return "DAVERAGE("
            xlfDcount -> return "DCOUNT("
            xlfDcounta -> return "DCOUNTA("
            xlfDget -> return "DGET("
            xlfDmax -> return "DMAX("
            xlfDmin -> return "DMIN("
            xlfDproduct -> return "DPRODUCT("
            xlfDstdev -> return "DSTDEV("
            xlfDstdevp -> return "DSTDEVP("
            xlfDsum -> return "DSUM("
            xlfDvar -> return "DVAR("
            xlfDvarp -> return "DVARP("
            xlfVar -> return "VAR("
            xlfText -> return "TEXT("
            xlfLinest -> return "LINEST("
            xlfTrend -> return "TREND("
            xlfLogest -> return "LOGEST("
            xlfGrowth -> return "GROWTH("
            xlfGoto -> return "GOTO("
            xlfHalt -> return "HALT("
            xlfPv -> return "PV("
            xlfFv -> return "FV("
            xlfNper -> return "NPER("
            xlfPmt -> return "PMT("
            xlfRate -> return "RATE("
            xlfMirr -> return "MIRR("
            xlfIrr -> return "IRR("
            xlfRand -> return "RAND("
            xlfMatch -> return "MATCH("
            xlfDate -> return "DATE("
            xlfTime -> return "TIME("
            xlfDay -> return "DAY("
            xlfMonth -> return "MONTH("
            xlfYear -> return "YEAR("
            xlfWeekday -> return "WEEKDAY("
            xlfHour -> return "HOUR("
            xlfMinute -> return "MINUTE("
            xlfSecond -> return "SECOND("
            xlfNow -> return "NOW("
            xlfAreas -> return "AREAS("
            xlfRows -> return "ROWS("
            xlfColumns -> return "COLUMNS("
            xlfOffset -> return "OFFSET("
            xlfAbsref -> return "ABSREF("
            xlfRelref -> return "RELREF("
            xlfArgument -> return "ARGUMENT("
            xlfSearch -> return "SEARCH("
            xlfTranspose -> return "TRANSPOSE("
            xlfError -> return "ERROR("
            xlfStep -> return "STEP("
            xlfType -> return "TYPE("
            xlfEcho -> return "ECHO("
            xlfSetName -> return "SETNAME("
            xlfCaller -> return "CALLER("
            xlfDeref -> return "DEREF("
            xlfWindows -> return "WINDOWS("
            xlfSeries -> return "SERIES("
            xlfDocuments -> return "DOCUMENTS("
            xlfActiveCell -> return "ACTIVECELL("
            xlfSelection -> return "SELECTION("
            xlfResult -> return "RESULT("
            xlfAtan2 -> return "ATAN2("
            xlfAsin -> return "ASIN("
            xlfAcos -> return "ACOS("
            xlfChoose -> return "CHOOSE("
            xlfHlookup -> return "HLOOKUP("
            xlfVlookup -> return "VLOOKUP("
            xlfLinks -> return "LINKS("
            xlfInput -> return "INPUT("
            xlfIsref -> return "ISREF("
            xlfGetFormula -> return "GETFORMULA("
            xlfGetName -> return "GETNAME("
            xlfSetValue -> return "SETVALUE("
            xlfLog -> return "LOG("
            xlfExec -> return "EXEC("
            xlfChar -> return "CHAR("
            xlfLower -> return "LOWER("
            xlfUpper -> return "UPPER("
            xlfProper -> return "PROPER("
            xlfLeft -> return "LEFT("
            xlfRight -> return "RIGHT("
            xlfExact -> return "EXACT("
            xlfTrim -> return "TRIM("
            xlfReplace -> return "REPLACE("
            xlfSubstitute -> return "SUBSTITUTE("
            xlfCode -> return "CODE("
            xlfNames -> return "NAMES("
            xlfDirectory -> return "DIRECTORY("
            xlfFind -> return "FIND("
            xlfCell -> return "CELL("
            xlfIserr -> return "ISERR("
            xlfIstext -> return "ISTEXT("
            xlfIsnumber -> return "ISNUMBER("
            xlfIsblank -> return "ISBLANK("
            xlfT -> return "T("
            xlfN -> return "N("
            xlfFopen -> return "FOPEN("
            xlfFclose -> return "FCLOSE("
            xlfFsize -> return "SIZE("
            xlfFreadln -> return "FREADLN("
            xlfFread -> return "FREAD("
            xlfFwriteln -> return "FWRITELN("
            xlfFwrite -> return "FWRITE("
            xlfFpos -> return "FPOS("
            xlfDatevalue -> return "DATEVALUE("
            xlfTimevalue -> return "TIMEVALUE("
            xlfSln -> return "SLN("
            xlfSyd -> return "SYD("
            xlfDdb -> return "DDB("
            xlfGetDef -> return "GETDEF("
            xlfReftext -> return "REFTEXT("
            xlfTextref -> return "TEXTREF("
            XLF_INDIRECT -> return "INDIRECT("
            xlfRegister -> return "REGISTER("
            xlfCall -> return "CALL("
            xlfAddBar -> return "ADDBAR("
            xlfAddMenu -> return "ADDMENU("
            xlfAddCommand -> return "ADDCOMMAND("
            xlfEnableCommand -> return "ENABLECOMMAND("
            xlfCheckCommand -> return "CHECKCOMMAND("
            xlfRenameCommand -> return "RENAMECOMMAND("
            xlfShowBar -> return "SHOWBAR("
            xlfDeleteMenu -> return "DELETEMENU("
            xlfDeleteCommand -> return "DELETECOMMAND("
            xlfGetChartItem -> return "CHARTITEM("
            xlfDialogBox -> return "DIALOGBOX("
            xlfClean -> return "CLEAN("
            xlfMdeterm -> return "MDETERM("
            xlfMinverse -> return "MINVERSE("
            xlfMmult -> return "MMULT("
            xlfFiles -> return "FILES("
            xlfIpmt -> return "IPMT("
            xlfPpmt -> return "PPMT("
            xlfCounta -> return "COUNTA("
            xlfCancelKey -> return "CANCELKEY("
            xlfInitiate -> return "INITIATE("
            xlfRequest -> return "REQUEST("
            xlfPoke -> return "POKE("
            xlfExecute -> return "EXECUTE("
            xlfTerminate -> return "TERMINATE("
            xlfRestart -> return "RESTART("
            xlfHelp -> return "HELP("
            xlfGetBar -> return "GETBAR("
            xlfProduct -> return "PRODUCT("
            xlfFact -> return "FACT("
            xlfGetCell -> return "GETCELL("
            xlfGetWorkspace -> return "GETWORKSPACE("
            xlfGetWindow -> return "GETWINDOW("
            xlfGetDocument -> return "GETDOCUMENT("
            xlfIsnontext -> return "ISNONTEXT("
            xlfGetNote -> return "GETNOTE("
            xlfNote -> return "NOTE("
            xlfStdevp -> return "STDEVP("
            xlfVarp -> return "VARP("
            xlfTrunc -> return "TRUNC("
            xlfIslogical -> return "ISLOGICAL("
            xlfDeleteBar -> return "DELETEBAR("
            xlfUnregister -> return "UNREGISTER("
            xlfUsdollar -> return "USDOLLAR("
            xlfFindb -> return "FINDB("
            xlfSearchb -> return "SEARCHB("
            xlfReplaceb -> return "REPLACEB("
            xlfLeftb -> return "LEFTB("
            xlfRightb -> return "RIGHTB("
            xlfMidb -> return "MIDB("
            xlfLenb -> return "LENB("
            xlfRoundup -> return "ROUNDUP("
            xlfRounddown -> return "ROUNDDOWN("
            xlfAsc -> return "ASC("
            xlfDbcs -> return "DBCS("
            xlfRank -> return "RANK("
            xlfAddress -> return "ADDRESS("
            xlfDays360 -> return "DAYS360("
            xlfToday -> return "TODAY("
            xlfVdb -> return "VDB("
            xlfMedian -> return "MEDIAN("
            xlfSumproduct -> return "SUMPRODUCT("
            xlfSinh -> return "SINH("
            xlfCosh -> return "COSH("
            xlfTanh -> return "TANH("
            xlfAsinh -> return "ASINH("
            xlfAcosh -> return "ACOSH("
            xlfAtanh -> return "ATANH("
            xlfCreateObject -> return "CREATEOBJECT("
            xlfVolatile -> return "VOLATILE("
            xlfLastError -> return "LASTERROR("
            xlfCustomUndo -> return "CUSTOMUNDO("
            xlfCustomRepeat -> return "CUSTOMREPEAT("
            xlfFormulaConvert -> return "FORMULACONVERT("
            xlfGetLinkInfo -> return "GETLINKINFO("
            xlfTextBox -> return "TEXTBOX("
            xlfInfo -> return "INFO("
            xlfGroup -> return "GROUP("
            xlfGetObject -> return "GETOBJECT("
            xlfDb -> return "DB("
            xlfPause -> return "PAUSE("
            xlfResume -> return "RESUME("
            xlfFrequency -> return "FREQUENCY("
            xlfAddToolbar -> return "ADDTOOLBAR("
            xlfDeleteToolbar -> return "DELETETOOLBAR("
            xlfResetToolbar -> return "RESETTOOLBAR("
            xlfEvaluate -> return "EVALUATE("
            xlfGetToolbar -> return "GETTOOLBAR("
            xlfGetTool -> return "GETTOOL("
            xlfSpellingCheck -> return "SPELLINGCHECK("
            xlfErrorType -> return "ERROR.TYPE("
            xlfAppTitle -> return "APPTITLE("
            xlfWindowTitle -> return "WINDOWTITLE("
            xlfSaveToolbar -> return "SAVETOOLBAR("
            xlfEnableTool -> return "ENABLETOOL("
            xlfPressTool -> return "PRESSTOOL("
            xlfRegisterId -> return "REGISTERID("
            xlfGetWorkbook -> return "GETWORKBOOK("
            xlfAvedev -> return "AVEDEV("
            xlfBetadist -> return "BETADIST("
            xlfGammaln -> return "GAMMALN("
            xlfBetainv -> return "BETAINV("
            xlfBinomdist -> return "BINOMDIST("
            xlfChidist -> return "CHIDIST("
            xlfChiinv -> return "CHIINV("
            xlfCombin -> return "COMBIN("
            xlfConfidence -> return "CONFIDENCE("
            xlfCritbinom -> return "CRITBINOM("
            xlfEven -> return "EVEN("
            xlfExpondist -> return "EXPONDIST("
            xlfFdist -> return "FDIST("
            xlfFinv -> return "FINV("
            xlfFisher -> return "FISHER("
            xlfFisherinv -> return "FISHERINV("
            xlfFloor -> return "FLOOR("
            xlfGammadist -> return "GAMMADIST("
            xlfGammainv -> return "GAMMAINV("
            xlfCeiling -> return "CEILING("
            xlfHypgeomdist -> return "HYPGEOMDIST("
            xlfLognormdist -> return "LOGNORMDIST("
            xlfLoginv -> return "LOGINV("
            xlfNegbinomdist -> return "NEGBINOMDIST("
            xlfNormdist -> return "NORMDIST("
            xlfNormsdist -> return "NORMSDIST("
            xlfNorminv -> return "NORMINV("
            xlfNormsinv -> return "NORMSINV("
            xlfStandardize -> return "STANDARDIZE("
            xlfOdd -> return "ODD("
            xlfPermut -> return "PERMUT("
            xlfPoisson -> return "POISSON("
            xlfTdist -> return "TDIST("
            xlfWeibull -> return "WEIBULL("
            xlfSumxmy2 -> return "SUMXMY2("
            xlfSumx2my2 -> return "SUMX2MY2("
            xlfSumx2py2 -> return "SUMX2PY2("
            xlfChitest -> return "CHITEST("
            xlfCorrel -> return "CORREL("
            xlfCovar -> return "COVAR("
            xlfForecast -> return "FORECAST("
            xlfFtest -> return "FTEST("
            xlfIntercept -> return "INTERCEPT("
            xlfPearson -> return "PEARSON("
            xlfRsq -> return "RSQ("
            xlfSteyx -> return "STEYX("
            xlfSlope -> return "SLOPE("
            xlfTtest -> return "TTEST("
            xlfProb -> return "PROB("
            xlfDevsq -> return "DEVSQ("
            xlfGeomean -> return "GEOMEAN("
            xlfHarmean -> return "HARMEAN("
            xlfSumsq -> return "SUMSQ("
            xlfKurt -> return "KURT("
            xlfSkew -> return "SKEW("
            xlfZtest -> return "ZTEST("
            xlfLarge -> return "LARGE("
            xlfSmall -> return "SMALL("
            xlfQuartile -> return "QUARTILE("
            xlfPercentile -> return "PERCENTILE("
            xlfPercentrank -> return "PERCENTRANK("
            xlfMode -> return "MODE("
            xlfTrimmean -> return "TRIMMEAN("
            xlfTinv -> return "TINV("
            xlfMovieCommand -> return "MOVIECOMMAND("
            xlfGetMovie -> return "GETMOVIE("
            xlfConcatenate -> return "CONCATENATE("
            xlfPower -> return "POWER("
            xlfPivotAddData -> return "PIVOTADDDATA("
            xlfGetPivotTable -> return "GETPIVOTTABLE("
            xlfGetPivotField -> return "GETPIVOTFIELD("
            xlfGetPivotItem -> return "GETPIVOTITEM("
            xlfRadians -> return "RADIANS("
            xlfDegrees -> return "DEGREES("
            xlfSubtotal -> return "SUBTOTAL("
            XLF_SUM_IF -> return "SUMIF("
            xlfCountif -> return "COUNTIF("
            xlfCountblank -> return "COUNTBLANK("
            xlfScenarioGet -> return "SCENARIOGET("
            xlfOptionsListsGet -> return "OPTIONSLISTSGET("
            xlfIspmt -> return "ISPMT("
            xlfDatedif -> return "DATEDIF("
            xlfDatestring -> return "DATESTRING("
            xlfNumberstring -> return "NUMBERSTRING("
            xlfRoman -> return "ROMAN("
            xlfOpenDialog -> return "OPENDIALOG("
            xlfSaveDialog -> return "SAVEDIALOG("
            xlfViewGet -> return "VIEWGET("
            xlfGetPivotData -> return "GETPIVOTDATA("
            xlfHyperlink -> return "HYPERLINK("
            xlfPhonetic -> return "PHONETIC("
            xlfAverageA -> return "AVERAGEA("
            xlfMaxA -> return "MAXA("
            xlfMinA -> return "MINA("
            xlfStDevPA -> return "STDEVPA("
            xlfVarPA -> return "VARPA("
            xlfStDevA -> return "STDEVA("
            xlfVarA -> return "VARA("
            // ADD-IN FORMULAS
            // Financial Formulas AddIns
            xlfAccrintm -> return "ACCRINTM("
            xlfAccrint -> return "ACCRINT("
            xlfCoupDayBS -> return "COUPDAYBS("
            xlfCoupDays -> return "COUPDAYS("
            xlfCoupDaysNC -> return "COUPDAYSNC("
            xlfCumIPmt -> return "CUMIPMT("
            xlfCumPrinc -> return "CUMPRINC("
            xlfCoupNCD -> return "COUPNCD("
            xlfCoupPCD -> return "COUPPCD("
            xlfCoupNUM -> return "COUPNUM("
            xlfDollarDE -> return "DOLLARDE("
            xlfDollarFR -> return "DOLLARFR("
            xlfEffect -> return "EFFECT("
            xlfINTRATE -> return "INTRATE("
            xlfXIRR -> return "XIRR("
            xlfXNPV -> return "XNPV("
            xlfYIELD -> return "YIELD("
            xlfPRICE -> return "PRICE("
            xlfPRICEDISC -> return "PRICEDISC("
            xlfDISC -> return "DISC("
            xlfPRICEMAT -> return "PRICEMAT("
            xlfDURATION -> return "DURATION("
            xlfMDURATION -> return "MDURATION("
            xlfTBillEq -> return "TBILLEQ("
            xlfTBillPrice -> return "TBILLPRICE("
            xlfTBillYield -> return "TBILLYIELD("
            xlfYieldDisc -> return "YIELDDISC("
            xlfYieldMat -> return "YIELDMAT("
            xlfFVSchedule -> return "FVSCHEDULE("
            xlfAmorlinc -> return "AMORLINC("
            xlfAmordegrc -> return "AMORDEGRC("
            xlfOddFPrice -> return "ODDFPRICE("
            xlfOddFYield -> return "ODDFYIELD("
            xlfOddLPrice -> return "ODDLPRICE("
            xlfOddLYield -> return "ODDLYIELD("
            xlfNOMINAL -> return "NOMINAL("
            xlfRECEIVED -> return "RECEIVED("
            // Engineering Formulas AddIns
            xlfBIN2DEC -> return "BIN2DEC("
            xlfBIN2HEX -> return "BIN2HEX("
            xlfBIN2OCT -> return "BIN2OCT("
            xlfDEC2BIN -> return "DEC2BIN("
            xlfDEC2HEX -> return "DEC2HEX("
            xlfDEC2OCT -> return "DEC2OCT("
            xlfHEX2BIN -> return "HEX2BIN("
            xlfHEX2DEC -> return "HEX2DEC("
            xlfHEX2OCT -> return "HEX2OCT("
            xlfOCT2BIN -> return "OCT2BIN("
            xlfOCT2DEC -> return "OCT2DEC("
            xlfOCT2HEX -> return "OCT2HEX("
            xlfCOMPLEX -> return "COMPLEX("
            xlfGESTEP -> return "GESTEP("
            xlfDELTA -> return "DELTA("
            xlfIMAGINARY -> return "IMAGINARY("
            xlfIMREAL -> return "IMREAL("
            xlfIMARGUMENT -> return "IMARGUMENT("
            xlfIMABS -> return "IMABS("
            xlfIMDIV -> return "IMDIV("
            xlfIMCONJUGATE -> return "IMCONJUGATE("
            xlfIMCOS -> return "IMCOS("
            xlfIMSIN -> return "IMSIN("
            xlfIMEXP -> return "IMEXP("
            xlfIMPOWER -> return "IMPOWER("
            xlfIMSQRT -> return "IMSQRT("
            xlfIMSUB -> return "IMSUB("
            xlfIMSUM -> return "IMSUM("
            xlfIMPRODUCT -> return "IMPRODUCT("
            xlfIMLN -> return "IMLN("
            xlfIMLOG10 -> return "IMLOG10("
            xlfIMLOG2 -> return "IMLOG2("
            xlfCONVERT -> return "CONVERT("
            xlfDOUBLEFACT -> return "FACTDOUBLE("
            xlfGCD -> return "GCD("
            xlfLCM -> return "LCM("
            xlfMROUND -> return "MROUND("
            xlfMULTINOMIAL -> return "MULTINOMIAL("
            xlfQUOTIENT -> return "QUOTIENT("
            xlfRANDBETWEEN -> return "RANDBETWEEN("
            xlfSERIESSUM -> return "SERIESSUM("
            xlfSQRTPI -> return "SQRTPI("
            xlfERF -> return "ERF("
            // information Add-ins
            xlfISEVEN -> return "ISEVEN("
            xlfISODD -> return "ISODD("
            // Date/Time Add-in Formulas
            xlfNETWORKDAYS -> return "NETWORKDAYS("
            xlfEDATE -> return "EDATE("
            xlfEOMONTH -> return "EOMONTH("
            xlfWEEKNUM -> return "WEEKNUM("
            xlfWORKDAY -> return "WORKDAY("
            xlfYEARFRAC -> return "YEARFRAC("
            // Statistical
            xlfAVERAGEIF -> return "AVERAGEIF("
            xlfAVERAGEIFS -> return "AVERAGEIFS("
            xlfCOUNTIFS -> return "COUNTIFS("
            // Logical
            xlfIFERROR -> return "IFERROR("
            // Math
            xlfSUMIFS -> return "SUMIFS("
        }
        return ""
    }

    // Num Params for all PTGFUNCs  - ptgfuncvar's have variable # args (hence the name ...)
    fun getNumParams(iftab: Int): Int {
        if (iftab == xlfNa) return 0 // na
        if (iftab == xlfPi) return 0 // Pi
        if (iftab == xlfRound) return 2 // Round
        if (iftab == xlfRept) return 2 // rept
        if (iftab == xlfMid) return 3 // Mid
        if (iftab == xlfMod) return 2 // Mod
        if (iftab >= xlfDcount && iftab <= xlfDstdev) return 3 // Dxxx formulas
        if (iftab == xlfDvar) return 3 // DVar
        if (iftab == xlfRand) return 0 // Rand
        if (iftab == xlfDate) return 3 // Date
        if (iftab == xlfTime) return 3 // Time
        if (iftab == xlfDay) return 1 // Day
        if (iftab == xlfNow) return 0 // now
        if (iftab == xlfAtan2) return 2 // Atan2
        if (iftab == xlfLog) return 2 // Log
        if (iftab == xlfLeft) return 2 // Left
        if (iftab == xlfRight) return 2 // Right
        if (iftab == xlfTrim) return 1 // Trim
        if (iftab == xlfText) return 2 // Text
        if (iftab == xlfReplace) return 4 // Replace
        if (iftab == xlfExact) return 2 // Exact
        if (iftab == 165) return 2  //TODO:
        if (iftab == xlfDproduct) return 3 // DProduct
        if (iftab == xlfDstdevp) return 3 // DStdDevp
        if (iftab == xlfDvarp) return 3 // DVarP
        if (iftab == xlfDcounta) return 3 // DCountA
        if (iftab == xlfRoundup) return 2 // Roundup
        if (iftab == xlfRounddown) return 2 // Rounddown
        if (iftab == xlfToday) return 0 // today
        if (iftab == xlfDget) return 3 // DGet
        if (iftab == xlfCombin) return 2 // Combin
        if (iftab == xlfFloor) return 2 // Floor
        if (iftab == xlfCeiling) return 2 // Ceiling
        if (iftab == xlfPower) return 2 // Power
        if (iftab == xlfCountif) return 2 // CountIf
        if (iftab == xlfQuartile) return 2
        if (iftab == xlfFrequency) return 2
        if (iftab == xlfCorrel) return 2
        if (iftab == xlfCovar) return 2
        if (iftab == xlfSlope) return 2
        if (iftab == xlfIntercept) return 2
        if (iftab == xlfPearson) return 2
        if (iftab == xlfRsq) return 2
        if (iftab == xlfSteyx) return 2
        if (iftab == xlfCritbinom) return 3
        if (iftab == xlfForecast) return 3
        if (iftab == xlfTrend) return 2
        if (iftab == xlfIsnumber) return 1
        if (iftab == xlfMmult) return 2
        if (iftab == xlfHour) return 1
        if (iftab == xlfMinute) return 1
        if (iftab == xlfMonth) return 1
        if (iftab == xlfYear) return 1
        if (iftab == xlfSecond) return 1
        if (iftab == xlfSqrt) return 1
        if (iftab == xlfExp) return 1
        if (iftab == xlfMirr) return 3
        if (iftab == xlfSyd) return 4
        if (iftab == xlfSln) return 3
        if (iftab == xlfIspmt) return 4
        if (iftab == xlfBinomdist) return 4
        if (iftab == xlfChidist) return 2
        if (iftab == xlfChiinv) return 2
        if (iftab == xlfChitest) return 2
        if (iftab == xlfConfidence) return 3
        if (iftab == xlfFtest) return 2
        if (iftab == xlfSumx2my2) return 2
        if (iftab == xlfSumx2py2) return 2
        if (iftab == xlfSumxmy2) return 2
        if (iftab == xlfLookup) return 3
        if (iftab == xlfTrue) return 0
        if (iftab == xlfFalse) return 0
        if (iftab == xlfExpondist) return 3
        if (iftab == xlfFdist) return 3
        if (iftab == xlfFinv) return 3
        if (iftab == xlfLoginv) return 2
        if (iftab == xlfNegbinomdist) return 3
        if (iftab == xlfNormdist) return 4
        if (iftab == xlfNorminv) return 3
        if (iftab == xlfNormsinv) return 1
        if (iftab == xlfStandardize) return 3
        if (iftab == xlfPermut) return 2
        if (iftab == xlfPoisson) return 3
        if (iftab == xlfSumx2my2) return 2
        if (iftab == xlfSumx2py2) return 2
        if (iftab == xlfSumxmy2) return 2
        if (iftab == xlfTdist) return 2
        if (iftab == xlfLarge) return 2
        return if (iftab == xlfSmall) 2 else 1
//if we are lucky - rest all should be 1 param!
    }
}
