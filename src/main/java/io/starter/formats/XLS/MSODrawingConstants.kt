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
package io.starter.formats.XLS

/**
 * Msodrawing record-related constants
 * NOT A COMPLETE LIST!
 */
object MSODrawingConstants {
    // 20070910 KSC: constants for MSO record + atom ids (records keep atoms and other containers; atoms contain info and are kept inside containers)
    // common record header for both:  ver, inst, fbt, len; fbt deterimes record type (0xF000 to 0xFFFF)
    // MsoDrawingGroup
    val MSOFBTDGGCONTAINER = 0xF000                // Drawing Group Container (Msodrawinggroup)
    val MSOFBTDGG = 0xF006                        // Drawing Group Record (of Msodrawinggoup)
    val MSOFBTCLSID = 0xF016                    // Clipboard format
    val MSOFBTOPT = 0xF00B                        // Property Table Record for newly created shapes, array of FOPTEs
    val MSOFBTCOLORMRU = 0xF11A
    val MSOFBTSPLITMENUCOLORS = 0xF11E
    val MSOFBTBSTORECONTAINER = 0xF001            // Stores BLIPS (=pix) in a separate container
    val MSOFBTBSE = 0xF007                    // BLIP Store Entry Record
    val MSOFBTCALLOUTRULE = 0xF017            // One callout rule per callout shape
    val MSOFBTBLIP = 0xF018
    // MsoDrawing
    val MSOFBTDGCONTAINER = 0xF002                    // Drawing Container - (Msodrawing records contained in MsoDrawinggroup)
    val MSOFBTDG = 0xF008                        // Basic drawing info
    val MSOFBTREGROUPITEMS = 0xF118            // Mappings to reconstitute groups
    val MSOFBTCOLORSCHEME = 0xF120
    val MSOFBTSECONDARYOPT = 0xF121            // secondary opt block - "the property table msofbtOpt may be split into as many as 3 blocks"
    val MSOFBTTERTIARYOPT = 0xF122        // default properties of new shapes - only those props which differ from the per-property defaults are saved
    val MSOFBTSPGRCONTAINER = 0xF003            // Patriarch shape, with all non-bg non-deleted shapes inside it
    val MSOFBTSPCONTAINER = 0xF004                // Shape Container
    val MSOFBTSPGR = 0xF009                // Group-shape-specific info  (i.e. shapes that are groups)
    val MSOFBTSP = 0xF00A                    // A shape atom rec (inst= shape type) rec= shape ID + group of flags
    val MSOFBTTEXTBOX = 0xF00C                // if the shape has text
    val MSOFBTCLIENTTEXTBOX = 0xF00D        // for clipboard stream
    val MSOFBTANCHOR = 0xF00E                // Anchor or location fo a shape (if streamed to a clipboard)
    val MSOFBTCHILDANCHOR = 0xF00F            //   " ", if shape is a child of a group shape
    val MSOFBTCLIENTANCHOR = 0xF010        //   " ", for top-level shapes
    val MSOFBTCLIENTDATA = 0xF011            // content is determined by host
    val MSOFBTCONNECTORRULE = 0xF012    // connector rule
    val MSOFBTALIGNRULE = 0xF013
    val MSOFBTARCRULE = 0xF014
    val MSOFBTCLIENTRULE = 0xF015
    val MSOFBTOLEOBJECT = 0xF11F        //
    val MSOFBTDELETEDPSPL = 0xF11D
    val MSOFBTSOLVERCONTAINER = 0xF005    // 	the rules governing shapes; count of rules
    val MSOFBTSELECTION = 0xF119
    // MSOBI == encoded BLIP types
    val msobiUNKNOWN = 0
    val msobiWMF = 0x216
    val msobiEMF = 0x3D4
    val msobiPICT = 0x542
    val msobiPNG = 0x6E0
    val msobiJFIF = 0x46A
    val msobiJPEG = msobiJFIF
    val msobiDIB = 0x7A8
    val msobiCLIENT = 0x800
    //
    val msofbtBlipFirst = 0xF018            // used to calculate fbt of BLIP in MsofbtBSE

    // pertinent Property Table/MSOFBTOPT property id's
    val msooptfLockAgainstGrouping = 127    // BOOL - Do not group this shape
    val msofbtlTxid = 128                    // LONG - id for the text, value determined by the host
    //margins relative to shape's inscribed text rectangle (in EMUs)
    val dxTextLeft = 129    // 	LONG	1/10 inch
    val dyTextTop = 130 // 	LONG	1/20 inch
    val dxTextRight = 131    // 	LONG 	1/10 inch
    val dyTextBottom = 132    //	LONG 	1/20 inch
    // How to anchor the text
    val anchorText = 135 // 	MSOANCHOR	def= Top
    val msofbttxdir = 139                    // MSOTXDIR - Bi-Di Text direction
    val msooptfFitTextToShape = 191        // BOOL - Size text to fit shape size
    val msooptpib = 260                    // IMsoBlip	- id of Blip to display == imageIndex
    val msooptpibName = 261                // WCHAR - Blip File Name == imageName
    val msooptpibFlags = 262                // MSOBLIPFLAGS - Blip Flags
    val msooptpictureActive = 319            // Server is active (OLE objects only)default= false
    // fill attrbutes
    val msooptFillType = 384
    val msooptfillColor = 385                // MSOCLR - foreground color
    val msooptFillOpacity = 386
    val msooptfillBackColor = 387            // MSOCLR - background color
    val msooptFillBackOpacity = 388
    val msooptFillCrMod = 389    // foreground color of the fill for black-and-white display mode.
    val msooptFillBlip = 390
    val msooptFillBlipName = 391        // fillBlipName_complex -- specifies additional data for the fillBlipName property
    val msooptFillBlipFlags = 392        // specifies how to interpret the fillBlipName_complex property
    val msooptFillWidth = 393        // the width of the fill. This property applies only to texture, picture, and pattern fills.
    // A signed integer that specifies the width of the fill in units that are specified by the fillDztype property, as defined in section 2.3.7.24. If fillDztype equals msodztypeDefault, this value MUST be ignored. The default value for this property is 0x00000000.
    val msooptFillHeight = 394        // A signed integer that specifies the height of the fill in units that are specified by the fillDztype property, as defined in section 2.3.7.24. If fillDztype equals msodztypeDefault, this value MUST be ignored. The default value for this property is 0x00000000.
    val msooptFillAngle = 395        // A value of type FixedPoint, as specified in [MS-OSHARED] section 2.2.1.6, that specifies the angle of the gradient fill. Zero degrees represents a vertical vector from bottom to top. The default value for this property is 0x00000000.
    val msooptFillFocus = 396        // specifies the relative position of the last color in the shaded fill.
    val msooptFillToLeft = 397
    val msooptFillToTop = 398
    val msooptFillToRight = 399
    val msooptFillToBottom = 400        // A signed integer that specifies the left boundary, in EMUs, of the bounding rectangle of the shaded fill. If the Fill Style BooleanfillUseRect property, as defined in section 2.3.7.43, equals 0x0, this value MUST be ignored. The default value for this property is 0x00000000.
    val msooptFillRectLeft = 401        // A signed integer that specifies the top boundary, in EMUs, of the bounding rectangle of the shaded fill. If the Fill Style BooleanfillUseRect property, as defined in section 2.3.7.43, equals 0x0, this value MUST be ignored. The default value for this property is 0x00000000.
    val msooptFillRectTop = 402
    val msooptFillRectRight = 403
    val msooptFillRectBottom = 404
    val msooptFillDztype = 405        // An MSODZTYPE enumeration value, as defined in section 2.4.12, that specifies how the fillWidth, as defined in section 2.3.7.12, and fillHeight, as defined in section 2.3.7.13, properties are interpreted. The default value for this property is msodztypeDefault.
    val msooptFillShadePreset = 406    // A signed integer that specifies the preset colors of the gradient fill. This value MUST be from 0x00000088 through 0x0000009F, inclusive. if the fillShadeColors_complex property, as defined in section 2.3.7.27, exists, this value MUST be ignored. The default value for this property is 0x00000000.
    val msooptFillShadeColors = 407    // The number of bytes of data in the fillShadeColors_complex property. If opid.fComplex equals 0x0, this value MUST be 0x00000000. The default value for this property is 0x00000000.
    val msooptFillOriginX = 408
    val msooptFillOriginY = 409
    val msooptFillShapeOriginX = 410
    val msooptFillShapeOriginY = 411
    val msooptFillShadeType = 412    // : An MSOSHADETYPE record, as defined in section 2.2.50, that specifies how the shaded fill is computed. The default value for this property is msoshadeDefault.
    val msooptFFilled = 443
    // line attributes
    val msooptfNoFillHitTest = 447            // BOOL - Hit test a shape as though filled
    val msooptlineColor = 448                // MSOCLR - Color of line
    val msooptlineWidth = 459                // LONG - 1pt= 12700 EMUs
    val msooptLineMiterLimit = 460
    val msooptLineStyle = 461
    val msooptLineDashing = 462
    val msooptLineDashStyle = 463
    val msooptLineStartArrowhead = 464
    val msooptLineEndArrowhead = 465
    val msooptLineStartArrowWidth = 466
    val msooptLineStartArrowLength = 467
    val msooptLineEndArrowWidth = 468
    val msooptLineEndArrowLength = 469
    val msooptLineJoinStyle = 470
    val msooptLineEndCapStyle = 471
    val msooptFArrowheadsOK = 507
    val msooptLine = 508
    val msooptFHitTestLine = 509
    val msooptLineFillShape = 510
    val msooptfNoLineDrawDash = 511        // BOOL - Draw a dashed line if no line
    // Shadow attributes - many to add
    val msooptfshadowColor = 513            // MSOCLR - foreground color - default = 0x808080
    val msooptfShadowObscured = 575        // BOOL - Excel5-style Shadow
    val msooptfBackground = 831            // BOOL - If true, this is the background shape
    val msooptwzName = 896                    // WCHAR - Shape Name (present only if explicitly set in Named Range box)
    val msooptwzDescription = 897            // WCHAR - Alternate text
    // group shape attributes (selected)
    val msooptidRegroup = 904                // LONG - Regroup id
    val msooptMetroBlob = 937                /*The shape‘s 2007 representation in Office Open XML format.
     																The actual data is a package in Office XML format, which can simply be opened as a zip file.
     																This zip file contains an XML file with the root element ―sp‖.
     																Refer to the publically available Office Open XML documentation for more information about this data.
     																In case we lose any property when converting a 2007 Office Art shape to 2003 shape,
     																we use this blob to retrieve the original Office Art property data when opening the file in 2007.
     																See Appendix F for more information*/
    // misc
    val msooptGroupShapeProperties = 959    // The Group Shape Boolean Properties specify a 32-bit field of Boolean properties for either a shape or a group.
    // ...
    // right now, we support GIF, JPG and PNG
    val IMAGE_TYPE_GIF = 0
    val IMAGE_TYPE_EMF = 2
    val IMAGE_TYPE_WMF = 3
    val IMAGE_TYPE_PICT = 4
    val IMAGE_TYPE_JPG = 5
    val IMAGE_TYPE_PNG = 6
    val IMAGE_TYPE_DIB = 7

    // Shape Types
    /**
     * Internally, a shape type is defined as a fixed set of property values,
     * the most important being the geometry of the shape (the pVertices property, etc.).
     * Each shape stores in itself only those properties that differ from its shape type.
     * When a shape is asked for a property that isn't in its local table,
     * it looks in the shape type's table.
     * If the shape type doesn't define a value for the property,
     * then the property's default value is used.
     */
    val msosptMin = 0
    val msosptNotPrimitive = msosptMin
    val msosptRectangle = 1
    val msosptRoundRectangle = 2
    val msosptEllipse = 3
    val msosptDiamond = 4
    val msosptIsocelesTriangle = 5
    val msosptRightTriangle = 6
    val msosptParallelogram = 7
    val msosptTrapezoid = 8
    val msosptHexagon = 9
    val msosptOctagon = 10
    val msosptPlus = 11
    val msosptStar = 12
    val msosptArrow = 13
    val msosptThickArrow = 14
    val msosptHomePlate = 15
    val msosptCube = 16
    val msosptBalloon = 17
    val msosptSeal = 18
    val msosptArc = 19
    val msosptLine = 20
    val msosptPlaque = 21
    val msosptCan = 22
    val msosptDonut = 23
    val msosptTextSimple = 24
    val msosptTextOctagon = 25
    val msosptTextHexagon = 26
    val msosptTextCurve = 27
    val msosptTextWave = 28
    val msosptTextRing = 29
    val msosptTextOnCurve = 30
    val msosptTextOnRing = 31
    val msosptStraightConnector1 = 32
    val msosptBentConnector2 = 33
    val msosptBentConnector3 = 34
    val msosptBentConnector4 = 35
    val msosptBentConnector5 = 36
    val msosptCurvedConnector2 = 37
    val msosptCurvedConnector3 = 38
    val msosptCurvedConnector4 = 39
    val msosptCurvedConnector5 = 40
    val msosptCallout1 = 41
    val msosptCallout2 = 42
    val msosptCallout3 = 43
    val msosptAccentCallout1 = 44
    val msosptAccentCallout2 = 45
    val msosptAccentCallout3 = 46
    val msosptBorderCallout1 = 47
    val msosptBorderCallout2 = 48
    val msosptBorderCallout3 = 49
    val msosptAccentBorderCallout1 = 50
    val msosptAccentBorderCallout2 = 51
    val msosptAccentBorderCallout3 = 52
    val msosptRibbon = 53
    val msosptRibbon2 = 54
    val msosptChevron = 55
    val msosptPentagon = 56
    val msosptNoSmoking = 57
    val msosptSeal8 = 58
    val msosptSeal16 = 59
    val msosptSeal32 = 60
    val msosptWedgeRectCallout = 61
    val msosptWedgeRRectCallout = 62
    val msosptWedgeEllipseCallout = 63
    val msosptWave = 64
    val msosptFoldedCorner = 65
    val msosptLeftArrow = 66
    val msosptDownArrow = 67
    val msosptUpArrow = 68
    val msosptLeftRightArrow = 69
    val msosptUpDownArrow = 70
    val msosptIrregularSeal1 = 71
    val msosptIrregularSeal2 = 72
    val msosptLightningBolt = 73
    val msosptHeart = 74
    val msosptPictureFrame = 75
    val msosptQuadArrow = 76
    val msosptLeftArrowCallout = 77
    val msosptRightArrowCallout = 78
    val msosptUpArrowCallout = 79
    val msosptDownArrowCallout = 80
    val msosptLeftRightArrowCallout = 81
    val msosptUpDownArrowCallout = 82
    val msosptQuadArrowCallout = 83
    val msosptBevel = 84
    val msosptLeftBracket = 85
    val msosptRightBracket = 86
    val msosptLeftBrace = 87
    val msosptRightBrace = 88
    val msosptLeftUpArrow = 89
    val msosptBentUpArrow = 90
    val msosptBentArrow = 91
    val msosptSeal24 = 92
    val msosptStripedRightArrow = 93
    val msosptNotchedRightArrow = 94
    val msosptBlockArc = 95
    val msosptSmileyFace = 96
    val msosptVerticalScroll = 97
    val msosptHorizontalScroll = 98
    val msosptCircularArrow = 99
    val msosptNotchedCircularArrow = 100
    val msosptUturnArrow = 101
    val msosptCurvedRightArrow = 102
    val msosptCurvedLeftArrow = 103
    val msosptCurvedUpArrow = 104
    val msosptCurvedDownArrow = 105
    val msosptCloudCallout = 106
    val msosptEllipseRibbon = 107
    val msosptEllipseRibbon2 = 108
    val msosptFlowChartProcess = 109
    val msosptFlowChartDecision = 110
    val msosptFlowChartInputOutput = 111
    val msosptFlowChartPredefinedProcess = 112
    val msosptFlowChartInternalStorage = 113
    val msosptFlowChartDocument = 114
    val msosptFlowChartMultidocument = 115
    val msosptFlowChartTerminator = 116
    val msosptFlowChartPreparation = 117
    val msosptFlowChartManualInput = 118
    val msosptFlowChartManualOperation = 119
    val msosptFlowChartConnector = 120
    val msosptFlowChartPunchedCard = 121
    val msosptFlowChartPunchedTape = 122
    val msosptFlowChartSummingJunction = 123
    val msosptFlowChartOr = 124
    val msosptFlowChartCollate = 125
    val msosptFlowChartSort = 126
    val msosptFlowChartExtract = 127
    val msosptFlowChartMerge = 128
    val msosptFlowChartOfflineStorage = 129
    val msosptFlowChartOnlineStorage = 130
    val msosptFlowChartMagneticTape = 131
    val msosptFlowChartMagneticDisk = 132
    val msosptFlowChartMagneticDrum = 133
    val msosptFlowChartDisplay = 134
    val msosptFlowChartDelay = 135
    val msosptTextPlainText = 136
    val msosptTextStop = 137
    val msosptTextTriangle = 138
    val msosptTextTriangleInverted = 139
    val msosptTextChevron = 140
    val msosptTextChevronInverted = 141
    val msosptTextRingInside = 142
    val msosptTextRingOutside = 143
    val msosptTextArchUpCurve = 144
    val msosptTextArchDownCurve = 145
    val msosptTextCircleCurve = 146
    val msosptTextButtonCurve = 147
    val msosptTextArchUpPour = 148
    val msosptTextArchDownPour = 149
    val msosptTextCirclePour = 150
    val msosptTextButtonPour = 151
    val msosptTextCurveUp = 152
    val msosptTextCurveDown = 153
    val msosptTextCascadeUp = 154
    val msosptTextCascadeDown = 155
    val msosptTextWave1 = 156
    val msosptTextWave2 = 157
    val msosptTextWave3 = 158
    val msosptTextWave4 = 159
    val msosptTextInflate = 160
    val msosptTextDeflate = 161
    val msosptTextInflateBottom = 162
    val msosptTextDeflateBottom = 163
    val msosptTextInflateTop = 164
    val msosptTextDeflateTop = 165
    val msosptTextDeflateInflate = 166
    val msosptTextDeflateInflateDeflate = 167
    val msosptTextFadeRight = 168
    val msosptTextFadeLeft = 169
    val msosptTextFadeUp = 170
    val msosptTextFadeDown = 171
    val msosptTextSlantUp = 172
    val msosptTextSlantDown = 173
    val msosptTextCanUp = 174
    val msosptTextCanDown = 175
    val msosptFlowChartAlternateProcess = 176
    val msosptFlowChartOffpageConnector = 177
    val msosptCallout90 = 178
    val msosptAccentCallout90 = 179
    val msosptBorderCallout90 = 180
    val msosptAccentBorderCallout90 = 181
    val msosptLeftRightUpArrow = 182
    val msosptSun = 183
    val msosptMoon = 184
    val msosptBracketPair = 185
    val msosptBracePair = 186
    val msosptSeal4 = 187
    val msosptDoubleWave = 188
    val msosptActionButtonBlank = 189
    val msosptActionButtonHome = 190
    val msosptActionButtonHelp = 191
    val msosptActionButtonInformation = 192
    val msosptActionButtonForwardNext = 193
    val msosptActionButtonBackPrevious = 194
    val msosptActionButtonEnd = 195
    val msosptActionButtonBeginning = 196
    val msosptActionButtonReturn = 197
    val msosptActionButtonDocument = 198
    val msosptActionButtonSound = 199
    val msosptActionButtonMovie = 200
    val msosptHostControl = 201
    //Host controls extend various user interface (UI) objects in the Word and Excel object models
    val msosptTextBox = 202
    val msosptMax = 0x0FFF
    val msosptNil = 0x0FFF
}
