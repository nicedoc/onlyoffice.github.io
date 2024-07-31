import { map_base64 } from '../resources/list_base64.js'
import { ZONE_TYPE, ZONE_SIZE, ZONE_TYPE_NAME } from './model/feature.js'
import { biyueCallCommand, dispatchCommandResult } from "./command.js";

var loading = false // 正在绘制中
var list_command = [] // 操作列表
var list_wait_command = [] // 等待执行的操作列表


var c_oAscRelativeFromH = {
	Character: 0,
	Column: 1,
	InsideMargin: 2,
	LeftMargin: 3,
	Margin: 4,
	OutsideMargin: 5,
	Page: 6,
	RightMargin: 7
}

var c_oAscRelativeFromV = {
	BottomMargin: 0,
	InsideMargin: 1,
	Line: 2,
	Margin: 3,
	OutsideMargin: 4,
	Page: 5,
	Paragraph: 6,
	TopMargin: 7
}

function handleFeature(options) {
	options.size = Object.assign({}, ZONE_SIZE[options.zone_type], (options.size || {}))
	if (options.v == undefined) {
		options.v = 1
	}
	options.page_num = options.p || 0
	options.type = 'feature'
	options.type_name = ZONE_TYPE_NAME[options.zone_type]
	addCommand(options)
	if (loading) {
		console.log('loading...')
		return
	}
	drawList([options]).then(() => {
		setLoading(false)
		handleNext()
	})
}

function drawExtroInfo(list) {
	list.forEach(e => {
		e.page_num = e.p || 0
		if (e.v == undefined) {
			e.v = 1
		}
		e.size = Object.assign({}, ZONE_SIZE[e.zone_type], (e.size || {}))
		e.type = 'feature'
		e.type_name = ZONE_TYPE_NAME[e.zone_type]
	})
	return drawList(list)
}
// 整理参数
function addCommand(options) {
	if (!list_wait_command) {
		list_wait_command = []
	}
	var index = list_wait_command.findIndex((e) => {
		return e.zone_type == options.zone_type && e.v == options.v
	})
	if (index == -1) {
		list_wait_command.push(options)
		return
	}
	list_wait_command[index] = options
}

function handleNext() {
	setLoading(false)
	if (list_wait_command && list_wait_command.length > 0) {
		var newlist = []
		var type = list_wait_command[0].type
		newlist.push(Object.assign({}, list_wait_command[0]))
		var end = -1
		for (var i = 1; i < list_wait_command.length; ++i) {
			if (list_wait_command[i].type != type) {
				list_wait_command.splice(0, i)
				end = i
				break
			}
			newlist.push(Object.assign({}, list_wait_command[i]))
		}
		if (end == -1) {
			list_wait_command = []
		}
		if (type == 'feature') {
			drawList(newlist).then(() => {
				handleNext()
			})
		} else if (type == 'header') {
			drawHeader(
				newlist[newlist.length - 1].cmd,
				newlist[newlist.length - 1].title
			)
		}
	}
}

function handleHeader(cmdType, examTitle) {
	addCommand({
		type: 'header',
		title: examTitle,
		cmd: cmdType,
	})
	if (loading) {
		console.log('loading...')
		return
	}
	drawHeader(cmdType, examTitle)
}

function drawHeader(cmdType, examTitle) {
	Asc.scope.header_cmd = cmdType
	Asc.scope.header_exam_title = examTitle
	Asc.scope.qrcode_url = map_base64.qrcode
	loading = true
	biyueCallCommand(window, function() {
		var cmdType = Asc.scope.header_cmd
		var examTitle = Asc.scope.header_exam_title
		var qrcode_url = Asc.scope.qrcode_url
		var MM2TWIPS = 25.4 / 72 / 20
		var scale = 0.25
		var oDocument = Api.GetDocument()
		var oSections = oDocument.GetSections()
		if (!oSections || oSections.length == 0) {
			return
		}
		if (oSections[0].GetHeader('title', false)) {
			oSections[0].RemoveHeader('title')
		}
		if (oSections.length > 1) {
			if (oSections[1].GetHeader('default', false)) {
				oSections[1].RemoveHeader('default')
			}
		}
		if (cmdType == 'close') {
			return
		}
		function createAgainDrawing() {
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(
				3600,
				Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
			)
			var oDrawing = Api.CreateShape(
				'rect',
				42 * scale * 36e3,
				24 * scale * 36e3,
				oFill,
				oStroke
			)
			var drawDocument = oDrawing.GetContent()
			var paragraphs = drawDocument.GetAllParagraphs()
			if (paragraphs && paragraphs.length > 0) {
				var oRun = Api.CreateRun()
				oRun.AddText('再练')
				paragraphs[0].AddElement(oRun)
				paragraphs[0].SetColor(3, 3, 3, false)
				paragraphs[0].SetFontSize(14)
				paragraphs[0].SetJc('center')
			}
			oDrawing.SetVerticalTextAlign('center')
			oDrawing.SetPaddings(0, 0, 0, 0)
			return oDrawing
		}

		function setHeader(oSection, oHeader, showAgain) {
			if (!oHeader) {
				return
			}
			var pmargins = oSection.Section.PageMargins
			var pSize = oSection.Section.PageSize
			var pw = pSize.W - pmargins.Left - pmargins.Right
			var oParagraph = oHeader.GetElement(0)
			oParagraph.SetTabs(
				[1, (pw * 0.5) / MM2TWIPS, pw / MM2TWIPS],
				['left', 'center', 'right']
			)
			oParagraph.AddTabStop()
			if (showAgain) {
				oParagraph.AddDrawing(createAgainDrawing())
			}
			oParagraph.AddTabStop()
			oParagraph.AddText(examTitle)
			oParagraph.AddTabStop()
			var oDrawing2 = Api.CreateImage(
				qrcode_url,
				45 * scale * 36e3,
				45 * scale * 36e3
			)
			oParagraph.AddDrawing(oDrawing2)
			oParagraph.SetBottomBorder('single', 8, 3, 153, 153, 153)
		}
		var firstSection = oSections[0]
		firstSection.SetTitlePage(true)
		var oHeader = firstSection.GetHeader('title', true)
		setHeader(firstSection, oHeader, true)

		if (oSections.length > 1) {
			var oHeader2 = oSections[1].GetHeader('default', true)
			setHeader(oSections[1], oHeader2, false)
		}
	}, false, true).then(res => {
		handleNext()
	})
}

function deleteAllFeatures(exceptList, specifyFeatures) {
	Asc.scope.exceptList = exceptList
	Asc.scope.specifyFeatures = specifyFeatures
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects()
		var exceptList = Asc.scope.exceptList
		var specifyFeatures = Asc.scope.specifyFeatures
		function deleteAccurate(oDrawing) {
			var run = oDrawing.Drawing.GetRun()
			if (run) {
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0) {
						oDrawing.Delete()
						if (oParagraph) {
							oParagraph.RemoveElement(ipos)
						}
						return
					}
				}
			}
			oDrawing.Delete()
		}
		function deleteDrawing(titleObj, oDrawing) {
			if (titleObj.feature.sub_type == 'ask_accurate' ||
			titleObj.feature.zone_type == 'pagination' ||
			titleObj.feature.zone_type == 'statistics') {
				deleteAccurate(oDrawing)
			} else {
				oDrawing.Delete()
			}
		}
		if (drawings) {
			for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
				var oDrawing = drawings[j]
				if (oDrawing.Drawing.docPr) {
					var title = oDrawing.Drawing.docPr.title
					if (title && title.indexOf('feature') >= 0) {
						var titleObj = JSON.parse(title)
						if (titleObj.feature && titleObj.feature.zone_type) {
							if (exceptList) {
								var inExcept = exceptList.findIndex(e => {
									return e.zone_type == titleObj.feature.zone_type
								})
								if (inExcept >= 0) {
									continue
								}
							}
							if (specifyFeatures) {
								var inSpecify = specifyFeatures.find(e => {
									return title.indexOf(e) >= 0
								})
								if (inSpecify) {
									deleteDrawing(titleObj, oDrawing)
								}
							} else {
								deleteDrawing(titleObj, oDrawing)
							}
						}
					}
				}
			}
		}
		function getFirstParagraph(oControl) {
			if (!oControl) {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
			for (var i = 0; i < paragraphs.length; ++i) {
				var oParagraph = paragraphs[i]
				if (oParagraph) {
					var parent1 = oParagraph.Paragraph.Parent
					var parent2 = parent1.Parent
					if (parent2 && parent2.Id == oControl.Sdt.GetId()) {
						return oParagraph
					}
				}
			}
			return null
		}
		var handledNumbering = {}
		function hideSimple(oParagraph) {
			if (!oParagraph) {
				return
			}
			var oNumberingLevel = oParagraph.GetNumbering()
			if (!oNumberingLevel) {
				return
			}
			var level = oNumberingLevel.Lvl
			var oNum = oNumberingLevel.Num
			var oNumberingLvl = oNum.GetLvl(level)
			if (!oNumberingLvl) {
				return
			}
			var LvlText = oNumberingLvl.LvlText || []
			if (LvlText && LvlText.length) {
				if (LvlText[0].Value!='\ue6a1') {
					return
				}
			}
			var key = `${oNum.Id}_${level}`
			if (handledNumbering[key]) {
				return
			}
			handledNumbering[key] = 1
			var suffix = ''
			if (LvlText.length > 1 && LvlText[LvlText.length - 1].Type == 1) {
				suffix = LvlText[LvlText.length - 1].Value
			}
			var sType = 'decimal'
			if (oNumberingLvl.Format == 8) {
				sType = 'chineseCounting'
			} else if (oNumberingLvl.Format == 9) {
				sType = 'chineseCountingThousand'
			} else if (oNumberingLvl.Format == 10) {
				sType = 'chineseLegalSimplified'
			} else if (oNumberingLvl.Format == 14) {
				sType = 'decimalEnclosedCircle'
			} else if (oNumberingLvl.Format == 15) {
				sType = 'decimalEnclosedCircleChinese'
			} else if (oNumberingLvl.Format == 21) {
				sType = 'decimalZero'
			} else if (oNumberingLvl.Format == 46) {
				sType = 'lowerLetter'
			} else if (oNumberingLvl.Format == 47) {
				sType = 'lowerRoman'	
			} else if (oNumberingLvl.Format == 60) {
				sType = 'upperLetter'
			} else if (oNumberingLvl.Format == 61) {
				sType = 'upperRoman'
			}
			oNumberingLevel.SetTemplateType('bullet', '');
			var str = ''
			var find = false
			for (var i = 0; i < LvlText.length; ++i) {
				if (LvlText[i].Type == 2) {
					str += `%${level+1}`
				} else {
					if (LvlText[i].Value == ' ') {
						if (find) {
							str += LvlText[i].Value
						}
					} else if (LvlText[i].Value != '\ue6a1') {
						str += LvlText[i].Value
						find = true
					}
				}
			}
			oNumberingLevel.SetCustomType(sType, str, "left")
			var oTextPr = oNumberingLevel.GetTextPr();
			oTextPr.SetFontFamily("iconfont");
		}
		var controls = oDocument.GetAllContentControls()
		if (controls) {
			for (var j = 0, jmax = controls.length; j < jmax; ++j) {
				var oControl = controls[j]
				if (oControl.GetClassType() == 'blockLvlSdt') {
					var firstParagraph = getFirstParagraph(oControl)
					hideSimple(firstParagraph)
				}
			}
		}
	}, false, true).then(() => {
		console.log('功能区都已删除')
	})
}

function drawList(list) {
	setLoading(true)
	Asc.scope.feature_wait_handle = list
	Asc.scope.ZONE_TYPE = ZONE_TYPE
	return biyueCallCommand(window, function() {
		var ZONE_TYPE = Asc.scope.ZONE_TYPE
		var MM2TWIPS = 25.4 / 72 / 20
		var oDocument = Api.GetDocument()
		var objs = oDocument.GetAllDrawingObjects()
		var oSections = oDocument.GetSections()
		var feature_wait_handle = Asc.scope.feature_wait_handle
		feature_wait_handle = feature_wait_handle.filter(e => {
			return e.zone_type
		})
		var elementsCount = oDocument.GetElementsCount()
		var lastElement = oDocument.GetElement(elementsCount - 1)
		var pageCount = oDocument.GetPageCount()

		var lastParagraph = null
		if (lastElement.GetClassType() == 'paragraph') {
			lastParagraph = lastElement
		} else {
			lastParagraph = Api.CreateParagraph()
			oDocument.push(lastParagraph)
		}
		var res = {
			list: [],
		}
		function addImageToCell(oTable, nRow, nCell, url, w, h, mleft, mright) {
			if (!oTable) {
				return
			}
			var cell = oTable.GetCell(nRow, nCell)
			if (!cell) {
				return
			}
			var oCellContent = cell.GetContent()
			if (oCellContent) {
				var p = oCellContent.GetElement(0)
				if (p && p.GetClassType() == 'paragraph') {
					var oImage = Api.CreateImage(url, w * 36e3, h * 36e3)
					p.AddDrawing(oImage)
				}
			}
			cell.SetWidth('twips', w / MM2TWIPS)
			cell.SetVerticalAlign('center')
			if (mleft != undefined) {
				cell.SetCellMarginLeft(mleft)
			}
			if (mright != undefined) {
				cell.SetCellMarginRight(mright)
			}
			cell.SetCellMarginTop(0)
			cell.SetCellMarginBottom(0)
		}
		function addTextToCell(
			oTable,
			nRow,
			nCell,
			texts,
			jc,
			w,
			fontSize,
			mleft,
			mright
		) {
			if (!oTable) {
				return
			}
			var cell = oTable.GetCell(nRow, nCell)
			var oCellContent = cell.GetContent()
			if (oCellContent) {
				var p = oCellContent.GetElement(0)
				if (p && p.GetClassType() == 'paragraph') {
					var oRun = Api.CreateRun()
					for (var i = 0; i < texts.length; ++i) {
						oRun.AddText(texts[i])
						if (i < texts.length - 1) {
							oRun.AddLineBreak()
						}
					}
					p.AddElement(oRun)
					p.SetFontSize(fontSize)
					p.SetJc(jc)
					p.SetColor(33, 33, 33, false)
				}
			}
			cell.SetVerticalAlign('center')
			cell.SetCellMarginLeft(0)
			cell.SetCellMarginRight(0)
			cell.SetWidth('twips', w / MM2TWIPS)
			if (mleft != undefined) {
				cell.SetCellMarginLeft(mleft)
			}
			if (mright != undefined) {
				cell.SetCellMarginRight(mright)
			}
		}
		// 获取用于放置drawing的paragraph
		/* ApiDocument.AddDrawingToPage的逻辑是
		// ApiDocument.Document.GoToPage(nPage);
		// var paragraph = ApiDocument.Document.GetCurrentParagraph();		
		// paragraph.AddToParagraph(drawing); => var oRun = new AscCommonWord.ParaRun;  oRun.AddToContent(0, this); oParagraph.AddToContent(0, oRun)
		*/
		// 因此当存在段落跨页时，若直接调用AddDrawingToPage，会出现渲染到前页的情况
		// 需要动态计算出适合的paragraph
		// GoToPage后，调用GetCurrentParagraph，此时拿到的未必是第一个段落
		function GetParagraphForDraw(page_num, posType) {
			var pageCount = oDocument.GetPageCount()
			if (page_num >= pageCount) {
				return null
			}
			if (page_num == 0) {
				if (posType == 'end' && page_num == pageCount - 1) {
					return lastParagraph.Paragraph
				}
				var nearestPos = oDocument.Document.Get_NearestPos(page_num, 0, 0)
				if (nearestPos.Paragraph) {
					return nearestPos.Paragraph
				}
			} else {
				if (posType == 'end') {
					if (page_num == pageCount - 1 && lastParagraph) {
						return lastParagraph.Paragraph
					}
					var PageSize = oSections[0].Section.PageSize
					var PageMargins = oSections[0].Section.PageMargins
					var nearestPos = oDocument.Document.Get_NearestPos(page_num, PageSize.W - PageMargins.Right, PageSize.H - PageMargins.Bottom)
					if (nearestPos.Paragraph) {
						if (nearestPos.Paragraph.Pages.length > 0) {
							if (nearestPos.Paragraph.GetAbsolutePage(0) == page_num) {
								var bounds = nearestPos.Paragraph.Pages[0].Bounds
								if (bounds && bounds.Bottom - bounds.Top > 4) {
									return nearestPos.Paragraph
								} else {
									var previosParagraph = Api.LookupObject(nearestPos.Paragraph.Id).GetPrevious()
									if (previosParagraph) {
										return previosParagraph.Paragraph
									}
								}
							} else {
								// todo..
							}
						}
						return nearestPos.Paragraph
					}
				}
			}
			return null
		}

		function getPageNumberDrawing(shapeWidth, shapeHeight) {
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(0, Api.CreateNoFill())
			var oDrawing = Api.CreateShape(
				'rect',
				shapeWidth * 36e3,
				shapeHeight * 36e3,
				oFill,
				oStroke
			)
			var drawDocument = oDrawing.GetContent()
			var paragraphs = drawDocument.GetAllParagraphs()
			if (paragraphs && paragraphs.length > 0) {
				paragraphs[0].AddPageNumber()
				paragraphs[0].SetJc('center')
				paragraphs[0].SetColor(3, 3, 3, false)
				var twips = shapeHeight / (25.4 / 72 / 20)
				paragraphs[0].SetFontSize(twips / 10)
				paragraphs[0].SetSpacingAfter(0)
				oDrawing.SetPaddings(0, 0, 0, 0)
			}
			oDrawing.SetVerticalTextAlign('center')
			oDrawing.SetWrappingStyle('inFront')
			var titleobj = {
				feature: {
					zone_type: 'pagination'
				}
			}
			oDrawing.Drawing.Set_Props({
				title: JSON.stringify(titleobj),
			})
			return oDrawing
		}
		function setPaginationAlign(oDrawing, align, PageMargins, margin) {
			if (!oDrawing) {
				return
			}
			if (align == 'center') {
				oDrawing.SetHorAlign('page', align)
			} else if (align == 'left') {
				oDrawing.SetHorPosition('leftMargin', margin * 36e3)
			} else if (align == 'right') {
				oDrawing.SetHorPosition('rightMargin', (0 - margin) * 36e3)
			}
			var drawDocument = oDrawing.GetContent()
			var paragraphs = drawDocument.GetAllParagraphs()
			if (paragraphs && paragraphs.length > 0) {
				paragraphs[0].SetJc(align)
			}
		}
		function getTypeDrawing(oParagraph, zone_type) {
			if (!oParagraph) {
				return false
			}
			var count = oParagraph.GetElementsCount()
			for (var i = 0; i < count; ++i) {
				var oRun = oParagraph.GetElement(i)
				if (oRun && oRun.GetClassType() == 'run') {
					for (var j = 0, jmax = oRun.Run.Content.length; j < jmax; ++j) {
						if (oRun.Run.Content[j].GetType && oRun.Run.Content[j].GetType() == 22) {
							var title = oRun.Run.Content[j].docPr.title || '{}'
							var titleObj = JSON.parse(title)
							if (titleObj.feature && titleObj.feature.zone_type == zone_type) {
								return oRun.Run.Content[j]
							}
						}
					}
				}
			}
			return null
		}
		function deleteDrawing(oDrawing) {
			var run = oDrawing.Drawing.GetRun()
			if (run) {
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0) {
						oDrawing.Delete()
						if (oParagraph) {
							oParagraph.RemoveElement(ipos)
						}
						return
					}
				}
			}
			oDrawing.Delete()
		}
		feature_wait_handle.forEach((options) => {
			var props_title = JSON.stringify({
				feature: {
					zone_type: options.type_name,
					v: options.v
				}
			})
			var find = objs.find(e => {
				return e.Drawing && e.Drawing.docPr && e.Drawing.docPr.title == props_title
			})
			var oDrawing = find
			var result = {
				code: 0,
				zone_type: options.zone_type,
				v: options.v,
				cmd: options.cmd,
			}
			if (options.cmd == 'close') {	// 关闭
				if (oDrawing) {
					result.code = 1
					if (options.type_name == 'statistics') {
						var statisticsList = objs.filter(e => {
							return e.Drawing && e.Drawing.docPr && e.Drawing.docPr.title == props_title
						})
						if (statisticsList) {
							statisticsList.forEach(e => {
								deleteDrawing(e)
							})
						}
						var props_title2 = JSON.stringify({
							feature: {
								zone_type: 'pagination',
								v: options.v
							}
						})
						var findList2 = objs.filter(e => {
							return e.Drawing && e.Drawing.docPr && e.Drawing.docPr.title == props_title2
						})
						if (findList2) {
							findList2.forEach(e => {
								deleteDrawing(e)
							})
						}
					} else {
						oDrawing.Delete()
					}
				} else {
					result.code = 2
					result.message = '未找到该区域'
				}
			} else {
				if (!oDrawing) {
					if (options.size) {
						var shapeWidth = options.size.w
						var shapeHeight = options.size.h
						if (options.zone_type == ZONE_TYPE.QRCODE) {
							var showtext = options.texts && options.texts.length
							if (showtext) {
								shapeWidth += shapeWidth * 1.5
								shapeHeight += 2
							}
							var oFill = Api.CreateNoFill()
							var oStroke = showtext ? Api.CreateStroke(0, Api.CreateNoFill()) : Api.CreateStroke(1, Api.CreateSolidFill(Api.CreateRGBColor(225, 225, 225)))
							oDrawing = Api.CreateShape(
								'rect',
								shapeWidth * 36e3,
								shapeHeight * 36e3,
								oFill,
								oStroke
							)
							if (showtext) {
								var oTable = Api.CreateTable(2, 1)
								addImageToCell(
									oTable,
									0,
									0,
									options.url,
									options.size.imgSize,
									options.size.imgSize
								)
								oTable
									.GetCell(0, 0)
									.SetCellBorderRight('single', 1, 0.1, 255, 255, 255)
								oTable.SetTableBorderTop('single', 1, 0.1, 255, 255, 255)
								oTable.SetTableBorderBottom('single', 1, 0.1, 255, 255, 255)
								oTable.SetTableBorderLeft('single', 1, 0.1, 255, 255, 255)
								oTable.SetTableBorderRight('single', 1, 0.1, 255, 255, 255)
								addTextToCell(
									oTable,
									0,
									1,
									options.texts, // texts内容类似['编码-未生成', '吕老师']
									'center',
									shapeWidth - options.size.w - 2,
									14
								)
								oTable.SetWidth('twips', shapeWidth / MM2TWIPS)
								var oDrawingContent = oDrawing.GetContent()
								oDrawingContent.AddElement(0, oTable)
							} else {
								oDrawing.SetPaddings(0, 0, 0, 0)
								var shapeContent = oDrawing.GetContent()
								var p = shapeContent.GetElement(0)
								if (p && p.GetClassType() == 'paragraph') {
									var oImage = Api.CreateImage(options.url, (options.size.imgSize) * 36e3, (options.size.imgSize) * 36e3)
									p.AddDrawing(oImage)
									p.SetSpacingAfter(0)
								}
							}
						} else if (options.zone_type == ZONE_TYPE.THER_EVALUATION || options.zone_type == ZONE_TYPE.SELF_EVALUATION) {
							var flowers = options.flowers || []
							var oTable = Api.CreateTable(2 + flowers.length, 1)
							var scale = 0.25
							var flowersize = 24
							var fw = flowersize * scale
							var fh = flowersize * scale
							var textw = 24
							var w = 21.33 * scale + textw + fw * flowers.length
							var fmargin = 1 / MM2TWIPS
							addImageToCell(
								oTable,
								0,
								0,
								options.icon_url,
								21.33 * scale,
								30 * scale,
								0,
								0
							)
							addTextToCell(
								oTable,
								0,
								1,
								[options.label],
								'center',
								textw,
								20,
								0,
								fmargin
							)
							var cellindex = 2
							if (flowers) {
								flowers.forEach((url, index) => {
									addImageToCell(oTable, 0, cellindex + index, url, fw, fh, fmargin, fmargin)	
								})
							}
							oTable.GetRow(0).SetHeight('auto', shapeHeight / MM2TWIPS)

							shapeWidth = w + 1
							var oFill = Api.CreateNoFill()
							var oStroke = Api.CreateStroke(0, Api.CreateNoFill())
							oDrawing = Api.CreateShape(
								'rect',
								shapeWidth * 36e3,
								shapeHeight * 36e3,
								oFill,
								oStroke
							)
							var oDrawingContent = oDrawing.GetContent()
							oDrawingContent.AddElement(0, oTable)
							oDrawing.SetPaddings(0, 0, 0, 0)
						} else if (options.url) {
							var oFill = Api.CreateNoFill()
							var oStroke = Api.CreateStroke(0, Api.CreateNoFill())
							oDrawing = Api.CreateShape(
								'rect',
								shapeWidth * 36e3,
								shapeHeight * 36e3,
								oFill,
								oStroke
							)
							oDrawing.SetPaddings(0, 0, 0, 0)
							var shapeContent = oDrawing.GetContent()
							var p = shapeContent.GetElement(0)
							if (p && p.GetClassType() == 'paragraph') {
								var oImage = Api.CreateImage(options.url, (options.size.w) * 36e3, (options.size.h) * 36e3)
								p.AddDrawing(oImage)
							}
						} else if (options.label) {
							var oFill = Api.CreateNoFill()
							var stroke_width = options.size.stroke_width || 0.1
							var oStroke = Api.CreateStroke(
								stroke_width * 36e3,
								Api.CreateSolidFill(Api.CreateRGBColor(225, 225, 225))
							)
							oDrawing = Api.CreateShape(
								options.size.shape_type || 'rect',
								shapeWidth * 36e3,
								shapeHeight * 36e3,
								oFill,
								oStroke
							)
							var drawDocument = oDrawing.GetContent()
							var paragraphs = drawDocument.GetAllParagraphs()
							if (paragraphs && paragraphs.length > 0) {
								var oRun = Api.CreateRun()
								oRun.AddText(options.label)
								paragraphs[0].AddElement(oRun)
								paragraphs[0].SetColor(3, 3, 3, false)
								paragraphs[0].SetFontSize(14)
								paragraphs[0].SetSpacingAfter(0)
								var jc = options.size && options.size.jc ? options.size.jc : 'center'
								paragraphs[0].SetJc(jc)
								if (jc == 'center') {
									oDrawing.SetPaddings(0, 0, 0, 0)
								} else {
									oDrawing.SetPaddings(2 * 36e3, 0, 0, 0)
								}
							}
							oDrawing.SetVerticalTextAlign('center')
						}
						if (oDrawing) {
							oDrawing.Drawing.Set_Props({
								title: props_title,
							})
							if (options.zone_type == ZONE_TYPE.STATISTICS) {
								var numberDrawing = getPageNumberDrawing(20, options.size.pagination.font_size)
								var pstyle = options.size.pagination.align_style
								oDocument.SetEvenAndOddHdrFtr(pstyle != 'center');
								oSections.forEach((section, index) => {
									var footerList = []
									// 首页
									var oTitleFooter = section.GetFooter("title", false)
									if (oTitleFooter) {
										footerList.push({
											type: 'title',
											oFooter: oTitleFooter
										})
									} else {
										var oTitleHeader = section.GetHeader('title', false)
										if (oTitleHeader) { // 存在首页页眉
											oTitleFooter = section.GetFooter('title', true);
											footerList.push({
												type: 'title',
												oFooter: oTitleFooter
											})
										}
									}
									if (pstyle != 'center') {
										var evenFooter = section.GetFooter('even', false)
										if (!evenFooter) {
											evenFooter = section.GetFooter('even', true)
										}
										footerList.push({
											type: 'even',
											oFooter: evenFooter
										})
									}
									var oDefaultFooter = section.GetFooter("default", false)
									if (!oDefaultFooter) {
										oDefaultFooter = section.GetFooter("default", true)
									}
									footerList.push({
										type: 'default',
										oFooter: oDefaultFooter
									})
									footerList.forEach((footerObj) => {
										var oFooter = footerObj.oFooter
										var elementCount = oFooter.GetElementsCount()
										if (elementCount > 2) {
											for(var i = elementCount - 1; i > 0; i--) {
												oFooter.RemoveElement(i)
											}
										}
										var PageMargins = section.Section.PageMargins
										var oParagraph = oFooter.GetElement(0)
										// 统计
										var typeDrawing = getTypeDrawing(oParagraph, options.type_name)
										if (typeDrawing) {
											typeDrawing.Set_PositionH(7, false, PageMargins.Right - options.size.right - options.size.w, false);
											typeDrawing.Set_PositionV(0, false, PageMargins.Bottom - options.size.bottom - options.size.h, false)
											typeDrawing.Set_DrawingType(2);
										} else {
											var oAdd = oDrawing.Copy()
											var drawing = oAdd.Drawing
											drawing.Set_PositionH(7, false, PageMargins.Right - options.size.right - options.size.w, false);
											drawing.Set_PositionV(0, false, PageMargins.Bottom - options.size.bottom - options.size.h, false)
											drawing.Set_DrawingType(2);
											oParagraph.SetJc('center')
											var titleobj = {
												feature: {
													zone_type: options.type_name,
													v: options.v
												}
											}
											oAdd.Drawing.Set_Props({
												title: JSON.stringify(titleobj),
											})
											oParagraph.AddDrawing(oAdd)
										}
										// 页码
										var oAddNum = getTypeDrawing(oParagraph, 'pagination')
										var needAdd = false
										if (!oAddNum) {
											oAddNum = numberDrawing.Copy()
											needAdd = true
										}
										if (oAddNum) {
											oAddNum = numberDrawing.Copy()
											var titleobj = {
												feature: {
													zone_type: 'pagination',
													v: options.v
												}
											}
											oAddNum.Drawing.Set_Props({
												title: JSON.stringify(titleobj)
											})
											var align = 'center'
											if (pstyle == 'oddLeftEvenRight') {
												align = footerObj.type == 'even' ? 'right' : 'left'
											} else if (pstyle == 'oddRightEvenLeft') {
												align = footerObj.type == 'even' ? 'left' : 'right'
											}
											setPaginationAlign(oAddNum, align, PageMargins, options.size.pagination.margin)
											oAddNum.SetVerPosition('bottomMargin', (PageMargins.Bottom - options.size.pagination.bottom - options.size.pagination.font_size) * 36e3)
											if (needAdd) {
												oParagraph.AddDrawing(oAddNum)
											}
										}
									})									
								})								
							} else {
								var page_num = options.page_num || options.p
								if (options.zone_type == ZONE_TYPE.THER_EVALUATION ||
									options.zone_type == ZONE_TYPE.SELF_EVALUATION ||
									options.zone_type == ZONE_TYPE.PASS ||
									options.zone_type == ZONE_TYPE.END ||
									options.zone_type == ZONE_TYPE.IGNORE) {
										page_num = pageCount - 1
									}
								var paragraph = GetParagraphForDraw(page_num, 'end')
								if (paragraph) {
									var drawing = oDrawing.Drawing;
										drawing.Set_PositionH(6, false, options.x, false);
										drawing.Set_PositionV(5, false, options.y, false);
										drawing.Set_DrawingType(2);
										if (lastParagraph && paragraph.Id == lastParagraph.Paragraph.Id) {
											// lastParagraph.Paragraph.AddToParagraph(drawing)
											lastParagraph.AddDrawing(oDrawing)
											oDrawing.Drawing.Set_Parent(lastParagraph.Paragraph)
										} else {
											paragraph.AddToParagraph(drawing);
										}
								} else {
									console.log('cannot find paragraph for draw', page_num)
								}
							}
						}
					}
				} else {
					if (oDrawing && options.zone_type != ZONE_TYPE.STATISTICS) {
						oDrawing.SetHorPosition('page', options.x * 36e3)
						oDrawing.SetVerPosition('page', options.y * 36e3)
					}
				}
				result.code = 1
				result.x = options.x
				result.y = options.y
				result.drawing_id = oDrawing ? oDrawing.Drawing.Id : 0
			}
			res.list.push(result)
		})
		console.log('=====================drawList end ')
		return res
	}, false, true)
}

function setLoading(v) {
	loading = v
}

function setInteraction(type, quesIds) {
	Asc.scope.interaction_type = type
	Asc.scope.interaction_quesIds = quesIds
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.node_list = window.BiyueCustomData.node_list
	return biyueCallCommand(window, function() {
		var interaction_type = Asc.scope.interaction_type
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var question_map = Asc.scope.question_map || {}
		var node_list = Asc.scope.node_list || []
		var MM2TWIPS = 25.4 / 72 / 20
		var quesIds = Asc.scope.interaction_quesIds
		if (!controls) {
			return
		}
		var handledNumbering = {}
		function showSimple(oParagraph, vshow) {
			var oNumberingLevel = oParagraph.GetNumbering()
			if (!oNumberingLevel) {
				return
			}
			var level = oNumberingLevel.Lvl
			var oNum = oNumberingLevel.Num
			if (!oNum) {
				return
			}
			var oNumberingLvl = oNum.GetLvl(level)
			if (!oNumberingLvl) {
				return
			}
			var LvlText = oNumberingLvl.LvlText || []
			if (LvlText && LvlText.length) {
				if (LvlText[0].Value=='\ue6a1') {
					if (vshow) {
						console.log('当前简单互动已显示')
						return
					}
				} else {
					if (!vshow) {
						console.log('当前简单互动本就未显示')
						return
					}
				}
			}
			var key = `${oNum.Id}_${level}`
			if (handledNumbering[key]) {
				return
			}
			handledNumbering[key] = 1
			var sType = 'decimal'
			if (oNumberingLvl.Format == 8) {
				sType = 'chineseCounting'
			} else if (oNumberingLvl.Format == 9) {
				sType = 'chineseCountingThousand'
			} else if (oNumberingLvl.Format == 10) {
				sType = 'chineseLegalSimplified'
			} else if (oNumberingLvl.Format == 14) {
				sType = 'decimalEnclosedCircle'
			} else if (oNumberingLvl.Format == 15) {
				sType = 'decimalEnclosedCircleChinese'
			} else if (oNumberingLvl.Format == 21) {
				sType = 'decimalZero'
			} else if (oNumberingLvl.Format == 46) {
				sType = 'lowerLetter'
			} else if (oNumberingLvl.Format == 47) {
				sType = 'lowerRoman'	
			} else if (oNumberingLvl.Format == 60) {
				sType = 'upperLetter'
			} else if (oNumberingLvl.Format == 61) {
				sType = 'upperRoman'
			}
			var bulletStr = vshow ? '\ue6a1 ' : ''
			oNumberingLevel.SetTemplateType('bullet', bulletStr);
			var str = ''
			str += bulletStr
			if (vshow) {
				for (var i = 0; i < LvlText.length; ++i) {
					if (LvlText[i].Type == 2) {
						str += `%${level+1}`
					} else {
						if (LvlText[i].Value != '\ue6a1') {
							str += LvlText[i].Value
						}	
					}
				} 
			} else {
				var find = false
				for (var i = 0; i < LvlText.length; ++i) {
					if (LvlText[i].Type == 2) {
						str += `%${level+1}`
					} else {
						if (LvlText[i].Value == ' ') {
							if (find) {
								str += LvlText[i].Value
							}
						} else if (LvlText[i].Value != '\ue6a1') {
							str += LvlText[i].Value
							find = true
						}
					}
				}
			}
			oNumberingLevel.SetCustomType(sType, str, "left")
			var oTextPr = oNumberingLevel.GetTextPr();
			oTextPr.SetFontFamily("iconfont");
		}

		function getExistDrawing(draws, sub_type_list, write_id, index) {
			var list = []
			for (var i = 0; i < draws.length; ++i) {
				var title = draws[i].Drawing.docPr.title
				if (title) {
					var titleObj = JSON.parse(title)
					if (titleObj.feature) {
						if (titleObj.feature.zone_type == 'question' && sub_type_list.indexOf(titleObj.feature.sub_type) >= 0) {
							if (write_id) {
								if (titleObj.feature.write_id == write_id) {
									list.push(draws[i])
								}
							} else {
								list.push(draws[i])
							}
						}
					}
				}
			}
			return list
		}

		function getAccurateDrawing(index, write_id) {
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(
				3600,
				Api.CreateSolidFill(Api.CreateRGBColor(153, 153, 153))
			)
			var oDrawing = Api.CreateShape(
				'rect',
				6 * 36e3,
				4 * 36e3,
				oFill,
				oStroke
			)
			var drawDocument = oDrawing.GetContent()
			var paragraphs = drawDocument.GetAllParagraphs()
			if (paragraphs && paragraphs.length > 0) {
				var oRun = Api.CreateRun()
				oRun.AddText(index + '')
				oRun.SetFontSize(22)
				oRun.SetVertAlign('baseline')
				paragraphs[0].AddElement(oRun, 0)
				paragraphs[0].SetColor(153, 153, 153, false)
				paragraphs[0].SetJc('center')
				paragraphs[0].SetSpacingAfter(0)
				oDrawing.SetPaddings(0, 0, 0, 0)
			}
			var titleobj = {
				feature: {
					zone_type: 'question',
					type: 'ques_interaction',
					sub_type: 'ask_accurate',
					write_id: write_id
				}
			}
			oDrawing.Drawing.Set_Props({
				title: JSON.stringify(titleobj),
			})
			return oDrawing
		}

		function deleteAccurateRun(oRun) {
			if (oRun &&
				oRun.Run &&
				oRun.Run.Content &&
				oRun.Run.Content[0] &&
				oRun.Run.Content[0].docPr) {
				var title = oRun.Run.Content[0].docPr.title
				if (title) {
					var titleObj = JSON.parse(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						oRun.Delete()
						return true
					}
				}
			}
			return false
		}

		function deleShape(oShape) {
			if (!oShape) {
				return
			}
			var run = oShape.Drawing.GetRun()
			if (run) {
				var runParent = run.GetParent()
				if (runParent) {
					var oParent = Api.LookupObject(runParent.Id)
					if (oParent && oParent.GetClassType() == 'inlineLvlSdt') {
						var count = oParent.GetElementsCount()
						for (var c = 0; c < count; ++c) {
							var child = oParent.GetElement(c)
							if (child.GetClassType() == 'run' && child.Run.Id == run.Id) {
								deleteAccurateRun(child)
								break
							}
						}
						return true
					}
				}
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0) {
						oShape.Delete()
						if (oParagraph) {
							oParagraph.RemoveElement(ipos)
						}
						return true
					}
				}
			}
			oShape.Delete()
			return true
		}

		function addAskInteraction(oControl, askData, index, write_id) {
			var oDrawing = getAccurateDrawing(index, write_id)
			if (askData.sub_type == 'control') {
				var askControl = Api.LookupObject(askData.control_id)
				if(!askControl) {
					return
				}
				if (askControl.GetClassType() == 'inlineLvlSdt') {
					var elementCount = askControl.GetElementsCount()
					if (elementCount > 0) {
						var oRun = Api.CreateRun();
						oRun.AddDrawing(oDrawing)
						askControl.AddElement(oRun, 0)
					}
				} else if (askControl.GetClassType() == 'blockLvlSdt') {
					var paragraphs2 = askControl.GetAllParagraphs()
					if (paragraphs2 && paragraphs2.length > 0) {
						var pParagraph = paragraphs2[0]
						var oRun = Api.CreateRun()
						oRun.AddDrawing(oDrawing)
						pParagraph.AddElement(
							oRun,
							0
						)
						oDrawing.SetWrappingStyle('inline')
					}
				}
			} else if (askData.sub_type == 'write') {
				if (askData.shape_id) {
					var oShape = Api.LookupObject(askData.shape_id)
					if (oShape) {
						var shapeContent = oShape.GetContent()
						if (shapeContent) {
							var paragraphs3 = shapeContent.GetAllParagraphs()
							if (paragraphs3 && paragraphs3.length) {
								var pParagraph = paragraphs3[0]
								pParagraph.SetJc('left')
								var oRun = Api.CreateRun()
								oRun.AddDrawing(oDrawing)
								pParagraph.AddElement(
									oRun,
									0
								)
								oDrawing.SetWrappingStyle('inline')
							}
						}
					}
				}
			} else if (askData.sub_type == 'cell') {
				var oCell = Api.LookupObject(askData.cell_id)
				oCell.SetCellMarginLeft(0)
				var paragraphs3 = oCell.GetContent().GetAllParagraphs()
				if (paragraphs3 && paragraphs3.length) {
					var pParagraph = paragraphs3[0]
					pParagraph.SetJc('left')
					var oRun = Api.CreateRun()
					oRun.AddDrawing(oDrawing)
					pParagraph.AddElement(
						oRun,
						0
					)
					oDrawing.SetVerPosition("paragraph", 0);
					oDrawing.SetHorAlign('column', 'left');
					oDrawing.SetWrappingStyle('inFront')
				}
			} else if (askData.sub_type == 'identify') {
				// todo..
			}
		}

		function handleAccurate(oControl, ask_list, write_list) {
			if (!oControl) {
				return
			}
			if (!write_list || !ask_list) {
				return
			}
			var drawings = oControl.GetAllDrawingObjects()
			for (var i = 0; i < ask_list.length; ++i) {
				var askData = write_list.find(e => {
					return e.id == ask_list[i].id
				})
				if (!askData) {
					continue
				}
				var dlist = getExistDrawing(drawings, ['ask_accurate'], ask_list[i].id)
				if (interaction_type == 'accurate') {
					if (!dlist || dlist.length == 0) {
						addAskInteraction(oControl, askData, i + 1, ask_list[i].id)
					} else {
						var content = dlist[0].Drawing.GraphicObj.textBoxContent // shapeContent
						if (content && content.Content && content.Content.length) {
							var paragraph = content.Content[0]
							if (paragraph && paragraph.GetElementsCount()) {
								var run = paragraph.GetElement(0)
								if (run) {
									if (run.GetText() * 1 != i + 1) { // 序号run
										paragraph.ReplaceCurrentWord(0, `${i+1}`)
									}
								}
							}
						}
						

					}
				} else {
					for (var j = 0; j < dlist.length; ++j) {
						deleShape(dlist[j])
					}
				}
			}
		}

		function getFirstParagraph(oControl) {
			if (!oControl) {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
			for (var i = 0; i < paragraphs.length; ++i) {
				var oParagraph = paragraphs[i]
				if (oParagraph) {
					var parent1 = oParagraph.Paragraph.Parent
					var parent2 = parent1.Parent
					if (parent2 && parent2.Id == oControl.Sdt.GetId()) {
						return oParagraph
					}
				}
			}
			return null
		}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tag = JSON.parse(oControl.GetTag() || '{}')
			if (quesIds) {
				var qindex = quesIds.findIndex(e => {
					return e == tag.client_id
				})
				if (qindex == -1) {
					continue
				}
			}
			if (tag.regionType == 'question') {
				if (interaction_type != 'none') {
					if (!question_map[tag.client_id] || question_map[tag.client_id].level_type != 'question') {
						continue
					}
				}
				var ask_list = question_map[tag.client_id].ask_list
				var nodeData = node_list.find(e => {
					return e.id == tag.client_id
				})
				var write_list = nodeData ? (nodeData.write_list || []) : []
				// var allDraws = oControl.GetAllDrawingObjects()
				// var simpleDrawings = getExistDrawing(allDraws, ['simple'])
				// var accurateDrawings = getExistDrawing(allDraws, ['accurate', 'ask_accurate'])
				var firstParagraph = getFirstParagraph(oControl)
				if (firstParagraph) {
					showSimple(firstParagraph, interaction_type == 'simple' || interaction_type == 'accurate')
				}
				handleAccurate(oControl, ask_list, write_list)
			}
		}

	}, false, true).then(res => {
		console.log('setInteraction result:', res)
	})
}

export { handleFeature, handleHeader, drawExtroInfo, setLoading, deleteAllFeatures, setInteraction }
