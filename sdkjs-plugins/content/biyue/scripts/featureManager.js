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
	if (!options) { return }
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

function drawExtroInfo(list, imageDimensionsCache, calc) {
	if (!list) { return }
	list.forEach(e => {
		e.page_num = e.p || 0
		if (e.v == undefined) {
			e.v = 1
		}
		e.size = Object.assign({}, ZONE_SIZE[e.zone_type], (e.size || {}))
		e.type = 'feature'
		e.type_name = ZONE_TYPE_NAME[e.zone_type]
	})
	Asc.scope.imageDimensionsCache = imageDimensionsCache
	return drawList(list, calc)
}
// 整理参数
function addCommand(options) {
	if (!options) { return }
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
	if (list_wait_command && list_wait_command.length) {
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
// 在页眉绘制再练和二维码		目前弃用
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
// 删除所有功能区
function deleteAllFeatures(exceptList, specifyFeatures) {
	Asc.scope.exceptList = exceptList
	Asc.scope.specifyFeatures = specifyFeatures
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects()
		var exceptList = Asc.scope.exceptList
		var specifyFeatures = Asc.scope.specifyFeatures
		function deleteAccurate(oDrawing) {
			var paraDrawing = oDrawing.getParaDrawing()
			var run = paraDrawing ? paraDrawing.GetRun() : null
			if (run) {
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0 && oParagraph.GetClassType() == 'paragraph') {
						oDrawing.Delete()
						var element2 = oParagraph.GetElement(ipos)
						if (element2 && element2.GetClassType() == 'run' && element2.Run.Id == run.Id) {
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
			// titleObj.feature.zone_type == 'pagination' ||
			titleObj.feature.zone_type == 'statistics') {
				deleteAccurate(oDrawing)
			} else {
				oDrawing.Delete()
			}
		}
		if (drawings) {
			for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
				var oDrawing = drawings[j]
				var title = oDrawing.GetTitle()
				if (title && title.indexOf('feature') >= 0) {
					var titleObj = Api.ParseJSON(title)
					if (titleObj.feature && titleObj.feature.zone_type) {
						// 作答区和识别框不应该被删除
						if (titleObj.feature.sub_type == 'write' || titleObj.feature.sub_type == 'identify' || titleObj.feature.type == 'score') {
							continue
						}
						if (exceptList) {
							var inExcept = exceptList.findIndex(e => {
								return e == titleObj.feature.zone_type
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
		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
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
				if (LvlText[0].Value!='\ue749') {
					var targetInd = oParagraph.GetParentTableCell() ? 280 : 0
					oParagraph.SetIndFirstLine(targetInd)
					return
				}
			}
			var key = `${oNum.Id}_${level}`
			if (handledNumbering[key]) {
				var targetInd = oParagraph.GetParentTableCell() ? 280 : 0
				oParagraph.SetIndFirstLine(targetInd)
				return
			}
			handledNumbering[key] = 1
			var suffix = ''
			if (LvlText.length > 1 && LvlText[LvlText.length - 1].Type == 1) {
				suffix = LvlText[LvlText.length - 1].Value
			}
			var sType = Api.GetFormatTypeString(oNumberingLvl.Format)
			oNumberingLevel.SetTemplateType('bullet', '');
			var str = ''
			var find = false
			for (var i = 0; i < LvlText.length; ++i) {
				if (LvlText[i].Type == 2) {
					if (LvlText[i].Value == level) {
						str += `%${level+1}`
					}
				} else {
					if (LvlText[i].Value == ' ') {
						if (find) {
							str += LvlText[i].Value
						}
					} else if (LvlText[i].Value != '\ue749' && LvlText[i].Value != '\ue607') {
						str += LvlText[i].Value
						find = true
					}
				}
			}
			oNumberingLevel.SetCustomType(sType, str, "left")
			var oTextPr = oNumberingLevel.GetTextPr();
			oTextPr.SetFontFamily("iconfont");
			var targetInd = oParagraph.GetParentTableCell() ? 280 : 0
			oParagraph.SetIndFirstLine(targetInd)
			return true
		}
		var controls = oDocument.GetAllContentControls()
		if (controls) {
			for (var j = 0, jmax = controls.length; j < jmax; ++j) {
				var oControl = controls[j]
				if (oControl.GetClassType() == 'blockLvlSdt') {
					var childControls = oControl.GetAllContentControls() || []
					if (childControls.length) {
						for (var c = childControls.length - 1; c >= 0; --c) {
							var tag = Api.ParseJSON(childControls[c].GetTag())
							if (tag.regionType == 'num' && childControls[c].GetClassType() == 'inlineLvlSdt') {
								var parent = childControls[c].Sdt.Parent
								if (parent && parent.GetType() == 1) {
									var oParent = Api.LookupObject(parent.Id)
									if (oParent) {
										var pos = childControls[c].Sdt.GetPosInParent()
										if (pos >= 0) {
											oParent.RemoveElement(pos)
										}
									}
								}
							}
						}
					}
					var firstParagraph = getFirstParagraph(oControl)
					hideSimple(firstParagraph)
				}
			}
		}
	}, false, true)
}

function drawList(list, recalc = true) {
	setLoading(true)
	Asc.scope.feature_wait_handle = list
	Asc.scope.ZONE_TYPE = ZONE_TYPE
	Asc.scope.page_type = window.BiyueCustomData.page_type
	return biyueCallCommand(window, function() {
		var imageDimensionsCache = Asc.scope.imageDimensionsCache || {}
		var ZONE_TYPE = Asc.scope.ZONE_TYPE
		var MM2TWIPS = 25.4 / 72 / 20
		var oDocument = Api.GetDocument()
		var objs = oDocument.GetAllDrawingObjects()
		var oSections = oDocument.GetSections()
		var feature_wait_handle = Asc.scope.feature_wait_handle
		var page_type = Asc.scope.page_type
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
					var dimension = imageDimensionsCache[url]
					var h2 = h
					if (dimension && dimension.aspectRatio != undefined && dimension.aspectRatio != 0) {
						h2 = w / dimension.aspectRatio
					}
					var oImage = Api.CreateImage(url, w * 36e3, h2 * 36e3)
					if (oImage) {
						oImage.SetTitle(JSON.stringify({
							ignore: 1
						}))
					}
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
			oDrawing.SetTitle(JSON.stringify(titleobj))
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
							var title = oRun.Run.Content[j].docPr ? (oRun.Run.Content[j].docPr.title || '{}') : ''
							var titleObj = Api.ParseJSON(title)
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
			var paraDrawing = oDrawing.getParaDrawing()
			var run = paraDrawing ? paraDrawing.GetRun() : null
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
		for (var options of feature_wait_handle) {
			if (options.type_name == 'statistics') {
				continue
			}
			var props_title = JSON.stringify({
				feature: {
					zone_type: options.type_name,
					v: options.v
				}
			})
			var find = objs.find(e => {
				return e.GetTitle() == props_title
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
							return e.GetTitle() == props_title
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
							return e.GetTitle() == props_title2
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
									if (oImage) {
										oImage.SetTitle(JSON.stringify({
											ignore: 1
										}))
									}
									p.AddDrawing(oImage)
									p.SetSpacingAfter(0)
								}
							}
						} else if (options.zone_type == ZONE_TYPE.THER_EVALUATION || options.zone_type == ZONE_TYPE.SELF_EVALUATION) {
							var flowers = options.flowers || []
							var oTable = Api.CreateTable(2 + flowers.length, 1)
							oTable.SetTableTitle(JSON.stringify({ignore: 1}))
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
								if (oImage) {
									oImage.SetTitle(JSON.stringify({
										ignore: 1
									}))
								}
								p.AddDrawing(oImage)
							}
						} else if (options.label) {
							var oFill = Api.CreateNoFill()
							var stroke_width = options.size.stroke_width || 0.1
							var oStroke = Api.CreateStroke(
								(options.zone_type == ZONE_TYPE.AGAIN ? 0 : stroke_width) * 36e3,
								Api.CreateSolidFill(Api.CreateRGBColor(53, 53, 53))
							)
							oDrawing = Api.CreateShape(
								options.size.shape_type || 'rect',
								shapeWidth * 36e3,
								shapeHeight * 36e3,
								oFill,
								oStroke
							)
							if (options.zone_type == ZONE_TYPE.AGAIN) {
								var outLineStroke = Api.CreateStroke(stroke_width * 2 * 36e3, Api.CreateSolidFill(Api.CreateRGBColor(53, 53, 53)))
								outLineStroke.Ln.setPrstDash(0)
								oDrawing.SetOutLine(outLineStroke)
							}
							var drawDocument = oDrawing.GetContent()
							var paragraphs = drawDocument.GetAllParagraphs()
							if (paragraphs && paragraphs.length > 0) {
								var oRun = Api.CreateRun()
								oRun.AddText(options.label)
								paragraphs[0].AddElement(oRun)
								paragraphs[0].SetColor(53, 53, 53, false)
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
							oDrawing.SetTitle(props_title)
							if (options.zone_type == ZONE_TYPE.STATISTICS) {
								if (!options.size.pagination) {
									options.size.pagination = {
										font_size: 2.71,
										align_style: 'center',
										margin: 0,
										bottom: 0
									}
								}
								var numberDrawing = getPageNumberDrawing(20, options.size.pagination.font_size)
								var pstyle = options.size.pagination.align_style
								oDocument.SetEvenAndOddHdrFtr(pstyle != 'center');
								oSections.forEach((section, index) => {
									var footerList = []
									var oTitleFooter = section.GetFooter("title", false)
									// 首页
									if (section.Section.IsTitlePage()) {
										if (!oTitleFooter) {
											oTitleFooter = section.GetFooter("title", true)
										}
									}
									if (oTitleFooter) {
										footerList.push({
											type: 'title',
											oFooter: oTitleFooter
										})
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
										if (page_type == 0) {
											// 统计
											var sh = PageMargins.Right - (options.size.right || 0) - options.size.w
											var sv = PageMargins.Bottom - (options.size.bottom || 0) - options.size.h
											var typeDrawing = getTypeDrawing(oParagraph, options.type_name)
											if (typeDrawing) {
												typeDrawing.Set_PositionH(7, false, sh, false);
												typeDrawing.Set_PositionV(0, false, sv, false)
												typeDrawing.Set_DrawingType(2);
											} else {
												var oAdd = oDrawing.Copy()
												var drawing = oAdd.getParaDrawing()
												if (drawing) {
													drawing.Set_PositionH(7, false, sh, false);
													drawing.Set_PositionV(0, false, sv, false)
													drawing.Set_DrawingType(2);
													oParagraph.SetJc('center')
													var titleobj = {
														feature: {
															zone_type: options.type_name,
															v: options.v
														}
													}
													oAdd.SetTitle(JSON.stringify(titleobj))
													oParagraph.AddDrawing(oAdd)
													drawing.Set_Parent(oParagraph.Paragraph)
												}
											}
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
													v: options.v,
													footer_type: footerObj.type
												}
											}
											oAddNum.SetTitle(JSON.stringify(titleobj))
											var align = 'center'
											if (pstyle == 'oddLeftEvenRight') {
												align = footerObj.type == 'even' ? 'right' : 'left'
											} else if (pstyle == 'oddRightEvenLeft') {
												align = footerObj.type == 'even' ? 'left' : 'right'
											}
											var pmargin = options.size.pagination ? options.size.pagination.margin : 0
											var pbottom = options.size.pagination ? options.size.pagination.bottom : 0
											var psize = options.size.pagination ? options.size.pagination.font_size : 2.71
											setPaginationAlign(oAddNum, align, PageMargins, pmargin)
											oAddNum.SetVerPosition('bottomMargin', (PageMargins.Bottom -pbottom - psize) * 36e3)
											if (needAdd) {
												oParagraph.AddDrawing(oAddNum)
												if (oAddNum.getParaDrawing()) {
													oAddNum.getParaDrawing().Set_Parent(oParagraph.Paragraph)
												}
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
									var drawing = oDrawing.getParaDrawing();
										drawing.Set_PositionH(6, false, options.x, false);
										drawing.Set_PositionV(5, false, options.y, false);
										drawing.Set_DrawingType(2);
										if (lastParagraph && paragraph.Id == lastParagraph.Paragraph.Id) {
											// lastParagraph.Paragraph.AddToParagraph(drawing)
											lastParagraph.AddDrawing(oDrawing)
											drawing.Set_Parent(lastParagraph.Paragraph)
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
		}
		console.log('=====================drawList end ')
		return res
	}, false, false)
}

function setLoading(v) {
	loading = v
}

function setInteraction(type, quesIds, recalc = true) {
	Asc.scope.interaction_type_use = type
	Asc.scope.interaction_quesIds = quesIds
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.node_list = window.BiyueCustomData.node_list
	return biyueCallCommand(window, function() {
		var interaction_type_use = Asc.scope.interaction_type_use
		var interaction_type = interaction_type_use
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var question_map = Asc.scope.question_map || {}
		var node_list = Asc.scope.node_list || []
		var MM2TWIPS = 25.4 / 72 / 20
		var quesIds = Asc.scope.interaction_quesIds
		var allParagraphs = oDocument.GetAllParagraphs()
		if (!controls) {
			return
		}
		var handledNumbering = {}
		var SIMPLE_CHAR = '\ue749'
		var vInd = 280
		
		function updateParagraphInd(oParagraph, vshow) {
			var targetInd = 0
			if (oParagraph.GetParentTableCell()) {
				targetInd = vshow ? 0 : vInd
			} else {
				targetInd = vshow ? (0 - vInd) : 0
			}
			oParagraph.SetIndFirstLine(targetInd)
		}
		function syncSameParagraph(numbering, oParagraph, vshow) {
			allParagraphs.forEach(e => {
				if (e.Paragraph.Id != oParagraph.Paragraph.Id) {
					var n = e.GetNumbering()
					if (n) {
						if (n.Lvl == numbering.Lvl && n.Num && n.Num.Id == numbering.Num.Id) {
							updateParagraphInd(e, vshow)
						}
					}
				}
			})
		}
		function showControlSimple(oParagraph, vshow) {
			if (!oParagraph) {
				return
			}
			// 删除已有的
			var childControls = oParagraph.GetAllContentControls() || []
			if (childControls.length) {
				for (var c = childControls.length - 1; c >= 0; --c) {
					var tag = Api.ParseJSON(childControls[c].GetTag())
					if (tag.regionType == 'num' && childControls[c].GetClassType() == 'inlineLvlSdt') {
						var parent = childControls[c].Sdt.Parent
						if (parent && parent.GetType() == 1) {
							var oParent = Api.LookupObject(parent.Id)
							if (oParent) {
								var pos = childControls[c].Sdt.GetPosInParent()
								if (pos >= 0) {
									oParent.RemoveElement(pos)
								}
							}
						}
					}
				}
			}
			// 添加
			if (vshow) {
				// 删除原有的"\ue749"，多个时暂不处理
				var elementcount = oParagraph.GetElementsCount()
				if (elementcount > 0) {
					for (var i = 0; i < elementcount; ++i) {
						var oChild = oParagraph.GetElement(i)
						if (oChild.GetClassType() == 'run') {
							if (oChild.Run.IsEmpty()) {
								continue
							}
							var find = false
							for (var j = 0; j < oChild.Run.GetElementsCount(); ++j) {
								var oElement = oChild.Run.GetElement(j)
								if (oElement.GetType() == 1) {
									if (oElement.Value == 59209) {
										oChild.Run.RemoveElement(oElement)
										find = true
										break
									}
								}
							}
							if (find) {
								break
							}
						} else {
							break
						}
					}
				}
				var oInlineLvlSdt = Api.CreateInlineLvlSdt();
				var oRun = Api.CreateRun();
				oRun.SetFontFamily('iconfont')
				oRun.AddText("\ue749")
				oInlineLvlSdt.SetTag(JSON.stringify({
					regionType: 'num',
					color: '#ffffff40'
				}))
				oInlineLvlSdt.AddElement(oRun, 0);
				oParagraph.Paragraph.Add_ToContent(0, oInlineLvlSdt.Sdt)
			}
		}
		function showSimple(oParagraph, vshow) {
			if (!oParagraph) {
				return
			}
			var oNumberingLevel = oParagraph.GetNumbering()
			if (!oNumberingLevel) {
				showControlSimple(oParagraph, vshow)
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
			var indleft = vshow ? (0 - vInd) : 0
			var LvlText = oNumberingLvl.LvlText || []
			if (LvlText && LvlText.length) {
				if (LvlText[0].Value==SIMPLE_CHAR) {
					if (vshow) {
						//console.log('当前简单互动已显示', oParagraph.GetText(), indleft)
						updateParagraphInd(oParagraph, vshow)
						return
					}
				} else {
					if (!vshow) {
						// console.log('当前简单互动本就未显示', oParagraph.GetText())
						updateParagraphInd(oParagraph, vshow)
						return
					}
				}
			}
			var key = `${oNum.Id}_${level}`
			if (handledNumbering[key]) {
				updateParagraphInd(oParagraph, vshow)
				return
			}
			handledNumbering[key] = 1
			var sType = Api.GetFormatTypeString(oNumberingLvl.Format)
			var bulletStr = vshow ? `${SIMPLE_CHAR} ` : ''
			oNumberingLevel.SetTemplateType('bullet', bulletStr);
			var str = ''
			str += bulletStr
			if (vshow) {
				for (var i = 0; i < LvlText.length; ++i) {
					if (LvlText[i].Type == 2) {
						if (LvlText[i].Value == level) {
							str += `%${level+1}`
						}
					} else {
						if (LvlText[i].Value != '\ue749') {
							str += LvlText[i].Value
						}	
					}
				} 
			} else {
				var find = false
				for (var i = 0; i < LvlText.length; ++i) {
					if (LvlText[i].Type == 2) {
						if (LvlText[i].Value == level) {
							str += `%${level+1}`
						}
					} else {
						if (LvlText[i].Value == ' ') {
							if (find) {
								str += LvlText[i].Value
							}
						} else if (LvlText[i].Value != '\ue749') {
							str += LvlText[i].Value
							find = true
						}
					}
				}
			}
			oNumberingLevel.SetCustomType(sType, str, "left")

			var oTextPr = oNumberingLevel.GetTextPr();
			oTextPr.SetFontFamily("iconfont");
			updateParagraphInd(oParagraph, vshow)
			var numbering = oParagraph.GetNumbering()
			syncSameParagraph(numbering, oParagraph, vshow)
		}

		function getExistDrawing(draws, sub_type_list, write_id) {
			var list = []
			for (var i = 0; i < draws.length; ++i) {
				var titleObj = Api.ParseJSON(draws[i].GetTitle())
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
			return list
		}

		function getAccurateDrawing(index, write_id, width = 6, height = 4) {
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(
				3600,
				Api.CreateSolidFill(Api.CreateRGBColor(153, 153, 153))
			)
			var oDrawing = Api.CreateShape(
				'rect',
				width * 36e3,
				height * 36e3,
				oFill,
				oStroke
			)
			var drawDocument = oDrawing.GetContent()
			var paragraphs = drawDocument.GetAllParagraphs()
			if (paragraphs && paragraphs.length > 0) {
				if (index > 0) {
					var oRun = Api.CreateRun()
					oRun.AddText(index + '')
					oRun.SetFontSize(16)
					oRun.SetVertAlign('baseline')
					paragraphs[0].AddElement(oRun, 0)
					paragraphs[0].SetColor(153, 153, 153, false)
				}
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
			oDrawing.SetTitle(JSON.stringify(titleobj))
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
					var titleObj = Api.ParseJSON(title)
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
			var paraDrawing = oShape.getParaDrawing()
			var run = paraDrawing ? paraDrawing.GetRun() : null
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

		function addCellInteraction(cell_id, oDrawing) {
			var oCell = Api.LookupObject(cell_id)
			if (!oCell || oCell.GetClassType() != 'tableCell') {
				return
			}
			var oTable = oCell.GetParentTable()
			if (oTable.GetPosInParent() == -1) {
				return
			}
			oCell.SetCellMarginLeft(0)
			oCell.SetVerticalAlign('top')
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
		}

		function getControl(client_id, regionType) {
			return controls.find(e => {
				var tag = Api.ParseJSON(e.GetTag())
				if ((e.GetClassType() == 'blockLvlSdt' && e.GetPosInParent() >= 0) || (e.GetClassType() == 'inlineLvlSdt' && e.Sdt.GetPosInParent() >= 0)) {
					return tag.client_id == client_id && tag.regionType == regionType
				}
			})
		}

		function addAskInteraction(oControl, askData, index, write_id) {
			var oDrawing = getAccurateDrawing(index, write_id)
			if (askData.sub_type == 'control') {
				var askControl = getControl(askData.id, 'write') // Api.LookupObject(askData.control_id)
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
				var drawings = oControl.GetAllDrawingObjects() || []
				var oAskDrawing = drawings.find(e => {
					var drawingTitle = Api.ParseJSON(e.GetTitle())
					return drawingTitle.feature && drawingTitle.feature.client_id == askData.id
				})
				if (oAskDrawing) {
					var oShape = Api.LookupObject(oAskDrawing.Drawing.Id)
					if (oShape && oShape.GetClassType() == 'shape') {
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
								oShape.SetVerticalTextAlign("top")
								oDrawing.SetVerPosition("paragraph", 0);
								oDrawing.SetHorAlign('column', 'left');
								oDrawing.SetWrappingStyle('inFront')
							}
						}
					}
				}
			} else if (askData.sub_type == 'cell') {
				addCellInteraction(askData.cell_id, oDrawing)
			} else if (askData.sub_type == 'identify') {
				// todo..
			}
		}

		function handleControlAccurate(oControl, ask_list, write_list, type) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			var drawings = oControl.GetAllDrawingObjects()
			var dlist1 = getExistDrawing(drawings, ['ask_accurate'])
			if (type == 'none' || type == 'simple') {
				if (dlist1 && dlist1.length) {
					dlist1.forEach(e => {
						deleShape(e)
					})
				}
				return
			}
			if (!write_list || !ask_list) {
				return
			}
			for (var i = 0; i < ask_list.length; ++i) {
				var askData = write_list.find(e => {
					return e.id == ask_list[i].id
				})
				if (!askData) {
					continue
				}
				var dlist = getExistDrawing(drawings, ['ask_accurate'], ask_list[i].id)
				if (type == 'accurate') {
					if (!dlist || dlist.length == 0) {
						addAskInteraction(oControl, askData, i + 1, ask_list[i].id)
					} else {
						var content = dlist[0].Drawing.textBoxContent // shapeContent
						if (content && content.Content && content.Content.length) {
							var paragraph = content.Content[0]
							if (paragraph) {
								if (paragraph.GetElementsCount()) {
									var run = paragraph.GetElement(0)
									if (run) {
										if (run.GetText() * 1 != i + 1) { // 序号run
											paragraph.ReplaceCurrentWord(0, `${i+1}`)
										}
									}
								} else {
									var oPara = Api.LookupObject(paragraph.Id)
									oPara.AddText(`${i+1}`)
									oPara.SetColor(153, 153, 153, false)
									oPara.SetJc('center')
									oPara.SetSpacingAfter(0)
								}
							}
						}
						var fidx = dlist1.findIndex(e => {
							return e.Drawing.Id == dlist[0].Drawing.Id
						})
						if (fidx >= 0) {
							dlist1.splice(fidx, 1)
						}
					}
				} else {
					for (var j = 0; j < dlist.length; ++j) {
						deleShape(dlist[j])
						var fidx = dlist1.findIndex(e => {
							return e.Drawing.Id == dlist[j].Drawing.Id
						})
						if (fidx >= 0) {
							dlist1.splice(fidx, 1)
						}
					}
				}
			}
			dlist1.forEach(e => {
				deleShape(e)
			})
		}

		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
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
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			var targetQuesId = tag.mid ? tag.mid : tag.client_id
			if (quesIds) {
				var qindex = quesIds.findIndex(e => {
					return e == targetQuesId
				})
				if (qindex == -1) {
					continue
				}
			}
			if (tag.regionType != 'question') {
				continue
			}
			if (!question_map[targetQuesId]) {
				continue
			}
			if (interaction_type_use != 'none') {
				if (!question_map[targetQuesId] || question_map[targetQuesId].level_type != 'question') {
					continue
				}
				if (question_map[targetQuesId].mark_mode == 2) {
					if (interaction_type_use == 'accurate') {
						interaction_type = 'simple'
					}
				} else {
					if (interaction_type_use == 'useself') {
						interaction_type = question_map[targetQuesId].interaction
					}
				}
			}
			if (question_map[targetQuesId].ques_mode == 6) {
				interaction_type = 'none'
			}
			var ask_list = question_map[targetQuesId].ask_list
			var nodeData = node_list.find(e => {
				return e.id == tag.client_id
			})
			var write_list = nodeData ? (nodeData.write_list || []) : []
			// var allDraws = oControl.GetAllDrawingObjects()
			// var simpleDrawings = getExistDrawing(allDraws, ['simple'])
			// var accurateDrawings = getExistDrawing(allDraws, ['accurate', 'ask_accurate'])
			var isGatherChoice = (question_map[targetQuesId].ques_mode == 1 || question_map[targetQuesId].ques_mode == 5) && nodeData.use_gather
			var type = isGatherChoice ? 'none' : interaction_type
			var firstParagraph = getFirstParagraph(oControl)
			if (firstParagraph) {
				showSimple(firstParagraph, type != 'none')
			}
			if (isGatherChoice && nodeData.gather_cell_id) {
				// 在集中作答区添加互动
				var oCell = Api.LookupObject(nodeData.gather_cell_id)
				if (oCell && oCell.GetClassType() == 'tableCell') {
					var drawings = oCell.GetContent().GetAllDrawingObjects()
					var dlist = getExistDrawing(drawings, ['ask_accurate'], nodeData.id)
					if (!dlist || dlist.length == 0) {
						if (interaction_type != 'none') {
							var oDrawing = getAccurateDrawing(0, nodeData.id, 4, 2.5)
							addCellInteraction(nodeData.gather_cell_id, oDrawing)
						}
					} else if (interaction_type == 'none') {
						deleShape(dlist[0])
					}
				}
			}
			handleControlAccurate(oControl, ask_list, write_list, type)
		}

	}, false, recalc).then(res => {
		console.log('setInteraction result:', res)
	})
}

function updateChoice(recalc = true) {
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.choice_params = window.BiyueCustomData.choice_display
	console.log('Asc.scope.choice_params', Asc.scope.choice_params)
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list
		var oDocument = Api.GetDocument()
		var oTables = oDocument.GetAllTables() || []
		for (var i1 = oTables.length - 1; i1 >= 0; --i1) {
			var tableTitle = oTables[i1].GetTableTitle()
			if (tableTitle.indexOf('choiceGather') >= 0) {
				var pos = oTables[i1].GetPosInParent()
				var parent = oTables[i1].Table.GetParent()
				if (parent) {
					var oParent = Api.LookupObject(parent.Id)
					oParent.RemoveElement(pos)
				}
			}
		}
		var choice_params = Asc.scope.choice_params
		if (choice_params.style == 'brackets_choice_region') {
			node_list.forEach(e => {
				e.use_gather = false
			})
			return node_list
		}
		var question_map = Asc.scope.question_map
		var mark_type_info = Asc.scope.subject_mark_types
		var num_row = choice_params.num_row || 10
		var structs = []
		function AddItem(id, control_id) {
			if (question_map[id] && question_map[id].question_type > 0) {
				var ques_mode = 0
				if (mark_type_info && mark_type_info.list) {
					var find = mark_type_info.list.find(e => {
						return e.question_type_id == question_map[id].question_type
					})
					ques_mode = find ? find.ques_mode : 3
				}
				// todo..1是单选题，5是多选题，之后要考虑多选另行一组
				if (ques_mode == 1 || ques_mode == 5) {
					structs[structs.length - 1].items.push({
						id: id,
						control_id: control_id,
						name: question_map[id].ques_name || question_map[id].ques_default_name || ''
					})
				}
			}
		}
		function handleControl(oControl, i) {
			if (!oControl) {
				return
			}
			var tag = Api.ParseJSON(oControl.GetTag())
			if (!tag.client_id) {
				control_id
			}
			var nodeData = node_list.find(e => {
				return e.id == tag.client_id
			})
			if (!nodeData) {
				return
			}
			if (nodeData.level_type == 'struct') {
				if (structs.length) {
					structs[structs.length - 1].last_pos = i - 1
				}
				structs.push({
					struct_id: nodeData.id,
					control_id: oControl.Sdt.GetId(),
					pos: i,
					last_pos: i,
					items: []
				})
			} else {
				// 没有题组时，自己就是题组
				if (structs.length == 0) {
					structs.push({
						struct_id: 0,
						control_id: oControl.Sdt.GetId(),
						pos: i - 1,
						last_pos: i,
						items: []
					})
				}
				if (nodeData.level_type == 'question') {
					AddItem(nodeData.id, oControl.Sdt.GetId())
					var childControls = oControl.GetAllContentControls()
					if (childControls && childControls.length) {
						childControls.forEach((oChildControl) => {
							if (oChildControl.GetClassType() == 'blockLvlSdt') {
								var childtag = Api.ParseJSON(oChildControl.GetTag())
								if (childtag.client_id) {
									var nodeData2 = node_list.find(e => {
										return e.id == childtag.client_id
									})
									if (nodeData2 && nodeData2.level_type == 'question') {
										AddItem(childtag.client_id, oControl.Sdt.GetId())
									}
								}
							}
						})
					}
					structs[structs.length - 1].last_pos = i
				}
			}
		}
		var elementcount = oDocument.GetElementsCount()
		for (var i = 0; i < elementcount; ++i) {
			var oElement = oDocument.GetElement(i)
			if (oElement.GetClassType() == 'table') {
				var rows = oElement.GetRowsCount()
				for (var i1 = 0; i1 < rows; ++i1) {
					var oRow = oElement.GetRow(i1)
					var cells = oRow.GetCellsCount()
					for (var i2 = 0; i2 < cells; ++i2) {
						var oCell = oRow.GetCell(i2)
						var oCellContent = oCell.GetContent()
						var cnt1 = oCellContent.GetElementsCount()
						for (var i3 = 0; i3 < cnt1; ++i3) {
							var oElement2 = oCellContent.GetElement(i3)
							if (!oElement2) {
								continue
							}
							if (oElement2.GetClassType() == 'blockLvlSdt') {
								handleControl(oElement2, i)
							}
						}
					}
				}
			} else if (oElement.GetClassType() == 'blockLvlSdt') {
				handleControl(oElement, i)
			}
		}
		if (structs.length) {
			structs[structs.length - 1].last_pos = elementcount - 1
			var lastelement = oDocument.GetElement(elementcount - 1)
			if (lastelement.GetClassType() == 'paragraph') { // 有段落时，最后一行不算
				var text = lastelement.GetText()
				if (lastelement.GetElementsCount() == 0 || text == '' || text == '\r\n') {
					structs[structs.length - 1].last_pos = elementcount - 2
				}
			}
		}
		for (var s = structs.length-1; s >= 0; --s) {
			var queslist = structs[s].items
			if (!queslist || queslist.length == 0) {
				continue
			}
			var cellnum = num_row
			var rows = Math.ceil(queslist.length / cellnum)
			rows = rows * 2
			var oTable = null
			var titleObj = {
				choiceGather: {
					struct_id: structs[s].struct_id,
				},
				items: []
			}

			var oTable = Api.CreateTable(cellnum, rows)
			var oTableStyle = oDocument.CreateStyle('CustomTableStyle', 'table')
			oTable.SetStyle(oTableStyle)
			oTable.SetWidth('percent', 100)
			for (var i = 0; i < rows; ++i) {
				var oRow = oTable.GetRow(i)
				oRow.SetHeight("atLeast", i % 2 == 0 ? 360 : 720);
				var rowno = Math.floor(i / 2)
				var mergeCells = []
				for (var j = 0; j < cellnum; ++j) {
					var oCell = oRow.GetCell(j)
					oCell.SetWidth('percent', 100 / cellnum)
					oCell.SetCellBorderBottom("single", 1, 0, 53, 53, 53)
					oCell.SetCellBorderRight("single", 1, 0, 53, 153, 53)
					oCell.SetCellBorderTop("single", 1, 0, 53, 53, 53)
					oCell.SetCellBorderLeft("single", 1, 0, 53, 53, 53)
					var oCellContent = oCell.GetContent()
					if (rowno * cellnum + j >= queslist.length) {
						mergeCells.push(oCell)
						continue
					}
					if (i % 2 == 0) {
						if (queslist[rowno * cellnum + j] && queslist[rowno * cellnum + j].id && question_map[queslist[rowno * cellnum + j].id]) {
							var oParagraph = oCellContent.GetElement(0)
							if (oParagraph && oParagraph.GetClassType() == 'paragraph') {
								oParagraph.AddText(queslist[rowno * cellnum + j].name)
								var nodeData = node_list.find(e => {
									return e.id == queslist[rowno * cellnum + j].id
								})
								if (nodeData) {
									nodeData.use_gather = true
									var oCell2 = oTable.GetCell(i + 1, j)
									if (oCell2) {
										nodeData.gather_cell_id = oCell2.Cell.Id // 用于上传坐标时反向追溯
										titleObj.items.push({
											ques_id: nodeData.id,
											cell_id: oCell2.Cell.Id
										})
									}
								}
								oParagraph.SetJc('center')
								oParagraph.SetColor(0, 0, 0, false)
								oParagraph.SetFontSize(16)
							}
						} else {
							console.log(queslist[rowno * cellnum + j], question_map)
						}
					} else {
						// oCell.SetBackgroundColor(255, 191, 191, false)
					}
				}
				if (mergeCells.length) {
					oTable.MergeCells(mergeCells)
				}
			}
			oTable.SetTableTitle(JSON.stringify(titleObj))
			if (choice_params.area) {
				oDocument.AddElement(structs[s].last_pos + 1, oTable)
			} else {
				oDocument.AddElement(structs[s].pos + 1, oTable)
			}
		}
		return node_list
	}, false, recalc)
}

function handleChoiceUpdateResult(res) {
	if (res) {
		window.BiyueCustomData.node_list = res
		if (window.BiyueCustomData.interaction != 'none') {
			return setInteraction('useself')
		}
	}
	return new Promise((resolve, reject) => {
		resolve()
	})
}
// 显示或隐藏页码
function showOrHidePagination(v) {
	Asc.scope.vshow = v
	return biyueCallCommand(window, function(){
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		var vshow = Asc.scope.v
		drawings.forEach(e => {
			var titleObj = Api.ParseJSON(e.GetTitle())
			if (titleObj.feature && titleObj.feature.zone_type == 'pagination') {
				var oShape = Api.LookupObject(e.Drawing.Id)
				if (oShape && oShape.GetClassType() == 'shape') {
					var oShapeContent = oShape.GetContent()
					if (oShapeContent) {
						var paragraphs = oShapeContent.GetAllParagraphs() || []
						paragraphs.forEach(p => {
							p.SetColor(3, 3, 3, !vshow)
						})
					}
				}
			}
		})
	}, false, false)
}
// 单独更新统计图标
function drawStatistics(options, recalc) {
	Asc.scope.options = options
	return biyueCallCommand(window, function() {
		var options = Asc.scope.options || {}
		var oDocument = Api.GetDocument()
		var oSections = oDocument.GetSections()
		if (!oSections) {
			return
		}
		function deleteAccurate(oDrawing) {
			var paraDrawing = oDrawing.getParaDrawing()
			var run = paraDrawing ? paraDrawing.GetRun() : null
			if (run) {
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0 && oParagraph.GetClassType() == 'paragraph') {
						oDrawing.Delete()
						var element2 = oParagraph.GetElement(ipos)
						if (element2 && element2.GetClassType() == 'run' && element2.Run.Id == run.Id) {
							oParagraph.RemoveElement(ipos)
						}
						return
					}
				}
			}
			oDrawing.Delete()
		}
		if (options.cmd == 'close') {
			var drawings = oDocument.GetAllDrawingObjects() || []
			for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
				var oDrawing = drawings[j]
				var title = oDrawing.GetTitle()
				if (title && title.indexOf('feature') >= 0) {
					var titleObj = Api.ParseJSON(title)
					if (titleObj.feature && titleObj.feature.zone_type == 'statistics') {
						deleteAccurate(oDrawing)
					}
				}
			}
		} else if (options.cmd == 'open') {
			function updateFooter(oFooter, type, PageMargins, PageSize) {
				if (!oFooter) {
					return
				}
				var oParagraph = oFooter.GetElement(0)
				if (!oParagraph) {
					return
				}
				if (options.page_type == 0) {
					// 统计
					var stat = options.stat || {}
					var oStatsDrawing = Api.CreateImage(
						stat.url,
						(stat.width || 4.8) * 36e3,
						(stat.height || 4.8) * 36e3
					)
					var paraDrawing2 = oStatsDrawing.getParaDrawing()
					if (paraDrawing2) {
						// 统计以左上角为基点
						var sx = PageSize.W - stat.right || 0
						var sy = PageSize.H - stat.bottom || 0
						paraDrawing2.Set_PositionH(6, false, sx, false);
						paraDrawing2.Set_PositionV(5, false, sy, false)
						paraDrawing2.Set_DrawingType(2);
						var titleobj = {
							feature: {
								zone_type: 'statistics',
								footer_type: type,
								v: 1
							}
						}
						oStatsDrawing.SetTitle(JSON.stringify(titleobj))
						oParagraph.AddDrawing(oStatsDrawing)
						paraDrawing2.Set_Parent(oParagraph.Paragraph)
					}
				}
			}
			for (var i = 0; i < oSections.length; ++i) {
				var oSection = oSections[i]
				var PageMargins = oSection.Section.PageMargins
				var PageSize = oSection.Section.PageSize
				var footerList = []
				var oTitleFooter = oSection.GetFooter('title', false)
				if (oTitleFooter) {
					footerList.push({
						type: 'title',
						oFooter: oTitleFooter
					})
				}
				var evenFooter = oSection.GetFooter('even', false)
				if (evenFooter) {
					footerList.push({
						type: 'even',
						oFooter: evenFooter
					})
				}
				var oDefaultFooter = oSection.GetFooter('default', false)
				if (oDefaultFooter) {
					footerList.push({
						type: 'default',
						oFooter: oDefaultFooter
					})
				}
				footerList.forEach((footerObj) => {
					updateFooter(footerObj.oFooter, footerObj.type, PageMargins, PageSize)
				})
			}
		}
	}, false, recalc)
}
function removeAllHeaderFooter() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var oSections = oDocument.GetSections()
		if (!oSections) {
			return
		}
		for (var i = 0; i < oSections.length; ++i) {
			var oSection = oSections[i]
			oSection.RemoveHeader('default')
			oSection.RemoveHeader('title')
			oSection.RemoveHeader('even')
			oSection.RemoveFooter('default')
			oSection.RemoveFooter('even')
			oSection.RemoveFooter('title')
		}
	}, false, false)
}
function drawHeaderFooter(options, calc) {
	return removeAllHeaderFooter().then(()=> {
		return drawHeaderFooter0(options, calc)
	})
}
// 绘制页眉页脚
function drawHeaderFooter0(options, calc) {
	Asc.scope.options_header_footer = options
	return biyueCallCommand(window, function() {
		var options = Asc.scope.options_header_footer || {}
		console.log('drawHeaderFooter', options)
		var oDocument = Api.GetDocument()
		var oSections = oDocument.GetSections()
		if (!oSections) {
			return
		}
		var pstyle = options.pagination ? options.pagination.align_style : 'center'
		oDocument.SetEvenAndOddHdrFtr(pstyle != 'center');
		function updateText(obj, oParagraph, defaultAlign) {
			if (!oParagraph) {
				return
			}
			if (obj && obj.text) {
				oParagraph.AddText(obj.text)
				if (obj.font_bold) {
					oParagraph.SetBold(true)
				}
				if (obj.font_family) {
					oParagraph.SetFontFamily(obj.font_family)
				}
				if (obj.font_size) {
					var twips = obj.font_size / (25.4 / 72 / 20)
					oParagraph.SetFontSize(twips / 10)
				}
				oParagraph.SetJc(obj.align || defaultAlign)
			}
		}
		function updateHeader(oHeader, PageMargins) {
			if (!oHeader) {
				return
			}
			var elementCount = oHeader.GetElementsCount()
			if (elementCount > 2) {
				for(var i = elementCount - 1; i > 0; i--) {
					oHeader.RemoveElement(i)
				}
			}
			var oParagraph = oHeader.GetElement(0)
			if (!oParagraph) {
				return
			}
			oParagraph.RemoveAllElements()
			var header = options.header || {}
			updateText(header, oParagraph, 'center')
			oParagraph.SetBottomBorder(header.line_visible ? 'single' : 'none', 1, 2, 153, 153, 153)
			if (header.image_url) {
				var width = header.image_width || 10 // mm
				var height = header.image_height || 10 // mm
				var oDrawing = Api.CreateImage(header.image_url, width * 36e3, height * 36e3)
				oParagraph.AddDrawing(oDrawing)
				var paraDrawing = oDrawing.getParaDrawing()
				if (paraDrawing) {
					var x = header.image_x
					var y = header.image_y
					if (header.image_x == undefined) {
						x = PageMargins.Left
						y = PageMargins.Top - height - 1
					}
					paraDrawing.Set_PositionH(6, false, x, false)
					paraDrawing.Set_PositionV(5, false, y, false)
					paraDrawing.Set_DrawingType(2)
					paraDrawing.Set_Parent(oParagraph.Paragraph)
				}
			}
		}
		function getPageNumberDrawing(shapeWidth, shapeHeight, type) {
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
					zone_type: 'pagination',
					footer_type: type
				}
			}
			oDrawing.SetTitle(JSON.stringify(titleobj))
			return oDrawing
		}
		function setPaginationAlign(oDrawing, align, margin) {
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
			if (drawDocument) {
				var paragraphs = drawDocument.GetAllParagraphs()
				if (paragraphs && paragraphs.length > 0) {
					paragraphs[0].SetJc(align)
				}
			}
		}
		var numberDrawing = null
		function updateFooter(oFooter, type, PageMargins, PageSize) {
			if (!oFooter) {
				return
			}
			var elementCount = oFooter.GetElementsCount()
			if (elementCount > 2) {
				for(var i = elementCount - 1; i > 0; i--) {
					oFooter.RemoveElement(i)
				}
			}

			var oParagraph = oFooter.GetElement(0)
			if (!oParagraph) {
				return
			}
			oParagraph.RemoveAllElements()
			var footer = options.footer || {}
			oParagraph.SetTopBorder(footer.line_visible ? 'single' : 'none', 1, 2, 153, 153, 153)
			if (footer.line_visible) {
				var oDrawing = Api.CreateShape(
					'rect',
					28 * 36e3,
					6 * 36e3,
					Api.CreateSolidFill(Api.CreateRGBColor(255, 255, 255)),
					Api.CreateStroke(0, Api.CreateNoFill())
				)
				var drawContent = oDrawing.GetContent()
				var paragraphs = drawContent.GetAllParagraphs()
				if (paragraphs && paragraphs.length) {
					var oRun = Api.CreateRun()
					oRun.AddText('线外请勿作答')
					paragraphs[0].AddElement(oRun)
					paragraphs[0].SetColor(153, 153, 153, false)
					paragraphs[0].SetFontSize(18)
				}
				oDrawing.SetPaddings(0, 0, 0, 0)
				var paraDrawing = oDrawing.getParaDrawing()
				if (paraDrawing) {
					oDrawing.SetHorAlign('page', 'center')
					paraDrawing.Set_PositionV(6, false, -3, false)
					paraDrawing.Set_DrawingType(2)
				}
				oParagraph.AddDrawing(oDrawing)
				paraDrawing.Set_Parent(oParagraph.Paragraph)
			}
			updateText(footer, oParagraph, 'left')
			if (options.page_type == 0) {
				// 统计
				var stat = options.stat || {}
				var oStatsDrawing = Api.CreateShape('rect',
					(stat.width || 4.8) * 36e3,
					(stat.height || 4.8) * 36e3,
					Api.CreateSolidFill(Api.CreateRGBColor(255, 255, 255)),
					Api.CreateStroke(0, Api.CreateNoFill()))
				var statContent = oStatsDrawing.GetContent()
				var paragraphs = statContent.GetAllParagraphs()
				if (paragraphs && paragraphs.length) {
					paragraphs[0].AddText('\ue628')
					paragraphs[0].SetFontFamily('iconfont')
					paragraphs[0].SetFontSize(28)
					paragraphs[0].SetColor(153, 153, 153, false)
				}
				oStatsDrawing.SetPaddings(0, 0, 0, 0)
				// var oStatsDrawing = Api.CreateImage(
				// 	stat.url,
				// 	(stat.width || 4.8) * 36e3,
				// 	(stat.height || 4.8) * 36e3
				// )
				var paraDrawing2 = oStatsDrawing.getParaDrawing()
				if (paraDrawing2) {
					// 统计以左上角为基点
					oStatsDrawing.SetHorPosition('rightMargin', (PageMargins.Right - stat.right || 0) * 36e3)
					oStatsDrawing.SetVerPosition('page', (PageSize.H - stat.bottom || 0) * 36e3)
					paraDrawing2.Set_DrawingType(2);
					var titleobj = {
						feature: {
							zone_type: 'statistics',
							footer_type: type,
							v: 1
						}
					}
					oStatsDrawing.SetTitle(JSON.stringify(titleobj))
					oParagraph.AddDrawing(oStatsDrawing)
					paraDrawing2.Set_Parent(oParagraph.Paragraph)
				}
			}
			// 页码
			var oAddNum = getPageNumberDrawing(20, options.pagination.font_size, type) // numberDrawing.Copy()
			var align = 'center'
			if (options.pagination.align_style == 'oddLeftEvenRight') {
				align = type == 'even' ? 'right' : 'left'
			} else if (options.pagination.align_style == 'oddRightEvenLeft') {
				align = type == 'even' ? 'left' : 'right'
			}
			setPaginationAlign(oAddNum, align, options.pagination.margin)
			oAddNum.SetVerPosition('page', (PageSize.H - options.pagination.bottom || 0) * 36e3)
			oParagraph.AddDrawing(oAddNum)
			var paraDrawing3 = oAddNum.getParaDrawing()
			if (paraDrawing3) {
				paraDrawing3.Set_DrawingType(2)
				paraDrawing3.Set_Parent(oParagraph.Paragraph)
			}
		}
		for (var i = 0; i < oSections.length; ++i) {
			var oSection = oSections[i]
			var PageMargins = oSection.Section.PageMargins
			var PageSize = oSection.Section.PageSize
			var footerList = []
			var headerList = []
			var oTitleHeader = oSection.GetHeader('title', false)
			var oTitleFooter = oSection.GetFooter('title', false)
			if (oSection.Section.IsTitlePage()) {
				if (!oTitleHeader) {
					oTitleHeader = oSection.GetHeader('title', true)
				}
				if (!oTitleFooter) {
					oTitleFooter = oSection.GetFooter('title', true)
				}
			}
			if (oTitleFooter) {
				footerList.push({
					type: 'title',
					oFooter: oTitleFooter
				})
			}
			if (oTitleHeader) {
				headerList.push({
					type: 'title',
					oHeader: oTitleHeader
				})
			}
			if (pstyle != 'center') {
				var evenFooter = oSection.GetFooter('even', false)
				if (!evenFooter) {
					evenFooter = oSection.GetFooter('even', true)
				}
				footerList.push({
					type: 'even',
					oFooter: evenFooter
				})
			}
			var oDefaultFooter = oSection.GetFooter('default', false)
			if (!oDefaultFooter) {
				oDefaultFooter = oSection.GetFooter('default', true)
			}
			footerList.push({
				type: 'default',
				oFooter: oDefaultFooter
			})
			var oDefaultHeader = oSection.GetHeader('default', false)
			if (!oDefaultHeader) {
				oDefaultHeader = oSection.GetHeader('default', true)
			}
			headerList.push({
				type: 'default',
				oHeader: oDefaultHeader
			})
			var oEvenHeader = oSection.GetHeader('even', false)
			if (!oEvenHeader) {
				oEvenHeader = oSection.GetHeader('even', true)
			}
			headerList.push({
				type: 'even',
				oHeader: oEvenHeader
			})
			headerList.forEach((headerObj) => {
				var oHeader = headerObj.oHeader
				updateHeader(oHeader, PageMargins)
				oSection.SetHeaderDistance((PageMargins.Top - 6) / (25.4 / 72 / 20))
			})
			// numberDrawing = getPageNumberDrawing(20, options.pagination.font_size)
			footerList.forEach((footerObj) => {
				updateFooter(footerObj.oFooter, footerObj.type, PageMargins, PageSize)
				oSection.SetFooterDistance((PageMargins.Bottom - 4) / (25.4 / 72 / 20))
			})
		}
		console.log('==================== draw header footer end')
	}, false, calc)
}

export { handleFeature, handleHeader, drawExtroInfo, setLoading, deleteAllFeatures, setInteraction, updateChoice, handleChoiceUpdateResult, showOrHidePagination,drawHeaderFooter, drawStatistics }
