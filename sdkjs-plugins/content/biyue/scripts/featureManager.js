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
	drawList([options])
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
	drawList(list)
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
	loading = false
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
			drawList(newlist)
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
									oDrawing.Delete()
								}
							} else {
								oDrawing.Delete()
							}
						}
					}
				}
			}
		}
	}, false, true).then(() => {
		console.log('功能区都已删除')
	})
}

function drawList(list) {
	loading = true
	Asc.scope.feature_wait_handle = list
	Asc.scope.pos_list = window.BiyueCustomData.pos_list
	console.log('========================= drawList', list)
	Asc.scope.ZONE_TYPE = ZONE_TYPE
	biyueCallCommand(window, function() {
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
		function GetPageParagraphForDraw(page_num) {
			var pageCount = oDocument.GetPageCount()
			if (page_num >= pageCount) {
				console.log('+++++++++++ GetPageParagraphForDraw', page_num, pageCount)
				return
			}
			oDocument.Document.GoToPage(page_num)
			var curParagraph = oDocument.Document.GetCurrentParagraph()
			if (page_num == 0) {
				return curParagraph
			} else if (page_num > 0) {
				var Pages = curParagraph.Pages
				if (Pages) {
					if (Pages.length == 1) {
						return curParagraph
					} else {
						if (curParagraph.GetAbsolutePage(Pages.length - 1) == page_num) {
							var nextParagraph = Api.LookupObject(curParagraph.Id).GetNext()
							if (nextParagraph && nextParagraph.Paragraph.GetAbsolutePage(0) == page_num) {
								return nextParagraph.Paragraph
							} else {
								console.log('nextParagraph is null')
							}
						} else {
							console.log('curParagraph end page != page_num')
						}
					}
				}
				return curParagraph
			}
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
			if (options.cmd == 'close') {
				// 关闭
				if (oDrawing) {
					result.code = 1
					oDrawing.Delete()
				} else {
					result.code = 2
					result.message = '未找到该区域'
				}
			} else {
				if (!oDrawing) {
					if (options.size) {
						var shapeWidth = options.size.w
						var shapeHeight = options.size.h
						// if (options.zone_type == ZONE_TYPE.AGAIN) {
						//   var firstSection = oSections[0]
						//   var oHeader = firstSection.GetHeader("default", true);
						//   var oParagraph = oHeader.GetElement(0);
						//   console.log('oHeader', oHeader)
						//   console.log('oHeader', oParagraph)
						// }
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
							// oTable.GetCell(0, 6).SetWidth('twips', 10 / MM2TWIPS)
							// addImageToCell(
							// 	oTable,
							// 	0,
							// 	7,
							// 	'https://by-qa-image-cdn.biyue.tech/xiaotao.png',
							// 	21.33 * scale,
							// 	30 * scale,
							// 	0,
							// 	0
							// )
							// addTextToCell(
							// 	oTable,
							// 	0,
							// 	8,
							// 	[options.label],
							// 	'center',
							// 	textw,
							// 	20,
							// 	0,
							// 	fmargin
							// )
							// addImageToCell(oTable, 0, 9, furl, fw, fh, fmargin, fmargin)
							// addImageToCell(oTable, 0, 10, furl, fw, fh, fmargin, fmargin)
							// addImageToCell(oTable, 0, 11, furl, fw, fh, fmargin, fmargin)
							// addImageToCell(oTable, 0, 12, furl, fw, fh, fmargin, fmargin)

							// oTable.SetWidth('twips', w / MM2TWIPS)
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
								var finalSection = oSections[oSections.length - 1]
								var oFooter = finalSection.GetFooter("default", true)
								var paragraph = oFooter.GetElement(0)
								var drawing = oDrawing.Drawing;
										// drawing.Set_PositionH(6, false, options.x, false);
										// drawing.Set_PositionV(5, false, options.y, false);
										paragraph.SetJc('right')
										paragraph.Paragraph.AddToParagraph(drawing);
							} else {
								var page_num = options.page_num || options.p
								var paragraph = GetParagraphForDraw(page_num, 'end')
								if (paragraph) {
									var drawing = oDrawing.Drawing;
										drawing.Set_PositionH(6, false, options.x, false);
										drawing.Set_PositionV(5, false, options.y, false);
										drawing.Set_DrawingType(2);
										paragraph.AddToParagraph(drawing);
								} else {
									console.log('cannot find paragraph for draw', page_num)
								}
							}
						}
					}
				} else {
					if (oDrawing) {
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
		return res
	}, false, true).then(res => {
		loading = false
		console.log('drawList result:', res)
		if (res && res.list) {
			var pos_list = window.BiyueCustomData.pos_list || []
			res.list.forEach((result) => {
				var index = pos_list.findIndex((e) => {
					return e.zone_type == result.zone_type && e.v == result.v
				})
				if (index >= 0) {
					if (result.cmd == 'close') {
						pos_list[index].drawing_id = null
					} else {
						pos_list[index].drawing_id = result.drawing_id
						pos_list[index].x = result.x
						pos_list[index].y = result.y
					}
				} else if (result.cmd == 'open') {
					pos_list.push({
						zone_type: result.zone_type,
						v: result.v,
						drawing_id: result.drawing_id,
						x: result.x,
						y: result.y,
					})
				}
			})
			window.BiyueCustomData.pos_list = pos_list
		}
		handleNext()
	})
}

function setInteraction(type) {
	Asc.scope.interaction_type = type
	return biyueCallCommand(window, function() {
		var interaction_type = Asc.scope.interaction_type
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var MM2TWIPS = 25.4 / 72 / 20
		if (!controls) {
			return
		}
		function getExistDrawing(draws, sub_type_list) {
			var list = []
			for (var i = 0; i < draws.length; ++i) {
				var title = draws[i].Drawing.docPr.title
				if (title) {
					var titleObj = JSON.parse(title)
					if (titleObj.feature) {
						if (titleObj.feature.zone_type == 'question' && sub_type_list.indexOf(titleObj.feature.sub_type) >= 0) {
							list.push(draws[i])
						}
					}
				}
			}
			return list
		}

		function addSimple(oControl) {
			if (!oControl) {
				return
			}
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(
				3600,
				Api.CreateSolidFill(Api.CreateRGBColor(225, 225, 225))
			)
			var oDrawing = Api.CreateShape(
				'rect',
				3.18 * 36e3,
				3.18 * 36e3,
				oFill,
				oStroke
			)
			var titleobj = {
				feature: {
					zone_type: 'question',
					type: 'ques_interaction',
					sub_type: 'simple',
					control_id: oControl.Sdt.GetId()
				}
			}
			oDrawing.Drawing.Set_Props({
				title: JSON.stringify(titleobj),
			})
			oDrawing.SetDistances(0, 0, 2 * 36e3, 0);
			var paragraphs = oControl.GetAllParagraphs()
			if (paragraphs && paragraphs.length > 0) {
				var pParagraph = paragraphs[0]
				var oRun = Api.CreateRun()
				oRun.AddDrawing(oDrawing)
				pParagraph.AddElement(
					oRun,
					0
				)
				oDrawing.SetWrappingStyle('tight')
			}
		}

		function addAskInteraction(askControl, index) {
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(
				3600,
				Api.CreateSolidFill(Api.CreateRGBColor(225, 225, 225))
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
				paragraphs[0].AddElement(oRun)
				paragraphs[0].SetColor(3, 3, 3, false)
				paragraphs[0].SetFontSize(14)
				paragraphs[0].SetJc('center')
				oDrawing.SetPaddings(0, 0, 0, 0)
			}
			var titleobj = {
				feature: {
					zone_type: 'question',
					type: 'ques_interaction',
					sub_type: 'ask_accurate',
					control_id: oControl.Sdt.GetId()
				}
			}
			oDrawing.Drawing.Set_Props({
				title: JSON.stringify(titleobj),
			})
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
		}
		function addAccurate(oControl) {
			if (!oControl) {
				return
			}
			var control_id = oControl.Sdt.GetId()
			var childControls = oControl.GetAllContentControls()
			var write_list = []
			if (childControls) {
				for (var i = 0; i < childControls.length; ++i) {
					var parentControl = childControls[i].GetParentContentControl()
					if (!parentControl || parentControl.Sdt.GetId() != control_id) {
						continue
					}
					var tag = JSON.parse(childControls[i].GetTag() || '{}')
					if (tag.regionType == 'write') {
						addAskInteraction(childControls[i], write_list.length + 1)
						write_list.push(childControls[i])
					}
				}
			}
			var num = write_list.length
			if (num == 0) {
				return
			}
			var rect = Api.asc_GetContentControlBoundingRect(control_id, true)
			var newRect = {
				Left: rect.X0,
				Right: rect.X1,
				Top: rect.Y0,
				Bottom: rect.Y1,
			}
			var controlContent = oControl.GetContent()
			if (controlContent) {
				var pageIndex = 0
				if (
					controlContent.Document &&
					controlContent.Document.Pages &&
					controlContent.Document.Pages.length > 1
				) {
					for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
						if (!oControl.Sdt.IsEmptyPage(p)) {
							pageIndex = p
							break
						}
					}
				}
				var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
				if (pagebounds) {
					newRect.Right = Math.max(pagebounds.Right, newRect.Right)
				}
			}
			var width = newRect.Right - newRect.Left
			var oTable = Api.CreateTable(num, 1)
			var oTableStyle = oDocument.CreateStyle('CustomTableStyle', 'table')
			var oTableStylePr = oTableStyle.GetConditionalTableStyle('wholeTable')
			oTable.SetTableLook(true, true, true, true, true, true)
			oTableStylePr.GetTableRowPr().SetHeight('atLeast', 8 / MM2TWIPS) // 高度至少多少trips
			var oTableCellPr = oTableStyle.GetTableCellPr()
			oTableCellPr.SetVerticalAlign('center')
			oTable.SetWrappingStyle(true)
			oTable.SetStyle(oTableStyle)
			oTable.SetCellSpacing(150)
			oTable.SetTableBorderTop('single', 1, 0.1, 255, 255, 255)
			oTable.SetTableBorderBottom('single', 1, 0.1, 255, 255, 255)
			oTable.SetTableBorderLeft('single', 1, 0.1, 255, 255, 255)
			oTable.SetTableBorderRight('single', 1, 0.1, 255, 255, 255)
			var Props = {
				CellSelect: true,
				Locked: false,
				PositionV: {
					Align: 1,
					RelativeFrom: 2,
					UseAlign: true,
					Value: 0,
				},
				PositionH: {
					Align: 4,
					RelativeFrom: 0,
					UseAlign: true,
					Value: 0,
				},
				TableDefaultMargins: {
					Bottom: 0,
					Left: 0,
					Right: 0,
					Top: 0,
				},
			}
			oTable.Table.Set_Props(Props)
			for (var i = 0; i < num; ++i) {
				var cell = oTable.GetCell(0, i)
				if (cell) {
					var cellcontent = cell.GetContent()
					if (cellcontent) {
						var oCellPara = cellcontent.GetElement(0)
						if (oCellPara) {
							oCellPara.AddText(`${i + 1}`)
							cell.SetWidth('twips', 8 / MM2TWIPS)
							oCellPara.SetJc('center')
							oCellPara.SetColor(0, 0, 0, false)
							oCellPara.SetFontSize(16)
						} else {
							console.log('oCellPra is null')
						}
					} else {
						console.log('cellcontent is null')
					}
				}
			}

			var shapew = num * 12
			var oDrawing = Api.CreateShape(
				'rect',
				shapew * 36e3,
				(8 + 4) * 36e3,
				Api.CreateNoFill(),
				Api.CreateStroke(0, Api.CreateNoFill())
			)
			var titleobj = {
				feature: {
					zone_type: 'question',
					type: 'ques_interaction',
					sub_type: 'accurate',
					control_id: control_id
				}
			}
			oDrawing.Drawing.Set_Props({
				title: JSON.stringify(titleobj),
			})
			oDrawing.SetPaddings(0, 0, 0, 0)
			var drawDocument = oDrawing.GetContent()
			drawDocument.AddElement(0, oTable)
			oDrawing.SetWrappingStyle('square')
			oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
			var oRun = Api.CreateRun()
			oRun.AddDrawing(oDrawing)
			var paragraphs = oControl.GetContent().GetAllParagraphs()
			var find = false
			if (paragraphs && paragraphs.length > 0) {
				for (var p = 0; p < paragraphs.length; ++p) {
					var paragraphParent = paragraphs[p].GetParentContentControl()
					if (paragraphParent && paragraphParent.Sdt.GetId() == oControl.Sdt.GetId()) {
						paragraphs[p].AddElement(oRun, 0)
						find = true
						console.log('add accurate success', oControl.Sdt.GetId())
						break
					}
				}
			}
			if (!find) {
				console.warn('cannot find paragraph')
			}
		}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tag = JSON.parse(oControl.GetTag() || '{}')
			if (tag.regionType == 'question' || tag.regionType == 'sub-question') {
				var allDraws = oControl.GetAllDrawingObjects()
				var simpleDrawings = getExistDrawing(allDraws, ['simple'])
				var accurateDrawings = getExistDrawing(allDraws, ['accurate', 'ask_accurate'])
				if (interaction_type == 'simple') {
					if (!simpleDrawings || simpleDrawings.length == 0) {
						addSimple(oControl)	
					}
					if (accurateDrawings && accurateDrawings.length) {
						for (var j = 0; j < accurateDrawings.length; ++j) {
							accurateDrawings[j].Delete()
						}
					}
					
				} else if (interaction_type == 'accurate') {
					if (!simpleDrawings || simpleDrawings.length == 0) {
						addSimple(oControl)
					}
					if (!accurateDrawings || accurateDrawings.length == 0) {
						addAccurate(oControl)
					}
				}
			}
		}

	}, false, true).then(res => {
		console.log('setInteraction result:', res)
	})
}

export { handleFeature, handleHeader, drawExtroInfo, deleteAllFeatures, setInteraction }
