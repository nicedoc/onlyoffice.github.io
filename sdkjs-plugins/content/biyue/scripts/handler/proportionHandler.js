// 这个文件专门处理占比相关
import { biyueCallCommand } from '../command.js'
import { setChoiceOptionLayout } from '../choiceQuestion.js'
import { isChoiceMode, isTextMode } from '../model/ques.js'
import { getWorkbookInteraction } from '../model/feature.js'
import { setInteraction } from '../featureManager.js'
// 批量修改占比
function batchProportion(idList, proportion) {
	if (!idList || idList.length == 0) {
		return
	}
	Asc.scope.change_id_list = idList
	Asc.scope.proportion = proportion
	return biyueCallCommand(window, function() {
		var idList = Asc.scope.change_id_list
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var TABLE_TITLE = 'questionTable'
		var target_proportion = Asc.scope.proportion ? Asc.scope.proportion * 1 : 1
		var newW = 100 / target_proportion
		var sections = oDocument.GetSections()
		var oSection = sections[0]
		var PageSize = oSection.Section.PageSize
		var PageMargins = oSection.Section.PageMargins
		var tw = PageSize.W - PageMargins.Left - PageMargins.Right
		var iTable = 0 // 目标表格数量
		var tables = {} // 要处理的表格
		var tempList = [] // 要删除添加的内容
		var rangeTableIds = []
		var wBefore = 0
		var targetStartPos = -1
		var effect_id_list = []
		function getControlsByClientId(cid) {
			var allControls = oDocument.GetAllContentControls() || []
			var findControls = allControls.filter(e => {
				var tag = Api.ParseJSON(e.GetTag())
				if (e.GetClassType() == 'blockLvlSdt') {
					return tag.client_id == cid && e.GetPosInParent() >= 0
				} else if (e.GetClassType() == 'inlineLvlSdt') {
					return e.Sdt && e.Sdt.GetPosInParent() >= 0 && tag.client_id == cid
				}
			})
			if (findControls && findControls.length) {
				return findControls[0]
			}
		}
		function addToTable(cellW, targetPos) {
			if (wBefore + cellW <= 100) {
				if (tables[iTable]) {
					tables[iTable].cells.push({
						icell: tables[iTable].cells.length,
						W: cellW
					})
					tables[iTable].W += cellW
				} else {
					if (targetStartPos >= 0) {
						++targetStartPos
					} else {
						targetStartPos = targetPos
					}
					tables[iTable] = {
						cells: [{
							icell: 0,
							W: cellW
						}],
						W: cellW,
						tablePos: targetStartPos
					}
				}
				wBefore += cellW
			} else {
				iTable++
				targetStartPos++
				tables[iTable] = {
					cells: [{
						icell: 0,
						W: cellW
					}],
					W: cellW,
					tablePos: targetStartPos
				}
				wBefore = cellW
			}
		}
		function addTable(oTable, tableWidth, overBreak) {
			var cellCount = oTable.GetRow(0).GetCellsCount()
			for (var c1 = 0; c1 < cellCount; ++c1) {
				var oCell = oTable.GetCell(0, c1)
				var cellBounds = oCell.Cell.GetPageBounds(0)
				var cellW = (cellBounds.Right - cellBounds.Left) * 100 / tw
				if (overBreak && c1 == 0) { // 超出范围的，当后续不需要再往上合并时，不再执行
					if (wBefore + cellW > 100) {
						wBefore = cellW
						return false
					}
				}
				var oCellContent = oCell.GetContent()
				var cellChildCount = oCellContent.GetElementsCount()
				for (var i3 = 0; i3 < cellChildCount; ++i3) {
					var oCellChild = oCellContent.GetElement(i3)
					if (oCellChild.GetClassType() == 'blockLvlSdt') {
						var tag = Api.ParseJSON(oCellChild.GetTag())
						if (tag.client_id && idList.find(e => {return e.id == tag.client_id})) {
							cellW = newW
						}
						effect_id_list.push(tag.client_id)
					}
					tempList.push({
						oElement: oCellChild,
						oParent: oCellContent,
						posinparent: i3
					})
					addToTable(cellW, oTable.GetPosInParent())
				}
			}
			return true
		}
		function AddControlToCell2(content, oCell) {
			if (!oCell) {
				console.log('AddControlToCell2 ocell is null')
				return
			}
			if (!content) {
				console.log('AddControlToCell2 content is null')
				return
			}
			oCell.SetCellMarginLeft(0)
			oCell.SetCellMarginRight(0)
			var cellContent = oCell.GetContent()
			content.forEach((e, index) => {
				cellContent.AddElement(index, e)
			})
			var lastindex = cellContent.GetElementsCount() - 1
			if (cellContent.GetElementsCount() > 1 && cellContent.GetElement(lastindex).GetClassType() == 'paragraph') {
				cellContent.RemoveElement(lastindex)
			}
		}
		// 先算出要处理的题目范围
		var beginPos = -1
		var endPos = -1
		for (var item of idList) {
			var id = item.id
			var quesData = question_map[id]
			if (!quesData) {
				continue
			}
			// 这里不考虑合并题
			var oControl = getControlsByClientId(id)
			if (!oControl) {
				continue
			}
			var oControlRange = oControl.GetRange()
			if (oControlRange && oControlRange.StartPos && oControlRange.StartPos.length && oControlRange.StartPos[0].Class) {
				var startElement = Api.LookupObject(oControlRange.StartPos[0].Class.Id)
				if (startElement.GetClassType() == 'document') {
					var pos = oControlRange.StartPos[0].Position
					if (beginPos == -1) {
						endPos = beginPos = pos
					} else {
						if (beginPos > pos) {
							beginPos = pos
						}
						if (endPos < pos) {
							endPos = pos
						}
					}
				}
			}
		}
		// 这里先不考虑表格嵌套的情况
		var elementCount1 = oDocument.GetElementsCount()
		for (var i1 = beginPos; i1 < elementCount1; ++i1) {
			var element = oDocument.GetElement(i1)
			if (element.GetClassType() == 'table') {
				if (element.GetTableTitle() != TABLE_TITLE) {
					if (i1 > endPos) {
						break
					} else {
						++targetStartPos
						if (tables[iTable]) {
							++iTable
						}
						wBefore = 0
						continue
					}
				}
				var tableBounds = element.Table.GetPageBounds(0)
				var tableWidth = tableBounds.Right - tableBounds.Left
				if (!addTable(element, tableWidth, i1 > endPos)) {
					break
				}
				rangeTableIds.push(element.Table.Id)
			} else {
				if (i1 > endPos) {
					break
				}
				if (element.GetClassType() == 'blockLvlSdt') {
					var tag = Api.ParseJSON(element.GetTag())
					if (!tag.client_id) {
						continue
					}
					var item = idList.find(e => {
						return e.id == tag.client_id
					})
					if (!item) {
						continue
					}
					element.Sdt.GetLogicDocument().PreventPreDelete = true
					// 第1个需要考虑往前合并
					var flag1 = 0
					if (i1 == beginPos) {
						flag1 = target_proportion > 1 ? 1 : 0
						// control本就不在表格中，且占比为1不考虑，否则需要将control合并到前一个表格中
						if (target_proportion > 1 && i1 > 0) {
							var preElement = oDocument.GetElement(i1 - 1)
							if (preElement && 
								preElement.GetClassType() == 'table' &&
								preElement.GetTableTitle() == TABLE_TITLE &&
								preElement.GetRowsCount() == 1
							) { // 前一个是题目表格
								var tableBounds = preElement.Table.GetPageBounds(0)
								var tableWidth = tableBounds.Right - tableBounds.Left
								var TableW = null
								if (preElement.TablePr && preElement.TablePr.TableW && preElement.TablePr.TableW.Type == 3) {
									TableW = preElement.TablePr.TableW
								}
								if (TableW && (TableW.W + newW) <= 100 || 
								(tw / target_proportion + tableWidth <= tw)) {
									// 前一个表格符合条件，最后一个单元格需要拆分
									addTable(preElement, tableWidth)
									rangeTableIds.push(preElement.Table.Id)
									flag1 = 2
								}
							}
						}
					} else {
						flag1 = 1
					}
					if (flag1 > 0) {
						addToTable(newW, element.GetPosInParent())
						effect_id_list.push(tag.client_id)
						tempList.push({
							oElement: element,
							oParent: oDocument,
							posinparent: element.GetPosInParent()
						})
					}
				} else {
					++targetStartPos
					if (tables[iTable]) {
						++iTable
					}
					wBefore = 0
				}
			}
		}
		// console.log('====== templist', tempList)
		// console.log('====== tables', tables)
		// console.log('====== rangeTableIds', rangeTableIds)
		// 从后往前删除内容
		for (var j = tempList.length - 1; j >= 0; --j) {
			tempList[j].oParent.RemoveElement(tempList[j].posinparent)
		}
		// 将删除的内容一一填入目标位置
		var count = 0
		var oTable = null
		for (var i = 0; i <= iTable; ++i) {
			var table = tables[i]
			if (!table) {
				continue
			}
			var targetCellCount = table.cells.length
			if (i < rangeTableIds.length) {
				oTable = Api.LookupObject(rangeTableIds[i])
				if (oTable) {
					var row = oTable.GetRow(0)
					var cellCount = row.GetCellsCount()
					if (targetCellCount > cellCount) { // 要拆分
						oTable.Split(oTable.GetCell(0, 0), 1, targetCellCount - cellCount + 1);
					} else if (targetCellCount < cellCount) { // 要合并
						var mergeCells = []
						for (var m = 0; m <= cellCount - targetCellCount; ++m) {
							mergeCells.push(oTable.GetCell(0, m))
						}
						oTable.MergeCells(mergeCells)
					}
					if (oTable.GetPosInParent() != table.tablePos) {
						oDocument.RemoveElement(oTable.GetPosInParent())
						oDocument.AddElement(table.tablePos, oTable)
					}
				}
			} else {
				oTable = Api.CreateTable(targetCellCount, 1)
				oTable.SetCellSpacing(0)
				oTable.SetTableTitle(TABLE_TITLE)
				oDocument.AddElement(table.tablePos, oTable)
			}
			if (oTable) {
				oTable.SetWidth('percent', tables[i].W)
				if (tables[i].cells.length == 1 && tables[i].cells[0].W == 100) {
					oDocument.AddElement(tables[i].tablePos, tempList[count].oElement)
					oDocument.RemoveElement(tables[i].tablePos + 1)
					count++
				} else {
					tables[i].cells.forEach((cell, cidx) => {
						AddControlToCell2([tempList[count].oElement], oTable.GetCell(0, cell.icell))
						count++
						if (cell.W > 0) {
							var cw = cell.W <= 0 ? 1 : cell.W
							var twips = oTable.Table.CalculatedTableW * (100 / cw)
							if (cw > 1) {
								oTable.GetCell(0, cell.icell).SetWidth('percent', cw)
							} else {
								oTable.GetCell(0, cell.icell).SetWidth('twips', twips)
							}
						}
					})
				}
			}
		}
		if (iTable >= rangeTableIds.length) {
			// todo..可能需要将多余的表格删除
		}
		return {
			proportion_list: idList.map(e => {
				return e.id
			}),
			effect_id_list: effect_id_list,
			proportion: target_proportion
		}
	}, false, false).then((res) => {
		return handleProportionSuccess(res)
	})
}
// 占比修改成功后的处理
function handleProportionSuccess(res) {
	if (window.BiyueCustomData.question_map) {
		res.proportion_list.forEach(id => {
			if (window.BiyueCustomData.question_map[id]) {
				window.BiyueCustomData.question_map[id].proportion = res.proportion
			}
		})
		var interaction_list = []
		var layout_list = []
		if (res.effect_id_list) {
			var workbookInteraction = getWorkbookInteraction()
			for (var id of res.effect_id_list) {
				var quesData = window.BiyueCustomData.question_map[id]
				if (!quesData || quesData.level_type == 'struct') {
					continue
				}
				if (isTextMode(quesData.ques_mode)) {
					continue
				}
				if (quesData.interaction === undefined) {
					if (workbookInteraction != 'none') {
						interaction_list.push(id)	
					}
				} else if (quesData.interaction != 'none') {
					interaction_list.push(id)
				}
				if (isChoiceMode(quesData.ques_mode) && quesData.part != undefined && quesData.part !== 1) {
					layout_list.push(id)
				}
			}
		}
		return new Promise((resolve, reject) => {
			if (interaction_list.length) {
				return setInteraction('useself', interaction_list, false).then(() => {
					return resolve()
				})
			} else {
				return resolve()
			}
		}).then(() => {
			if (layout_list.length) {
				return setChoiceOptionLayout({
					list: layout_list,	
					from: 'proportion'
				})
			} else {
				return new Promise((resolve, reject) => {
					resolve()
				})
			}
		})
	}
}
function changeProportion(idList, proportion) {
	if (!idList || idList.length == 0) {
		return
	}
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.change_id_list = idList
	Asc.scope.proportion = proportion
	return biyueCallCommand(window, function() {
		var idList = Asc.scope.change_id_list
		var question_map = Asc.scope.question_map || {}
		var target_proportion = Asc.scope.proportion || 1
		var oDocument = Api.GetDocument()
		var oControls = oDocument.GetAllContentControls()
		var TABLE_TITLE = 'questionTable'
		var newW = 100 / target_proportion
		var effect_id_list = []
		function AddControlToCell(oControl, oCell) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			var temp = []
			temp.push(oControl)
			var posinparent = oControl.GetPosInParent()
			var parent = oControl.Sdt.GetParent()
			var oParent = Api.LookupObject(parent.Id)
			oParent.RemoveElement(posinparent)
			oCell.SetWidth('percent', 100 / target_proportion)
			oCell.AddElement(0, temp[0])
			if (oCell.GetContent().GetElementsCount() == 2) {
				oCell.GetContent().RemoveElement(1)
			}
		}

		function AddControlToCell2(content, oCell) {
			if (!oCell) {
				console.log('AddControlToCell2 ocell is null')
				return
			}
			if (!content) {
				console.log('AddControlToCell2 content is null')
				return
			}
			oCell.SetCellMarginLeft(0)
			oCell.SetCellMarginRight(0)
			var cellContent = oCell.GetContent()
			content.forEach((e, index) => {
				cellContent.AddElement(index, e)
			})
			var lastindex = cellContent.GetElementsCount() - 1
			if (cellContent.GetElementsCount() > 1 && cellContent.GetElement(lastindex).GetClassType() == 'paragraph') {
				cellContent.RemoveElement(lastindex)
			}
		}
		// 上一个表格有足够的空间可以放，需要将上一个表格的最后一个单元格拆分，然后下面的表格内容挪到上面
		function func1(startTable, oControl) {
			if (!startTable || !oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			var templist = []
			var iTable = 0
			var tables = {}
			var wBefore = 0
			var toIndex = -1
			var tag = Api.ParseJSON(oControl.GetTag())
			var startTablePos = startTable.GetPosInParent()
			var dataHandled = {}
			dataHandled.parentTable = oControl.GetParentTable()
			dataHandled.parentTableCell = oControl.GetParentTableCell()
			if (dataHandled.parentTableCell) {
				dataHandled.cellIndex = dataHandled.parentTableCell.GetIndex()
				dataHandled.rowIndex = dataHandled.parentTableCell.GetRowIndex()
				dataHandled.table_id = dataHandled.parentTable.Table.Id
				dataHandled.table_pos = dataHandled.parentTable.GetPosInParent()
			} else {
				dataHandled.pos = oControl.GetPosInParent()
				if (dataHandled.pos < startTablePos) {
					if (tag.client_id) {
						effect_id_list.push(tag.client_id)
					}
					templist.push({
						content: [oControl]
					})
					tables[iTable] = {
						cells: [{
							icell: 0,
							W: newW
						}],
						W: newW
					}
					wBefore = newW
					var oControlParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
					oControlParent.RemoveElement(dataHandled.pos)
				}
			}
			startTablePos = startTable.GetPosInParent()
			var startTableParent = startTable.Table.Parent
			var oStartTableParent = Api.LookupObject(startTableParent.Id)
			var elementCount1 = oStartTableParent.GetElementsCount()
			for (var i1 = startTablePos; i1 < elementCount1; ++i1) {
				var element = oStartTableParent.GetElement(i1)
				if (element.GetClassType() != 'table') {
					if (dataHandled.pos != undefined && dataHandled.pos == i1 && element.GetClassType() == 'blockLvlSdt' && element.Sdt.GetId() == oControl.Sdt.GetId()) {
						if (tag.client_id) {
							effect_id_list.push(tag.client_id)
						}
						templist.push({
							content: [oControl]
						})
						if (wBefore + newW > 100) {
							iTable++
							tables[iTable] = {
								cells: [{
									icell: 0,
									W: newW
								}],
								W: newW
							}
							wBefore = newW
						} else {
							tables[iTable].cells.push({
								icell: tables[iTable].cells.length,
								W: newW
							})
							tables[iTable].W += newW
							wBefore += newW
						}
						var oControlParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
						oControlParent.RemoveElement(dataHandled.pos)
					}
					break
				}
				if (element.GetTableTitle() != TABLE_TITLE) {
					break
				}
				// todo..分栏里的表格尚未考虑
				// var tableBounds = element.Table.GetPageBounds(0)
				// var tableWidth = tableBounds.Right - tableBounds.Left
				var sectPr = element.Table.Get_SectPr()
				var sectWidth = sectPr.PageSize.W - sectPr.PageMargins.Left - sectPr.PageMargins.Right
				var rows = element.GetRowsCount()
				if (rows != 1) {
					break
				}
				var oTableRow = element.GetRow(0)
				var cellsCount = oTableRow.GetCellsCount()
				for (var icell = 0; icell < cellsCount; ++icell) {
					var oCell = oTableRow.GetCell(icell)
					var TableCellW = oCell.CellPr.TableCellW
					if (!TableCellW) {
						TableCellW = oCell.Cell.CompiledPr.Pr.TableCellW
					}
					if (!TableCellW) {
						continue
					}
					var pagebounds = oCell.Cell.GetPageBounds(0)
					var cellWidth = pagebounds.Right - pagebounds.Left
					var W = cellWidth * 100 / sectWidth // TableCellW.W
					var cellContent = oCell.GetContent()
					var contents = []
					var elementcount = cellContent.GetElementsCount()
					oControl.Sdt.GetLogicDocument().PreventPreDelete = true
					for (var k = 0; k < elementcount; ++k) {
						var cellChild = cellContent.GetElement(k)
						if (cellChild.GetClassType() == 'blockLvlSdt') { 
							var ctag = Api.ParseJSON(cellChild.Sdt.GetTag())
							if (ctag.client_id) {
								effect_id_list.push(ctag.client_id)
							}
						}
						contents.push(cellChild)
					}
					for (var k = 0; k < elementcount; ++k) {
						cellContent.RemoveElement(0)
					}
					templist.push({
						content: contents
					})
					if (dataHandled.table_id) {
						if (i1 < dataHandled.table_pos) { // 上一个表格
							if (!tables[iTable]) {
								tables[iTable] = {
									cells: [{
										icell: 0,
										W: W
									}],
									W: W
								}
							} else {
								tables[iTable].cells.push({
									icell: tables[iTable].cells.length,
									W: W
								})
								tables[iTable].W += W
							}
							wBefore += W
						} else {
							var W2 = W
							if (dataHandled.table_pos == i1 && dataHandled.cellIndex == icell) {
								W2 = newW
							}
							if (wBefore + W2 <= 100) {
								if (!tables[iTable]) {
									tables[iTable] = {
										cells: [{
											icell: 0,
											W: W2
										}],
										W: W2
									}
									wBefore = W2
								} else {
									tables[iTable].cells.push({
										icell: tables[iTable].cells.length,
										W: W2
									})
									tables[iTable].W += W2
									wBefore += W2
								}
							} else {
								iTable++
								tables[iTable] = {
									cells: [{
										icell: 0,
										W: W2
									}],
									W: W2
								}
								wBefore = W2
							}
						}
					} else { // 当前未处于表格中
						if (tables[iTable] && tables[iTable].W + W > 100) {
							++iTable
							wBefore = 0
						}
						if (!tables[iTable]) {
							tables[iTable] = {
								cells: [{
									icell: 0,
									W: W
								}],
								W: W
							}
						} else {
							tables[iTable].cells.push({
								icell: tables[iTable].cells.length,
								W: W
							})
							tables[iTable].W += W
						}
						wBefore += W
					}
				}
				++toIndex
			}
			var count = 0
			for (var i = 0; i <= iTable; ++i) {
				if (!tables[i]) {
					break
				}
				var targetcellscount = tables[i].cells.length
				var oTable
				if (i <= toIndex) {
					oTable = oStartTableParent.GetElement(startTablePos + i)
					if (oTable.GetClassType() != 'table') {
						break
					}
					var oRow = oTable.GetRow(0)
					var cellscount = oRow.GetCellsCount()
					if (cellscount < targetcellscount) { // 需要拆分
						oTable.Split(oTable.GetCell(0, 0), 1, targetcellscount - cellscount + 1);
					} else if (cellscount > targetcellscount) { // 需要合并
						var mergeCells = []
						for (var m = 0; m <= cellscount - targetcellscount; ++m) {
							mergeCells.push(oTable.GetCell(0, m))
						}
						oTable.MergeCells(mergeCells)
					}
				} else {
					oTable  = Api.CreateTable(targetcellscount, 1)
					oTable.SetCellSpacing(0)
					oTable.SetTableTitle(TABLE_TITLE)
				}
				oTable.SetWidth('percent', tables[i].W)
				if (tables[i].cells.length == 1 && tables[i].cells[0].W == 100) {
					oStartTableParent.AddElement(startTablePos + i, templist[count].content[0])
					oStartTableParent.RemoveElement(startTablePos + i + 1)
					count++
				} else {
					tables[i].cells.forEach((cell, cidx) => {
						AddControlToCell2(templist[count].content, oTable.GetCell(0, cell.icell))
						count++
						if (cell.W > 0) {
							var cw = cell.W <= 0 ? 1 : cell.W
							var twips = oTable.Table.CalculatedTableW * (100 / cw)
							if (cw > 1) {
								oTable.GetCell(0, cell.icell).SetWidth('percent', cw)
							} else {
								oTable.GetCell(0, cell.icell).SetWidth('twips', twips)
							}
						}					
						// oTable.GetCell(0, cell.icell).SetWidth('percent', cell.W)
					})
					if (i >= toIndex) {
						oStartTableParent.AddElement(startTablePos + i, oTable)
					}
				}
			}
			if (iTable < toIndex) {
				for (var j = 0; j < toIndex - iTable; ++j) {
					oStartTableParent.RemoveElement(startTablePos + iTable + j + 1)
				}
			}
		}

		for (var idx = 0; idx < idList.length; ++idx) {
			var id = idList[idx]
			var quesData = question_map[id]
			if (!quesData) {
				continue
			}
			var oControl = oControls.find(e => {
				var tag = Api.ParseJSON(e.GetTag())
				return tag.client_id == id && e.GetClassType() == 'blockLvlSdt'
			})
			if (!oControl) {
				continue
			}
			oControl.Sdt.GetLogicDocument().PreventPreDelete = true
			var oTable = oControl.GetParentTable()
			var posinparent = oControl.GetPosInParent()
			var parent = oControl.Sdt.GetParent()
			var oParent = Api.LookupObject(parent.Id)
			var preTable = null
			if (oParent.GetClassType() == 'document') {
				if (posinparent) {
					var previousElement = oParent.GetElement(posinparent - 1)
					if (previousElement && previousElement.GetClassType() == 'table' && previousElement.GetTableTitle() == TABLE_TITLE) {
						var TableW = previousElement.TablePr.TableW.W
						if (TableW + newW <= 100) {
							preTable = previousElement
						}
					}
				}
			} else if (oParent.GetClassType() == 'documentContent') {
				var contentParent = Api.LookupObject(oParent.Document.GetParent().Id)
				if (contentParent && contentParent.GetClassType() == 'tableCell') {
					var tablepos = oTable.GetPosInParent()
					if (tablepos) {
						var tableParent = Api.LookupObject(oTable.Table.Parent.Id)
						var pre = tableParent.GetElement(tablepos - 1)
						if (pre.GetClassType() == 'table' && pre.GetTableTitle() == TABLE_TITLE) {
							var TableW = pre.TablePr.TableW.W
							if(TableW + newW <= 100 ) {
								preTable = pre
							}
						}
					}
				}
			}
			if (preTable) {
				func1(preTable, oControl)
			} else if (oTable) {
				func1(oTable, oControl)
			} else {
				var addToNextTable = false
				var nextTable = oParent.GetElement(posinparent + 1)
				if (nextTable && nextTable.GetClassType() == 'table' && nextTable.GetTableTitle() == TABLE_TITLE) {
					var firstCell = nextTable.GetRow(0).GetCell(0)
					var TableCellW = firstCell.CellPr.TableCellW
					if (!TableCellW) {
						TableCellW = firstCell.Cell.CompiledPr.Pr.TableCellW
					}
					var firstCellW = TableCellW ? TableCellW.W : 100
					if (firstCellW + newW <= 100) {
						addToNextTable = true
						func1(nextTable, oControl)
					}
				}
				if (!addToNextTable) {
					oTable = Api.CreateTable(1, 1);
					oTable.SetCellSpacing(0)
					oTable.SetWidth('percent', newW);
					oTable.SetTableTitle(TABLE_TITLE)
					var oCell = oTable.GetRow(0).GetCell(0)
					AddControlToCell(oControl, oCell)
					oParent.AddElement(posinparent, oTable)
				}
			}
		}
		return {
			proportion_list: idList,
			effect_id_list: effect_id_list,
			proportion: target_proportion
		}
	}, false, false).then(res => {
		return handleProportionSuccess(res)
	})
}

export default {
	batchProportion,
	changeProportion
}