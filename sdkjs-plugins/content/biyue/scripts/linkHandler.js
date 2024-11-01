// 这个文件主要处理图片或表格关联相关操作
import { biyueCallCommand, dispatchCommandResult } from "./command.js";

function tagImageCommon(params) {
	Asc.scope.tag_params = params
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var tag_params = Asc.scope.tag_params
		var client_node_id = Asc.scope.client_node_id
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		if (tag_params.target_type == 'table') {
			var oTable = Api.LookupObject(tag_params.target_id)
			if (oTable && oTable.GetClassType && oTable.GetClassType() == 'table') {
				var title = Api.ParseJSON(oTable.GetTableTitle())
				if (!title.client_id) {
					client_node_id += 1
					title.client_id = client_node_id
				}
				title.ques_use = tag_params.ques_use.join('_')
				oTable.SetTableTitle(JSON.stringify(title))
			}
		} else {
			var oDrawing = drawings.find(e => {
				return e.Drawing.Id == tag_params.target_id
			})
			if (oDrawing) {
				var tag = Api.ParseJSON(oDrawing.GetTitle())
				if (tag.feature) {
					if (!tag.feature.client_id) {
						client_node_id += 1
						tag.feature.client_id = client_node_id
					}
					tag.feature.ques_use = tag_params.ques_use.join('_')
				} else {
					client_node_id += 1
					tag = {
						feature: {
							ques_use: tag_params.ques_use.join('_'),
							client_id: client_node_id
						}
					}
				}
				oDrawing.SetTitle(JSON.stringify(tag))
			}
		}
		return {
			client_node_id: client_node_id,
			drawing_id: tag_params.target_id,
			ques_use: tag_params.ques_use
		}
	}, false, false).then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			if (!window.BiyueCustomData.image_use) {
				window.BiyueCustomData.image_use = {}
			}
			window.BiyueCustomData.image_use[res.drawing_id] = res.ques_use
			return ShowLinkedWhenclickImage()
		}
	})
}

function imageAutoLink() {
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		var tables = oDocument.GetAllTables() || []
		var controls = oDocument.GetAllContentControls() || []
		var question_map = Asc.scope.question_map
		var client_node_id = Asc.scope.client_node_id
		// 判断是否在control中
		function getBelongControl(pageIndex, x1, y1, x2, y2) {
			for (var i = 0; i < controls.length; ++i) {
				var oControl = controls[i]
				if (oControl.GetClassType() != 'blockLvlSdt') {
					continue
				}
				var tag = Api.ParseJSON(oControl.GetTag() || '{}')
				if (tag.regionType != 'question' || !tag.client_id) {
					continue
				}
				var question_obj = question_map[tag.client_id]
					? question_map[tag.client_id]
					: {}
				// 当前题目必须是blockLvlSdt
				if (question_obj.level_type != 'question') {
					continue
				}
				
				var oControlContent = oControl.GetContent()
				var Pages = oControlContent.Document.Pages
				if (Pages) {
					for (var j = 0; j < Pages.length; ++j) {
						var page = Pages[j]
						if ( page.Bounds) {
							let w = page.Bounds.Right - page.Bounds.Left
							let h = page.Bounds.Bottom - page.Bounds.Top
							if (w > 0 && h > 0) {
								var pindex = oControl.Sdt.GetAbsolutePage(j)
								if (pindex == pageIndex) {
									if (page.Bounds.Left <= x1 && 
										page.Bounds.Top <= y1 &&
										(page.Bounds.Left + w) >= x2 &&
										(page.Bounds.Top + h) >= y2) {
											return {
												ques_id: tag.client_id,
												oControl: oControl
											}
									}
								}
							}
						}
					}
				}
			}
			return null
		}
		// 图片关联
		for (var i = 0; i < allDrawings.length; ++i) {
			var oDrawing = allDrawings[i]
			var Drawing = oDrawing.getParaDrawing()
			if (!Drawing) {
				continue
			}
			var title = Api.ParseJSON(oDrawing.GetTitle())
			if (title.feature && title.feature.zone_type) {
				continue
			}
			var belongControl = getBelongControl(Drawing.PageNum, 
				Drawing.X, 
				Drawing.Y, 
				Drawing.X + Drawing.Width,
				Drawing.Y + Drawing.Height)
			if (belongControl) {
				if (title.feature) {
					if (!title.feature.client_id) {
						client_node_id += 1
						title.feature.client_id = client_node_id
					}
					var quesuse = title.feature.ques_use
					if (quesuse) {
						if (typeof quesuse == 'number') {
							quesuse = quesuse + ''
						}
						if (typeof quesuse == 'string') {
							var uselist = quesuse.split('_')
							if (!(uselist.find(e => { return e == belongControl.ques_id}))) {
								uselist.push(belongControl.ques_id)
								title.feature.ques_use = uselist.join('_')
							}
						}
					} else {
						title.feature.ques_use = belongControl.ques_id + ''
					}
				} else {
					client_node_id += 1
					title = {
						feature: {
							ques_use: belongControl.ques_id + '',
							client_id: client_node_id
						}
					}
				}
				oDrawing.SetTitle(JSON.stringify(title))
			}
		}
		// 暂不考虑大小题的问题
		// 表格关联
		for (var i = 0; i < tables.length; ++i) {
			var oTable = tables[i]
			var strtitle = oTable.GetTableTitle()
			if (strtitle == 'questionTable' || strtitle == 'ignore') {
				continue
			}
			var title = Api.ParseJSON(strtitle)
			var pageCount = oTable.Table.GetPagesCount()
			var belong_id = 0
			for (var p = 0; p < pageCount; ++p) {
				var pageBounds = oTable.Table.GetPageBounds(p)
				var pagenum = oTable.Table.GetAbsolutePage(p)
				var belongControl = getBelongControl(pagenum, 
					pageBounds.Left, 
					pageBounds.Top, 
					pageBounds.Right,
					pageBounds.Bottom)
				if (belongControl) {
					if (belong_id && belong_id != belongControl.ques_id) {
						belong_id = 0
						break
					}
					belong_id = belongControl.ques_id
				}
			}
			if (belong_id) {
				if (!title.client_id) {
					client_node_id += 1
					title.client_id = client_node_id
				}
				var uselist = title.ques_use ? title.ques_use.split('_') : []
				if (!(uselist.find(e => { return e == belongControl.ques_id}))) {
					uselist.push(belongControl.ques_id)
				}
				title.ques_use = uselist.join('_')
				oTable.SetTableTitle(JSON.stringify(title))
			}
		}
		return {
			client_node_id
		}
	}, false, true)
}
function onAllCheck() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		var tables = oDocument.GetAllTables() || []
		var allList = []
		function getHtml(oRange) {
			if (!oRange) {
				return
			}
			oRange.Select()				
			let text_data = {
				data:     "",
				// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
				pushData: function (format, value) {
					if (!value) {
						console.log('========= value is null')
					} else {
						console.log('value is not null')
					}
					this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
				}
			};
			Api.asc_CheckCopy(text_data, 2);
			var rev = text_data.data
			// Api.asc_RemoveSelection();
			return rev
		}
		var i, j
		if (allDrawings.length) {
			for (i = 0; i < allDrawings.length; ++i) {
				var oDrawing = allDrawings[i]
				var paraDrawing = oDrawing.getParaDrawing()
				if (!paraDrawing) {
					continue
				}
				var title = Api.ParseJSON(oDrawing.GetTitle())
				if (title.feature && title.feature.zone_type) {
					continue
				}
				var parentRun = paraDrawing.GetRun()
				if (parentRun) {
					var oRun = Api.LookupObject(parentRun.Id)
					var elementCount = oRun.Run.GetElementsCount()
					var index = 0
					for (j = 0; j < elementCount; ++j) {
						var child = oRun.Run.GetElement(0)
						if (child.Id == oDrawing.Drawing.Id) {
							index = j
							break
						}
					}
					oDrawing.Select()
					var contentpos = oDocument.Document.GetContentPosition()
					var oRange = oRun.GetRange(contentpos, contentpos)
					var html = getHtml(oRange)
					if (html && html.length > 0) {
						allList.push({
							html: html,
							ques_use: title.feature && title.feature.ques_use ? title.feature.ques_use : '',
							type: 'drawing',
							target_id: oDrawing.Drawing.Id
						})
					}
				}
			}
		}
		if (tables.length) {
			for (i = 0; i < tables.length; ++i) {
				var oTable = tables[i]
				var tabletitle = oTable.GetTableTitle()
				if (tabletitle == 'questionTable' || tabletitle == 'ignore') {
					continue
				}
				var title = Api.ParseJSON(tabletitle)
				var html = getHtml(oTable.GetRange())
				if (html && html.length > 0) {
					allList.push({
						html: html,
						ques_use: title && title.ques_use ? title.ques_use : '',
						type: 'table',
						target_id: oTable.Table.Id
					})	
				}
			}
		}
		return allList
	}, false, false).then(res => {
		Asc.scope.linked_list = res
		window.biyue.showDialog('elementLinks', '关联检查', 'elementLinks.html', 900, 500, false)
	})
}
// 获取已关联列表
function onLinkedCheck() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		var tables = oDocument.GetAllTables() || []
		var linkedList = []
		function getHtml(oRange) {
			if (!oRange) {
				return
			}
			console.log('oRange', oRange)
			oRange.Select()				
			let text_data = {
				data:     "",
				// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
				pushData: function (format, value) {
					if (!value) {
						console.log('========= value is null')
					} else {
						console.log('value is not null')
					}
					this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
				}
			};
			Api.asc_CheckCopy(text_data, 2);
			var rev = text_data.data
			// Api.asc_RemoveSelection();
			return rev
		}
		if (allDrawings.length) {
			allDrawings.forEach(oDrawing => {
				var paraDrawing = oDrawing.getParaDrawing()
				var title = Api.ParseJSON(oDrawing.GetTitle())
				if (title.feature && title.feature.ques_use) {
					var parentRun = paraDrawing.GetRun()
					if (parentRun) {
						var oRun = Api.LookupObject(parentRun.Id)
						var elementCount = oRun.Run.GetElementsCount()
						var index = 0
						for (var i = 0; i < elementCount; ++i) {
							var child = oRun.Run.GetElement(0)
							if (child.Id == oDrawing.Drawing.Id) {
								index = i
								break
							}
						}
						oDrawing.Select()
						var contentpos = oDocument.Document.GetContentPosition()
						var oRange = oRun.GetRange(contentpos, contentpos)
						var html = getHtml(oRange)
						linkedList.push({
							html: html,
							ques_use: title.feature.ques_use,
							type: 'drawing',
							target_id: oDrawing.Drawing.Id
						})
					}
				}
			})
		}
		if (tables.length) {
			tables.forEach(oTable => {
				var title = Api.ParseJSON(oTable.GetTableTitle())
				if (title && title.ques_use) {
					var html = getHtml(oTable.GetRange())
					linkedList.push({
						html: html,
						ques_use: title.ques_use,
						type: 'table',
						target_id: oTable.Table.Id
					})
				}
			})
		}
		console.log('linkedList', linkedList)
		return linkedList
	}, false, false).then(res => {
		Asc.scope.linked_list = res
		window.biyue.showDialog('elementLinks', '关联检查', 'elementLinks.html', 900, 500, false)
	})
}

function updateLinkedInfo(info) {
	Asc.scope.link_info = info
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var link_info = Asc.scope.link_info
		var client_node_id = Asc.scope.client_node_id
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		var tables = oDocument.GetAllTables() || []
		function inList(Id, type) {
			for (var i = 0; i < link_info.length; ++i) {
				if (link_info[i].type == type && link_info[i].target_id == Id) {
					return link_info[i]
				}
			}
			return null
		}
		drawings.forEach(oDrawing => {
			var data = inList(oDrawing.Drawing.Id, 'drawing')
			if (data) {
				var tag = Api.ParseJSON(oDrawing.GetTitle())
				if (tag.feature) {
					if (!tag.feature.client_id) {
						client_node_id += 1
						tag.feature.client_id = client_node_id
					}
					tag.feature.ques_use = data.ques_use
				} else {
					client_node_id += 1
					tag = {
						feature: {
							ques_use: data.ques_use,
							client_id: client_node_id
						}
					}
				}
				oDrawing.SetTitle(JSON.stringify(tag))
			}
		})
		tables.forEach(oTable => {
			var data = inList(oTable.Table.Id, 'table')
			if (data) {
				var title = Api.ParseJSON(oTable.GetTableTitle())
				if (!title.client_id) {
					client_node_id += 1
					title.client_id = client_node_id
				}
				title.ques_use = data.ques_use
				oTable.SetTableTitle(JSON.stringify(title))
			}
		})
		return {
			client_node_id: client_node_id
		}
	}, false, false).then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
		}
	})
}
// 点击图片后，自动显示关联的题目
function ShowLinkedWhenclickImage(options, control_id) {
	Asc.scope.control_id = control_id
	Asc.scope.click_options = options
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var selectedDrawings = oDocument.GetSelectedDrawings() || []
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		var control_id = Asc.scope.control_id
		var question_map = Asc.scope.question_map
		var click_options = Asc.scope.click_options || {}
		var oState = oDocument.Document.SaveDocumentState()
		var ids = []
		var selectDrawingCount = selectedDrawings.length
		if (selectDrawingCount > 0) {
			selectedDrawings.forEach(oDrawing => {
				var tag = Api.ParseJSON(oDrawing.GetTitle())
				if (tag && tag.feature && tag.feature.ques_use) {
					ids = ids.concat(tag.feature.ques_use.split('_'))
				}
			})
		}
		var controls = oDocument.GetAllContentControls() || []
		var selectId = null
		controls.forEach(oControl => {
			var tag = Api.ParseJSON(oControl.GetTag())
			if (tag.color == '#ffff0020') {
				if (click_options.ctrlKey || !(ids.includes(tag.client_id + ''))) {
					tag.color = tag.clr
					oControl.SetTag(JSON.stringify(tag))
				}
			} else {
				if (!click_options.ctrlKey && ids.includes(tag.client_id + '')) {
					tag.color = '#ffff0020'
					oControl.SetTag(JSON.stringify(tag))
				}
			}
			if (ids.length == 0 && control_id && control_id == oControl.Sdt.GetId()) {
				if (tag.client_id && question_map[tag.client_id]) {
					selectId = tag.client_id
				} else {
					var parentControl = oControl.GetParentContentControl()
					if (parentControl) {
						var pTag = Api.ParseJSON(parentControl.GetTag())
						if (pTag.client_id && question_map[pTag.client_id]) {
							selectId = pTag.client_id
						}
					}
				}
				
			}
		})
		if (click_options.ctrlKey) {
			return
		}
		allDrawings.forEach(oDrawing => {
			var flag = 0
			var dtag = Api.ParseJSON(oDrawing.GetTitle())
			if (ids.length == 0 && control_id && selectId) {
				if (dtag.feature && dtag.feature.ques_use) {
					var ids2 = dtag.feature.ques_use.split('_')
					if (ids2.includes(selectId + '')) {
						oDrawing.SetShadow(null, 0, 110, null, 0, '#ffa0a0')
						flag = 1
					}
				}
			}
			if (flag == 0) {
				if (dtag.feature && dtag.feature.partical_no_dot) {
					oDrawing.SetShadow(null, 0, 100, null, 0, '#0fc1fd')
				} else {
					if (oDrawing.Drawing.spPr && oDrawing.Drawing.spPr.effectProps && oDrawing.Drawing.spPr.effectProps.EffectLst) {
						oDrawing.ClearShadow()
					}
				}
			}
		})
		oDocument.Document.LoadDocumentState(oState)
	}, false, false)
}

function locateItem(data) {
	Asc.scope.locate_data = data
	return biyueCallCommand(window, function() {
		var locate_data = Asc.scope.locate_data
		var oDocument = Api.GetDocument()
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		if (locate_data.type == 'drawing') {
			var oDrawing = allDrawings.find(e => e.Drawing.Id == locate_data.target_id)
			if (oDrawing) {
				oDrawing.Select()
			}
		} else if (locate_data.type == 'table') {
			var oTable = Api.LookupObject(locate_data.target_id)
			if (oTable) {
				oTable.Select()
			}
		}
	})
}

export {
	tagImageCommon,
	imageAutoLink,
	onLinkedCheck,
	updateLinkedInfo,
	onAllCheck,
	locateItem,
	ShowLinkedWhenclickImage
}