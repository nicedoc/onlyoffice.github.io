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
		function getJsonData(str) {
			if (!str || str == '' || typeof str != 'string') {
				return {}
			}
			try {
				return JSON.parse(str)
			} catch (error) {
				console.log('json parse error', error)
				return {}
			}
		}
		if (tag_params.target_type == 'table') {
			var oTable = Api.LookupObject(tag_params.target_id)
			if (oTable && oTable.GetClassType && oTable.GetClassType() == 'table') {
				var title = getJsonData(oTable.GetTableTitle())
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
				var tag = getJsonData(oDrawing.Drawing.docPr.title)
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
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(tag)
				})
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
		function getJsonData(str) {
			if (!str || str == '' || typeof str != 'string') {
				return {}
			}
			try {
				return JSON.parse(str)
			} catch (error) {
				console.log('json parse error', error)
				return {}
			}
		}
		// 判断是否在control中
		function getBelongControl(pageIndex, x1, y1, x2, y2) {
			for (var i = 0; i < controls.length; ++i) {
				var oControl = controls[i]
				if (oControl.GetClassType() != 'blockLvlSdt') {
					continue
				}
				var tag = getJsonData(oControl.GetTag() || '{}')
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
			if (!oDrawing.Drawing) {
				continue
			}
			var title = getJsonData(oDrawing.Drawing.docPr.title)
			if (title.feature && title.feature.zone_type) {
				continue
			}
			var belongControl = getBelongControl(oDrawing.Drawing.PageNum, 
				oDrawing.Drawing.X, 
				oDrawing.Drawing.Y, 
				oDrawing.Drawing.X + oDrawing.Drawing.Width,
				oDrawing.Drawing.Y + oDrawing.Drawing.Height)
			if (belongControl) {
				if (title.feature) {
					if (!title.feature.client_id) {
						client_node_id += 1
						title.feature.client_id = client_node_id
					}
					if (title.feature.ques_use) {
						var uselist = title.feature.ques_use.split('_')
						if (!(uselist.find(e => { return e == belongControl.ques_id}))) {
							uselist.push(belongControl.ques_id)
							title.feature.ques_use = uselist.join('_')
						}
					} else {
						title.feature.ques_use = title.ques_use.join('_')
					}
				} else {
					client_node_id += 1
					title = {
						feature: {
							ques_use: belongControl.ques_id,
							client_id: client_node_id
						}
					}
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(title)
				})				
			}
		}
		// 暂不考虑大小题的问题
		// 表格关联
		for (var i = 0; i < tables.length; ++i) {
			var oTable = tables[i]
			var strtitle = oTable.GetTableTitle()
			if (strtitle == 'questionTable') {
				continue
			}
			var title = getJsonData(strtitle)
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
// 获取已关联列表
function onLinkedCheck() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		var tables = oDocument.GetAllTables() || []
		var linkedList = []
		function getJsonData(str) {
			if (!str || str == '' || typeof str != 'string') {
				return {}
			}
			try {
				return JSON.parse(str)
			} catch (error) {
				console.log('json parse error', error)
				return {}
			}
		}
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
				var title = getJsonData(oDrawing.Drawing.docPr.title)
				if (title.feature && title.feature.ques_use) {
					var parentRun = oDrawing.Drawing.GetRun()
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
				var title = getJsonData(oTable.GetTableTitle())
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
		function getJsonData(str) {
			if (!str || str == '' || typeof str != 'string') {
				return {}
			}
			try {
				return JSON.parse(str)
			} catch (error) {
				console.log('json parse error', error)
				return {}
			}
		}
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
				var tag = getJsonData(oDrawing.Drawing.docPr.title)
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
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(tag)
				})
			}
		})
		tables.forEach(oTable => {
			var data = inList(oTable.Table.Id, 'table')
			if (data) {
				var title = getJsonData(oTable.GetTableTitle())
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

export {
	tagImageCommon,
	imageAutoLink,
	onLinkedCheck,
	updateLinkedInfo
}