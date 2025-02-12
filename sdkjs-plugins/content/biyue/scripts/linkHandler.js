// 这个文件主要处理图片或表格关联相关操作
import { biyueCallCommand } from "./command.js";
import { preGetExamTree, focusControl } from "./QuesManager.js";

function tagImageCommon(params) {
	Asc.scope.tag_params = params
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
			// console.log('[tagImageCommon] begin')
			var tag_params = Asc.scope.tag_params
			var client_node_id = Asc.scope.client_node_id
			var oDocument = Api.GetDocument()
			var drawings = oDocument.GetAllDrawingObjects() || []
			var uid = 0
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
					uid = title.tid
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
						tag.feature = {
							ques_use: tag_params.ques_use.join('_'),
							client_id: client_node_id
						}
					}
					oDrawing.SetTitle(JSON.stringify(tag))
					uid = tag.pid
				}
			}
			return {
				client_node_id: client_node_id,
				drawing_id: tag_params.target_id,
				ques_use: tag_params.ques_use,
				uid: uid
			}
	}, false, false, {name: 'tagImageCommon'}).then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			if (!window.BiyueCustomData.image_use) {
				window.BiyueCustomData.image_use = {}
			}
			window.BiyueCustomData.image_use[res.drawing_id] = res.ques_use
			window.biyue.sendToDialog('pictureIndex', 'updateUse', {
				from: 'ques_use',
				list: [{
					type: params.target_type,
					uid: res.uid,
					sort_id: res.uid ? res.uid.substring(2, res.uid.length) * 1 : 0,
					id: res.drawing_id,
					ques_use: res.ques_use
				}]
			}, false)
			if (params.ques_id) {
				return ShowLinkedWhenclickImage({
					client_id: params.ques_id
				})
			} else {
				return ShowLinkedWhenclickImage()
			}
		} else {
			return new Promise((resolve, reject) => {
				return resolve()
			})
		}
	})
}
// 点击就近关联
function imageAutoLink(ques_id, calc) {
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.ques_id = ques_id || 0
	Asc.scope.link_type = window.BiyueCustomData.link_type || 'all'
	Asc.scope.link_coverage_percent = window.BiyueCustomData.link_coverage_percent || 80
	return biyueCallCommand(window, function() {
			// console.log('[imageAutoLink] begin')
			var oDocument = Api.GetDocument()
			var allDrawings = oDocument.GetAllDrawingObjects() || []
			var tables = oDocument.GetAllTables() || []
			var controls = oDocument.GetAllContentControls() || []
			var question_map = Asc.scope.question_map
			var client_node_id = Asc.scope.client_node_id
			var ques_id = Asc.scope.ques_id
			var link_type = Asc.scope.link_type
			var link_coverage_percent = Asc.scope.link_coverage_percent / 100
			var rev = false
			var list_update = []
			// 获取重叠面积
			function getOverlapArea(quesFields, imageFields) {
				var overlapArea = 0
				for (var field of quesFields) {
					for (var image of imageFields) {
						if (image.page != field.page) {
							continue
						}
						let v = Math.max(0, Math.min(image.x + image.w, field.x + field.w) - Math.max(image.x, field.x)) *
								Math.max(0, Math.min(image.y + image.h, field.y + field.h) - Math.max(image.y, field.y))
						overlapArea += v
					}
				}
				return overlapArea
			}
			function getImageArea(imageFields) {
				var area = 0
				for (var image of imageFields) {
					area += image.w * image.h
				}
				return area
			}
			// 是否符合全包
			function allInFields(quesFields, fields) {
				var hasPartOverlap = false
				for (var field of fields) {
					// 每个区域都处于题目范围中
					var flag = 0
					for (var quesField of quesFields) {
						if (quesField.page != field.page) {
							continue
						}
						// y值在范围内
						if (field.y >= quesField.y - 1 && 
							field.y + field.h <= quesField.y + quesField.h + 1
						) {
							if (field.x >= quesField.x - 1 && 
								field.x + field.w <= quesField.x + quesField.w + 1
							) { // 全在范围内
								flag = 1
								break
							} else if (!(field.x + field.w < quesField.x || quesField.x + quesField.w < field.x)) { // 有x值交集
								flag = 2
								break
							}
						}
					}
					if (!flag) {
						return 0
					} else if (flag == 2) {
						hasPartOverlap = true
					}
				}
				return hasPartOverlap ? 2 : 1
			}
			// 获取经过的题目控件
			function getBelongQuesList(fields) {
				if (!fields || fields.length == 0) {
					return []
				}
				var list = []
				var partList = []
				var imageArea = getImageArea(fields)
				if (imageArea == 0) {
					return []
				}
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
					var quesFields = []
					if (Pages) {
						for (var j = 0; j < Pages.length; ++j) {
							var page = Pages[j]
							if ( page.Bounds) {
								let w = page.Bounds.Right - page.Bounds.Left
								let h = page.Bounds.Bottom - page.Bounds.Top
								if (w > 0 && h > 0) {
									var pindex = oControl.Sdt.GetAbsolutePage(j)
									var bounds = oControl.Sdt.GetContentBounds(j)
									quesFields.push({
										page: pindex,
										x: page.Bounds.Left,
										y: bounds.Top, // page.Bounds.Top,
										w: w,
										h: bounds.Bottom - bounds.Top // h
									})
								}
							}
						}
					}
					if (link_type == 'all') { // 全包判断
						var allInFlag = allInFields(quesFields, fields)
						if (allInFlag == 1) {
							list.push(tag.mid || tag.client_id)
						} else if (allInFlag == 2) {
							partList.push(tag.mid || tag.client_id)
						}
					} else if (link_type == 'area') { // 面积比例判断
						var overlapArea = getOverlapArea(quesFields, fields)
						if (overlapArea > 0 && overlapArea / imageArea >= link_coverage_percent) {
							list.push(tag.mid || tag.client_id)
						}
					}
				}
				if (list.length) {
					return list
				}
				// else if (partList.length == 1) {
				// 	return partList
				// }
				return null
			}
			function getNewQuesUse(quesuse, quesList) {
				// if (quesuse) {
				// 	if (typeof quesuse == 'number') {
				// 		quesuse = quesuse + ''
				// 	}
				// 	if (typeof quesuse == 'string') {
				// 		var uselist = quesuse.split('_')
				// 		for (var id of quesList) {
				// 			if (!uselist.find(e => { return e == id })) {
				// 				uselist.push(id)
				// 			}
				// 		}
				// 		return uselist.join('_')
				// 	}
				// }
				return quesList.join('_')
			}
			// 图片关联
			for (var i = 0; i < allDrawings.length; ++i) {
				var oDrawing = allDrawings[i]
				var Drawing = oDrawing.getParaDrawing()
				if (!Drawing) {
					continue
				}
				var title = Api.ParseJSON(oDrawing.GetTitle())
				if (title.ignore) {
					continue
				}
				if (title.feature && title.feature.zone_type) {
					continue
				}
				var quesList = getBelongQuesList([{
					page: Drawing.PageNum,
					x: Drawing.X,
					y: Drawing.Y,
					w: Drawing.Width,
					h: Drawing.Height
				}])
				if (!quesList || quesList.length == 0) {
					if (title.feature && title.feature.ques_use) {
						title.feature.ques_use = ''
						oDrawing.SetTitle(JSON.stringify(title))
					}
					continue
				}
				if (title.feature) {
					if (!title.feature.client_id) {
						client_node_id += 1
						title.feature.client_id = client_node_id
					}
					var quesuse = title.feature.ques_use
					title.feature.ques_use = getNewQuesUse(quesuse, quesList)
				} else {
					client_node_id += 1
					title.feature = {
						ques_use: quesList.join('_'),
						client_id: client_node_id
					}
				}
				rev = true
				oDrawing.SetTitle(JSON.stringify(title))
				list_update.push({
					type: 'drawing',
					uid: title.pid,
					sort_id: title.pid ? title.pid.substring(2, title.pid.length) * 1 : 0,
					id: oDrawing.Drawing.Id,
					ques_use: quesList,
					partical_no_dot: title.feature && title.feature.partical_no_dot
				})
			}
			// 暂不考虑大小题的问题
			// 表格关联
			for (var i = 0; i < tables.length; ++i) {
				var oTable = tables[i]
				var strtitle = oTable.GetTableTitle()
				if (strtitle == 'questionTable' || strtitle == JSON.stringify({ignore: 1})) {
					continue
				}
				var title = Api.ParseJSON(strtitle)
				var pageCount = oTable.Table.GetPagesCount()
				var tableFields = []
				for (var p = 0; p < pageCount; ++p) {
					var pageBounds = oTable.Table.GetPageBounds(p)
					var pagenum = oTable.Table.GetAbsolutePage(p)
					tableFields.push({
						page: pagenum,
						x: pageBounds.Left,
						y: pageBounds.Top,
						w: pageBounds.Right - pageBounds.Left,
						h: pageBounds.Bottom - pageBounds.Top
					})
				}
				var quesList = getBelongQuesList(tableFields)
				if (!quesList || quesList.length == 0) {
					if (title.ques_use) {
						title.ques_use = ''
						oTable.SetTableTitle(JSON.stringify(title))
					}
					continue
				}
				if (!title.client_id) {
					client_node_id += 1
					title.client_id = client_node_id
				}
				title.ques_use = getNewQuesUse(title.ques_use, quesList)
				oTable.SetTableTitle(JSON.stringify(title))
				list_update.push({
					type: 'table',
					uid: title.tid,
					sort_id: title.tid ? title.tid.substring(2, title.tid.length) * 1 : 0,
					id: oTable.Table.Id,
					ques_use: quesList,
					partical_no_dot: title.partical_no_dot
				})
				rev = true
			}
			return {
				client_node_id,
				rev,
				list_update
			}
	}, false, calc, {name: 'imageAutoLink'}).then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			if (res.list_update) {
				window.biyue.sendToDialog('pictureIndex', 'updateUse', {
					from: 'ques_use',
					list: res.list_update
				}, false)
			}
		}
		return new Promise((resolve, reject) => {resolve(res)})
	})
}
function onAllCheck() {
	return biyueCallCommand(window, function() {
			// console.log('[onAllCheck] begin')
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
					if (title.ignore) {
						continue
					}
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
							if (child && child.Id == oDrawing.Drawing.Id) {
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
					if (tabletitle == 'questionTable' || tabletitle == JSON.stringify({ignore: 1})) {
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
			var oContrls = oDocument.GetAllContentControls()
			oContrls.forEach(oControl => {
				var tag = Api.ParseJSON(oControl.GetTag())
				if (tag.color == '#ffff0020') {
					tag.color = tag.clr
					oControl.SetTag(JSON.stringify(tag))
				}
			})
			return allList
	}, false, false, {name: 'onAllCheck'}).then(res => {
		Asc.scope.linked_list = res
		return preGetExamTree()
	}).then(res => {
		Asc.scope.tree_info = res
		window.biyue.showDialog('elementLinks', '关联检查', 'elementLinks.html', 900, 500, false)
	})
}
// 获取已关联列表
function onLinkedCheck() {
	return biyueCallCommand(window, function() {
			// console.log('[onLinkedCheck] begin')
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
								if (child && child.Id == oDrawing.Drawing.Id) {
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
			return linkedList
	}, false, false, {name: 'onLinkedCheck'}).then(res => {
		Asc.scope.linked_list = res
		return preGetExamTree()
	}).then(res => {
		Asc.scope.tree_info = res
		window.biyue.showDialog('elementLinks', '关联检查', 'elementLinks.html', 900, 500, false)
	})
}

function updateLinkedInfo(info) {
	Asc.scope.link_info = info
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
			// console.log('[updateLinkedInfo] begin')
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
						tag.feature = {
							ques_use: data.ques_use,
							client_id: client_node_id
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
	}, false, false, {name: 'updateLinkedInfo'}).then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
		}
		return new Promise((resolve, reject) => {
			return resolve()
		})
	})
}
// 点击图片后，自动显示关联的题目
function ShowLinkedWhenclickImage(options, control_id) {
	Asc.scope.control_id = control_id
	Asc.scope.click_options = options
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	return biyueCallCommand(window, function() {
			// console.log('[ShowLinkedWhenclickImage] begin')
			var oDocument = Api.GetDocument()
			var selectedDrawings = oDocument.GetSelectedDrawings() || []
			var allDrawings = oDocument.GetAllDrawingObjects() || []
			var control_id = Asc.scope.control_id
			var question_map = Asc.scope.question_map
			var click_options = Asc.scope.click_options || {}
			var oState = oDocument.Document.SaveDocumentState()
			var ids = []
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
			if (!control_id && click_options.client_id) {
				var control = getControlsByClientId(click_options.client_id)
				if (control) {
					control_id = control.Sdt.GetId()
				}
			}
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
							oDrawing.SetShadow('tl', 40, 100, null, 1, '#f2ba02')
							flag = 1
						}
					}
				}
				if (flag == 0) {
					if (dtag.feature && dtag.feature.partical_no_dot) {
						oDrawing.SetShadow(null, 0, 100, null, 0, '#0fc1fd')
					} else {
						if (oDrawing.Drawing.spPr && dtag.feature && dtag.feature.ques_use) {
							oDrawing.ClearShadow()
						}
					}
				}
			})
			oDocument.Document.LoadDocumentState(oState)
	}, false, false, {name: 'ShowLinkedWhenclickImage'})
}

function locateItem(data) {
	Asc.scope.locate_data = data
	return biyueCallCommand(window, function() {
			// console.log('[locateItem] begin')
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
	}, false, false, {name: 'locateItem'})
}
// 获取所有图片
function getPictureList(outOfRange) {
	Asc.scope.picture_id = window.BiyueCustomData.picture_id || 0
	Asc.scope.table_id = window.BiyueCustomData.table_id || 0
	Asc.scope.outOfRange = outOfRange
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var oSections = oDocument.GetSections() || []
		var pageSize = {}
		if (oSections[0] && oSections[0].Section) {
			pageSize.width = oSections[0].Section.GetPageWidth()
			pageSize.height = oSections[0].Section.GetPageHeight()
		}
		var picture_id = Asc.scope.picture_id
		var outOfRange = Asc.scope.outOfRange
		var list = []
		var allDrawings = oDocument.GetAllDrawingObjects() || []
		var list_ignore = []
		function isOutOfRange(oDrawing) {
			if (!oDrawing || !oDrawing.Drawing) {
				return false
			}
			var bounds = oDrawing.Drawing.getRectBounds ? oDrawing.Drawing.getRectBounds() : oDrawing.Drawing.bounds
			if (!bounds) {
				return false
			}
			return bounds.l < 0 || bounds.l > pageSize.width || bounds.t < 0 || bounds.t > pageSize.height || bounds.r < 0 || bounds.r > pageSize.width || bounds.b < 0 || bounds.b > pageSize.height
		}
		function isTableOutOfRange(oTable) {
			if (!oTable || !oTable.Table) {
				return false
			}
			var bounds = oTable.Table.GetPageBounds ? oTable.Table.GetPageBounds(0) : oTable.Table.Bounds
			if (!bounds) {
				return false
			}
			return bounds.Left < 0 || bounds.Right > pageSize.width
		}
		for (var oDrawing of allDrawings) {
			var title = Api.ParseJSON(oDrawing.GetTitle())
			var Drawing = oDrawing.getParaDrawing()
			if (!Drawing) {
				continue
			}
			var title = Api.ParseJSON(oDrawing.GetTitle())
			if (!outOfRange && title.ignore == 1) {
				continue
			}
			if (title.feature && title.feature.zone_type) {
				continue
			}
			if (!title.pid) {
				title.pid = `d_${++picture_id}` 
				oDrawing.SetTitle(JSON.stringify(title))
			}
			var targetList = title.ignore == 2 ? list_ignore : list
			var obj = {
				type: 'drawing',
				uid: title.pid,
				sort_id: title.pid ? title.pid.substring(2, title.pid.length) * 1 : 0,
				id: oDrawing.Drawing.Id,
				ques_use: title.feature && title.feature.ques_use ? getQuesUse(title.feature.ques_use) : [],
				partical_no_dot: title.feature && title.feature.partical_no_dot
			}
			if (outOfRange) {
				if (isOutOfRange(oDrawing)) {
					targetList.push(obj)
				}
			} else {
				targetList.push(obj)
			}
		}
		var table_id = Asc.scope.table_id
		var tables = oDocument.GetAllTables() || []
		for (var oTable of tables) {
			var strtitle = oTable.GetTableTitle()
			if (strtitle == 'questionTable') {
				continue
			}
			var title = Api.ParseJSON(strtitle)
			if (!outOfRange && title.ignore == 1) {
				continue
			}
			if (!title.tid) {
				title.tid = `t_${++table_id}`
				oTable.SetTableTitle(JSON.stringify(title))
			}
			var targetList = title.ignore == 2 ? list_ignore : list
			var isOutOfRange = isTableOutOfRange(oTable)
			var obj = {
				type: 'table',
				uid: title.tid,
				sort_id: title.tid ? title.tid.substring(2, title.tid.length) * 1 : 0,
				id: oTable.Table.Id,
				ques_use: getQuesUse(title.ques_use),
				partical_no_dot: title.partical_no_dot
			}
			if (outOfRange) {
				if (isTableOutOfRange(oTable)) {
					targetList.push(obj)
				}
			} else {
				targetList.push(obj)
			}
		}
		function getQuesUse(ques_use) {
			if (!ques_use) {
				return []
			}
			var useList = ques_use.split('_')
			return useList
		}
		return {
			picture_id,
			table_id,
			list,
			list_ignore
		}
	}, false, false, {name: 'getPictureList'})
}

function handleIgnore(data) {
	Asc.scope.pic_data = data
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var picData = Asc.scope.pic_data
		if (picData.type == 'table') {
			var tables = oDocument.GetAllTables() || []
			var oTable = tables.find(e => {
				return oTable.Table.Id == picData.id
			})
			if (oTable) {
				var title = Api.ParseJSON(oTable.GetTableTitle())
				if (picData.ignore) {
					title.ignore = 2
					if (title.ques_use) {
						delete title.ques_use
					}
				} else if (title.ignore == 2) {
					delete title.ignore
				}
				oTable.SetTableTitle(JSON.stringify(title))
			}
		} else {
			var drawings = oDocument.GetAllDrawingObjects() || []
			var oDrawing = drawings.find(e => {
				return e.Drawing.Id == picData.id
			})
			if (oDrawing) {
				var title = Api.ParseJSON(oDrawing.GetTitle())
				if (picData.ignore) {
					title.ignore = 2
					if (title.feature && title.feature.ques_use) {
						delete title.feature.ques_use
					}
				} else if (title.ignore == 2) {
					delete title.ignore
				}
				oDrawing.SetTitle(JSON.stringify(title))
			}
		}
		
	}, false, false, {name: 'handleIgnore'})
}
// 更新局部不铺码
function updateImageParticalNoDot(params) {
	Asc.scope.tag_params = params
	return biyueCallCommand(window, function() {
		var params = Asc.scope.tag_params
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		var oDrawing = drawings.find(e => {
			return e.Drawing.Id == params.target_id
		})
		if (oDrawing) {
			var tag = Api.ParseJSON(oDrawing.GetTitle())
			if (params.partical_no_dot) {
				if (tag.feature) {
					tag.feature.partical_no_dot = 1
				} else {
					tag.feature = {
						partical_no_dot: 1
					}
				}
			} else {
				if (tag.feature && tag.feature.partical_no_dot) {
					delete tag.feature.partical_no_dot
				}
			}
			oDrawing.SetTitle(JSON.stringify(tag))
			if (params.partical_no_dot) {
				oDrawing.SetShadow('ctr', 0, 100, null, 0, '#007bff')
			} else {
				oDrawing.ClearShadow()
			}
		}
	}, false, false, {name: 'updateImageParticalNoDot'})
}

function handlePictureIndexMessage(modal, message) {
	switch(message.cmd) {
		case 'ignore':
			handleIgnore(message.data)
			break
		case 'link':
			tagImageCommon(message.data)
			break
		case 'locateQues':
			var event = new CustomEvent('focusQuestion', {
				detail: message.data
			})
			document.dispatchEvent(event)
			focusControl(message.data.ques_id).then(() => {
				ShowLinkedWhenclickImage({
					client_id: message.data.ques_id
				})
			})
			break
		case 'updateDot':
			updateImageParticalNoDot(message.data)
			break
		case 'autoLink':
			window.BiyueCustomData.link_type = message.data.link_type
			window.BiyueCustomData.link_coverage_percent = message.data.link_coverage_percent
			imageAutoLink(null, true).then(res => {
				return getPictureList()
			}).then(res => {
				if (res.picture_id) {
					window.BiyueCustomData.picture_id = res.picture_id
				}
				if (res.table_id) {
					window.BiyueCustomData.table_id = res.table_id
				}
				modal.command('pictureIndexMessage', {
					BiyueCustomData: window.BiyueCustomData,
					list: res.list,
					list_ignore: res.list_ignore,
					type: 'autolink'
				})
			})
			break
		case 'refresh':
			getPictureList(message.data == 'pictureList').then(res => {
				if (res.picture_id) {
					window.BiyueCustomData.picture_id = res.picture_id
				}
				if (res.table_id) {
					window.BiyueCustomData.table_id = res.table_id
				}
				if (message.data == 'pictureList') {
					modal.command('pictureListMessage', {
						list: res.list
					})
				} else {
					modal.command('pictureIndexMessage', {
						BiyueCustomData: window.BiyueCustomData,
						list: res.list,
						list_ignore: res.list_ignore
					})
				}
			})
			break
		default:
			break
	}
}

export {
	tagImageCommon,
	imageAutoLink,
	onLinkedCheck,
	updateLinkedInfo,
	onAllCheck,
	locateItem,
	ShowLinkedWhenclickImage,
	getPictureList,
	handlePictureIndexMessage
}