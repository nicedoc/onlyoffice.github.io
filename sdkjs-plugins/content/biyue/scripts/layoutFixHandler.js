import { biyueCallCommand, dispatchCommandResult } from "./command.js";
function layoutDetect(all) {
	Asc.scope.layout_all_range = !!all
	return removeAllComment().then(() => {
		return biyueCallCommand(window, function() {
			var oDocument = Api.GetDocument()
			var oRange = null
			if (Asc.scope.layout_all_range) {
				oRange = oDocument.GetRange()
			} else {
				oRange = oDocument.GetRangeBySelect()
				if (!oRange) {
					var currentContentControl = oDocument.Document.GetContentControl()
					if (currentContentControl) {
						var oControl = Api.LookupObject(currentContentControl.Id)
						if (oControl) {
							oRange = oControl.GetRange()
						}
					}
				}
			}
			if (!oRange) {
				return null
			}
			var paragraphs = oRange.GetAllParagraphs() || []
			var result = {
				has32: false, // 存在空格具有下划线属性
				has160: false, // 存在ASCII码160，这是一个不可打印字符，称为不间断空格（NBSP）
				has65307: false, // 中文分号
				has12288: false, // 中文空格
				hasTab: false, // 存在tab键
				hasWhiteBg: false, // 存在背景为白色的段落
				hasSmallImage: false, // 存在宽高过小的图片
				hasBookmark: false, // 存在书签
			}
			var type_map = {
				'hasSmallImage': '宽高过小的图片',
				'has160': '不间断空格',
				'hasTab': '括号里使用了tab',
				'has32': '空格下划线',
				'has65307': '中文分号',
				'has12288': '中文空格',
				'hasWhiteBg': '段落背景为白色',
				'hasBookmark': '存在书签'
			}
			function isWhite(shd) {
				return shd && shd.Fill && shd.Fill.r == 255 && shd.Fill.g == 255 && shd.Fill.b == 255 && shd.Fill.Auto == false
			}
			var bflag = []
			function handleRun(oRun, pid = ''){
				if (!oRun || oRun.GetClassType() != 'run') {
					return
				}
				var findlist = []
				var textpr = oRun.GetTextPr()
				if (textpr && textpr.TextPr) {
					if (isWhite(textpr.TextPr.Shd)) {
						result.hasWhiteBg = true
						findlist.push('hasWhiteBg')
					}
				}
				var runContent = oRun.Run.Content || []
				var isUnderline = oRun.GetUnderline()
				// 判断是否在括号内存在tab
				for (var k = 0, kmax = runContent.length; k < kmax; ++k) {
					var type = runContent[k].GetType()
					if (type == 22) { // drawing
						if (runContent[k].Width <= 1 && runContent[k].Height <= 1) {
							result.hasSmallImage = true
							findlist.push('hasSmallImage')
						}
					} else if (type == 1) {
						if (runContent[k].Value == 40 || runContent[k].Value == 65288) {
							bflag.push(1)
						} else if (runContent[k].Value == 41 || runContent[k].Value == 65289) {
							if (bflag.length) {
								if (bflag[bflag.length - 1] == 1) {
									bflag.pop()
								} else if (bflag[bflag.length - 1] == 2) {
									result.hasTab = true
									findlist.push('hasTab')
									break
								}
							}
						}
					} else if (type == 21) {
						if (isUnderline) {
							result.has32 = true
							findlist.push('has32')
						}
						if (bflag.length) {
							if (bflag[bflag.length - 1] == 1) {
								bflag.push(2)
							}
						}
					}
					if (runContent[k].Value == 32) {
						if (isUnderline) {
							result.has32 = true
							findlist.push('has32')
						}
					} else if (runContent[k].Value) {
						var vkey = `has${runContent[k].Value}`
						if (result[vkey] != undefined) {
							if (!result[vkey]) {
								result[vkey] = true
								findlist.push(vkey)
							}
						}
					}
				}
				if (findlist.length) {
					const uniqueArray = [...new Set(findlist)];
					var textlist = uniqueArray.map(e => { return type_map[e]})
					oRun.AddComment(textlist.join(','), "biyueFix", `repair${pid}:${uniqueArray.join('-')}`)
					return true
				}
				return false
			}
			function hasBookmark(oParagraph) {
				var elementcount = oParagraph.GetElementsCount()
				if (elementcount > 0) {
					for (var i = elementcount - 1; i >= 0; --i) {
						var oElement = oParagraph.Paragraph.Content[i]
						if (oElement.GetType && oElement.GetType() == 71) { // 书签
							return true
						}
					}
				}
				return false
			}
			for (var i = 0, imax = paragraphs.length; i < imax; ++i) {
				var oParagraph = paragraphs[i]
				bflag = []
				var controls = oParagraph.GetAllContentControls() || []
				controls.forEach(oControl => {
					if (oControl.GetClassType() == 'inlineLvlSdt') {
						var count1 = oControl.GetElementsCount()
						if (count1) {
							for (var k = 0; k < oControl.GetElementsCount(); ++k) {
								if (handleRun(oControl.GetElement(k), oParagraph.Paragraph.Id)) {
									k += 3
								}
							}
						}
					}
				})
				var pcount = oParagraph.GetElementsCount()
				for (var j = 0; j < oParagraph.GetElementsCount(); ++j) {
					if (oParagraph.GetElementsCount() > pcount * 4) {
						console.log('== 超出了4倍, 死循环了')
						break
					}
					var addrev = handleRun(oParagraph.GetElement(j), oParagraph.Paragraph.Id)
					if (addrev) {
						j += 4
					}
				}
				// 判断是否为白色背景
				var paraPr = oParagraph.GetParaPr()
				if (paraPr && paraPr.ParaPr && isWhite(paraPr.ParaPr.Shd)) {
					result.hasWhiteBg = true
					oParagraph.AddComment(type_map['hasWhiteBg'], "biyueFix", `repair${oParagraph.Paragraph.Id}}:hasWhiteBg`)
				}
				if (!result.hasBookmark && hasBookmark(oParagraph)) {
					result.hasBookmark = true
				}
			}
			return result
		}, false, false)
	}).then(res => {
		Asc.scope.layout_detect_result = res
		window.biyue.showDialog('layoutRepairWindow', '字符检测', 'layoutRepair.html', 250, 400, false)
	})
}

// 排版修复
function layoutRepair(cmdData) {
	Asc.scope.cmdData = cmdData
	return biyueCallCommand(window, function() {
		var cmdData = Asc.scope.cmdData
		var oDocument = Api.GetDocument()
		var oRange = null
		if (Asc.scope.layout_all_range) {
			oRange = oDocument.GetRange()
		} else {
			oRange = oDocument.GetRangeBySelect()
			if (!oRange) {
				var currentContentControl = oDocument.Document.GetContentControl()
				if (currentContentControl) {
					oRange = Api.LookupObject(currentContentControl.Id).GetRange()
				}
			}
		}
		if (!oRange) {
			return null
		}
		var allComments = oDocument.GetAllComments() || []
		var paragrahs = oRange.GetAllParagraphs() || []
		var oneSpaceWidth = 2.11 // 一个空格的宽度
		var bflag = []
		var type_map = {
			'hasSmallImage': '宽高过小的图片',
			'has160': '不间断空格',
			'hasTab': '括号里使用了tab',
			'has32': '空格下划线',
			'has65307': '中文分号',
			'has12288': '中文空格',
			'hasWhiteBg': '段落背景为白色',
			'hasBookmark': '存在书签'
		}
		function getTabReplaceTarget(width, target) {
			var str = ''
			var count = Math.ceil(width / oneSpaceWidth)
			for (var i = 0; i < count; ++i) {
				str += target
			}
			return str
		}
		// 将tab下划线替换为真的下划线
		function replaceUnderline(oRun, parent, pos) {
			var fixed = false
			if (!oRun || oRun.GetClassType() != 'run') {
				return fixed
			}
			var runContent = oRun.Run.Content || []
			if (!oRun.GetUnderline()) {
				return fixed
			}
			var find = false
			var split = false
			for (var k = 0; k < runContent.length; ++k) {
				var len = runContent.length
				var element2 = runContent[k]
				var elementType = element2.GetType()
				if (elementType == 21 || elementType == 2) {
					if (!find && k > 0) {
						var newRun = oRun.Run.Split_Run(k + 1)
						parent.Add_ToContent(pos + 1, newRun)
						oRun.Run.RemoveElement(element2)
						var str = elementType == 2 ? '_' : getTabReplaceTarget(element2.Width, '_')
						oRun.Run.AddText(str, k)
						fixed = true
						return fixed
					}
					find = true
					oRun.Run.RemoveElement(element2)
					var str = elementType == 2 ? '_' : getTabReplaceTarget(element2.Width, '_')
					oRun.Run.AddText(str, k)
					var addCount = str.length - 1
					if (k < len - 1) {
						var nextType = runContent[k + 1 + addCount].GetType()
						if (nextType != 2 && nextType != 21) {
							oRun.SetUnderline(false)
							var newRun = oRun.Run.Split_Run(k + 1 + addCount)
							newRun.SetUnderline(true)
							parent.Add_ToContent(pos + 1, newRun)
							split = true
							fixed = true
							break
						} else {
							k += addCount
						}
					}
				}
			}
			if (find) {
				if (!split) {
					oRun.SetUnderline(false)
					fixed = true
				}
			}
			return fixed
		}
		function isWhite(shd) {
			return shd && shd.Fill && shd.Fill.r == 255 && shd.Fill.g == 255 && shd.Fill.b == 255 && shd.Fill.Auto == false
		}
		function handleRun(oRun, parent, pos) {
			var fixed = false
			if (!oRun || oRun.GetClassType() != 'run') {
				return fixed
			}
			if (cmdData.type == 1 && cmdData.value == 'whitebg') {
				var textpr = oRun.GetTextPr()
				if (textpr && textpr.TextPr) {
					var shd = textpr.TextPr.Shd
					if (isWhite(shd)) {
						oRun.Run.Set_Shd(undefined);
						fixed = true
					}
				}
				return fixed
			}
			if (cmdData.type == 1 && cmdData.value == 32) {
				fixed = replaceUnderline(oRun, parent, pos)
				return fixed
			}
 			var runContent = oRun.Run.Content || []
			for (var k = 0; k < runContent.length; ++k) {
				var element2 = runContent[k]
				if (element2.GetType() == 22) { // drawing
					if (cmdData.type == 2 && cmdData.value == 'smallimg') {
						if (element2.Width <= 1 && element2.Height <= 1) {
							element2.PreDelete();
							oRun.Run.RemoveElement(element2)
							fixed = true
						}
					}
				}
				if (cmdData.type == 1 && cmdData.value != 32) {
					if (element2.Value == cmdData.value) {
						oRun.Run.RemoveElement(element2)
						if (cmdData.newValue == 32) {
							var replaceStr = cmdData.value == 12288 ? '  ' : ' '
							oRun.Run.AddText(replaceStr, k)
							k += replaceStr.length - 1
							fixed = true
						} else if (cmdData.newValue == 59) {
							oRun.Run.AddText(';', k)
							fixed = true
						}
					}
				}
			}
			return fixed
		}
		function hasRightBracket2(begin, runContents) {
			for (var k = begin; k < runContents.length; ++k) {
				if (runContents[k].GetType() == 1 && (runContents[k].Value == 41 || runContents[k].Value == 65289)) {
					return true
				}
			}
		}
		// 之后是否存在右括号
		function hasRightBracket(oParagraph, iElement, jRun, idxParent) {
			for (var j = iElement; j < oParagraph.GetElementsCount(); ++j) {
				var oElement = oParagraph.GetElement(j)
				if (oElement.GetClassType() == 'run') {
					var runContents = oElement.Run.Content || []
					var begin = iElement == j ? jRun : 0
					if (hasRightBracket2(begin, runContents)) {
						return true
					}
				} else if (oElement.GetClassType() == 'inlineLvlSdt') {
					var count2 = oElement.GetElementsCount()
					var begin = iElement == j ? idxParent : 0
					for (var idx = begin; idx < count2; ++idx) {
						if (oElement.GetElement(idx).GetClassType() == 'run') {
							var runContents = oElement.GetElement(idx).Run.Content || []
							var begin = iElement == j ? idxParent : 0
							if (hasRightBracket2(begin, runContents)) {
								return true
							}
						}
					}
				}
			}
		}
		function handleRun2(oParagraph, j, oElement, bflag, idxParent) {
			var runContents = oElement.Run.Content || []
			var fixed = false
			for (var k = 0; k < runContents.length; ++k) {
				var type = runContents[k].GetType()
				if (type == 1) {
					var value = runContents[k].Value
					if (value == 65288 || value == 40) { // 左括号
						bflag.push(1)
					} else if (value == 65289 || value == 41) { // 右括号
						bflag.pop()
					}
				} else if (type == 21 && bflag.length) {
					// 段落之后是否存在右括号
					var findRight = hasRightBracket(oParagraph, j, k, idxParent)
					if (findRight) {
						var element2 = runContents[k]
						oElement.Run.RemoveElement(element2)
						var str = getTabReplaceTarget(element2.Width, ' ')
						oElement.Run.AddText(str, k)
						k += str.length - 1
						fixed = true
					}
				}
			}
			return fixed
		}
		function removeComment(pid) {
			var cid = 'repair' + pid
			for (var c = 0; c < allComments.length; ++c) {
				var oComment = allComments[c]
				if (oComment.GetAutorName() != 'biyueFix') {
					continue
				}
				var userid = oComment.GetUserId()
				var arr1 = userid.split(':')
				if (arr1.length == 2 && cid == arr1[0]) {
					var arr2 = arr1[1].split('-')
					var index = arr2.findIndex(e2 => {
						return e2 == cmdData.keyname
					})
					if (index >= 0) {
						if (arr2.length == 1) {
							oComment.Delete()
						} else {
							arr2.splice(index, index + 1)
							var newText = arr2.map(e => { return type_map[e]})
							oComment.SetText(newText.join(','))
							oComment.SetUserId(`${arr1[0]}:${arr2.join('-')}`)
						}
					}
				}
			}
		}
		// 删除书签
		function removeBookmark(oParagraph) {
			var elementcount = oParagraph.GetElementsCount()
			if (elementcount > 0) {
				for (var i = elementcount - 1; i >= 0; --i) {
					var oElement = oParagraph.Paragraph.Content[i]
					if (oElement.GetType && oElement.GetType() == 71) { // 书签
						oParagraph.RemoveElement(i)
						++i
					}
				}
			}
		}
		if (cmdData.type == 1 && cmdData.value == 'tab') { // 将括号里的tab替换为空格
			var bflag = []
			for (var i = 0, imax = paragrahs.length; i < imax; ++i) {
				var oParagraph = paragrahs[i]
				var fixed = false
				for (var j = 0; j < oParagraph.GetElementsCount(); ++j) {
					var oElement = oParagraph.GetElement(j)
					if (oElement.GetClassType() == 'run') {
						if (handleRun2(oParagraph, j, oElement, bflag)) {
							fixed = true
						}
						
					} else if (oElement.GetClassType() == 'inlineLvlSdt') {
						var count2 = oElement.GetElementsCount()
						for (var idx = 0; idx < count2; ++idx) {
							if (oElement.GetElement(idx).GetClassType() == 'run') {
								if (handleRun2(oParagraph, j, oElement.GetElement(idx), bflag)) {
									fixed = true
								}
								
							}
						}
					}
				}
				if (fixed) {
					removeComment(oParagraph.Paragraph.GetId())
				}
			}
		} else {
			for (var i = 0, imax = paragrahs.length; i < imax; ++i) {
				var oParagraph = paragrahs[i]
				var fixed = false
				if (cmdData.type == 1 && cmdData.value == 'whitebg') {
					var oParaPr = oParagraph.GetParaPr();
					if (oParaPr) {
						if (oParaPr && oParaPr.ParaPr && isWhite(oParaPr.ParaPr.Shd)) {
							oParaPr.SetShd("clear", 255, 255, 255, true);
							fixed = true
						}
					}
				} else if (cmdData.type == 2 && cmdData.value == 'bookmark') { // 删除书签
					removeBookmark(oParagraph)
				}
				var controls = oParagraph.GetAllContentControls() || []
				controls.forEach(oControl => {
					if (oControl.GetClassType() == 'inlineLvlSdt') {
						var count1 = oControl.GetElementsCount()
						if (count1) {
							for (var k = 0; k < oControl.GetElementsCount(); ++k) {
								if (handleRun(oControl.GetElement(k), oControl.Sdt, k)) {
									fixed = true
								}
							}
						}
					}
				})
				for (var j = 0; j < oParagraph.GetElementsCount(); ++j) {
					if (handleRun(oParagraph.GetElement(j), oParagraph.Paragraph, j)) {
						fixed = true
					}
				}
				if (fixed) {
					removeComment(oParagraph.Paragraph.GetId())	
				}
			}
		}
	}, false, true)
}

function removeAllComment() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var allComments = oDocument.GetAllComments() || []
		allComments.forEach(oComment => {
			if (oComment.GetAutorName() == 'biyueFix') {
				oComment.Delete()
			}
		})
	}, false, false)
}

export {
	layoutDetect,
	layoutRepair,
	removeAllComment
}