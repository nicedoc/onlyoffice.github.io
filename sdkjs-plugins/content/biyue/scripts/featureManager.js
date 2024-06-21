import { map_base64 } from '../resources/list_base64.js'
import { ZONE_TYPE, ZONE_SIZE } from './model/feature.js'

var loading = false // 正在绘制中
var list_command = [] // 操作列表
var list_wait_command = [] // 等待执行的操作列表
let setupPostTask = function(window, task) {
  window.postTask = window.postTask || [];
  window.postTask.push(task);
}

function handleFeature(options) {
  if (options.zone_type == ZONE_TYPE.QRCODE) {
    options.url = map_base64.qrcode
  } else if (options.zone_type == ZONE_TYPE.AGAIN) {
    options.text = '再练'
  } else if (options.zone_type == ZONE_TYPE.PASS) {
    options.text = '通过'
  } else if (options.zone_type == ZONE_TYPE.IGNORE) {
    options.text = '日期/评语'
  }
  options.size = ZONE_SIZE[options.zone_type]
  if (options.v == undefined) {
    options.v = 1
  }
  options.page_num = options.p || 0
  options.type = 'feature'
  addCommand(options)
  if (loading) {
    console.log('loading...')
    return
  }
  drawList([options])
}
// 整理参数
function addCommand(options) {
  if (!list_wait_command) {
    list_wait_command = []
  }
  var index = list_wait_command.findIndex(e => {
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
      drawHeader(newlist[newlist.length - 1].cmd, newlist[newlist.length - 1].title)
    }
  }
}

function handleHeader(cmdType, examTitle) {
  addCommand({
    type: 'header',
    title: examTitle,
    cmd: cmdType
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
  setupPostTask(window, function (res) {
    handleNext()
  })
  window.Asc.plugin.callCommand(function() {
    var cmdType = Asc.scope.header_cmd
    var examTitle = Asc.scope.header_exam_title
    var qrcode_url = Asc.scope.qrcode_url
    var MM2TWIPS = (25.4 / 72 / 20)
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
      var oStroke = Api.CreateStroke(3600, Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125)))
      var oDrawing = Api.CreateShape("rect", 42 * scale * 36E3, 24 * scale * 36E3, oFill, oStroke)
      var drawDocument = oDrawing.GetContent()
      var paragraphs = drawDocument.GetAllParagraphs()
      if (paragraphs && paragraphs.length > 0) {
        var oRun = Api.CreateRun();
        oRun.AddText('再练')
        paragraphs[0].AddElement(oRun)
        paragraphs[0].SetColor(3, 3, 3, false)
        paragraphs[0].SetFontSize(14)
        paragraphs[0].SetJc('center')
      }
      oDrawing.SetVerticalTextAlign("center")
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
      var oParagraph = oHeader.GetElement(0);
      oParagraph.SetTabs([1, pw * 0.5 / MM2TWIPS, pw / MM2TWIPS], ["left", "center", "right"]);
      oParagraph.AddTabStop();
      if (showAgain) {
        oParagraph.AddDrawing(createAgainDrawing())
      }
      oParagraph.AddTabStop();
      oParagraph.AddText(examTitle)
      oParagraph.AddTabStop();
      var oDrawing2 = Api.CreateImage(qrcode_url, 45 * scale * 36E3, 45 * scale * 36E3)
      oParagraph.AddDrawing(oDrawing2)
      oParagraph.SetBottomBorder("single", 8, 3, 153, 153, 153);
    }
    var firstSection = oSections[0]
    firstSection.SetTitlePage(true);
    var oHeader = firstSection.GetHeader("title", true)
    setHeader(firstSection, oHeader, true)

    if (oSections.length > 1) {
      var oHeader2 = oSections[1].GetHeader("default", true)
      setHeader(oSections[1], oHeader2, false)
    }
  }, false, true, undefined)
}

function drawList(list) {
  loading = true
  setupPostTask(window, function (res) {
    loading = false
    console.log('drawList result:', res)
    if (res && res.list) {
      var pos_list = window.BiyueCustomData.pos_list || []
      res.list.forEach(result => {
        var index = pos_list.findIndex(e => {
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
            y: result.y
          })
        }
      })
      window.BiyueCustomData.pos_list = pos_list
    }
    handleNext()
  })
  Asc.scope.feature_wait_handle = list
  Asc.scope.pos_list = window.BiyueCustomData.pos_list
  Asc.scope.ZONE_TYPE = ZONE_TYPE
  window.Asc.plugin.callCommand(function () {
    var pos_list = Asc.scope.pos_list || []
    var ZONE_TYPE = Asc.scope.ZONE_TYPE
    var MM2TWIPS = (25.4 / 72 / 20)
    var oDocument = Api.GetDocument()
    var objs = oDocument.GetAllDrawingObjects()
    var feature_wait_handle = Asc.scope.feature_wait_handle
    var res = {
      list: []
    }
    var oSections = oDocument.GetSections()
    function addImageToCell(oTable, nRow, nCell, url, w, h, mleft, mright) {
      var cell = oTable.GetCell(nRow, nCell)
      if (!cell) {
        return
      }
      var oCellContent = cell.GetContent()
      if (oCellContent) {
        var p = oCellContent.GetElement(0)
        if (p && p.GetClassType() == "paragraph") {
          var oImage = Api.CreateImage(url, w * 36E3, h * 36E3)
          p.AddDrawing(oImage)
        }
      }
      cell.SetWidth("twips", w / MM2TWIPS)
      cell.SetVerticalAlign("center")
      if (mleft != undefined) {
        cell.SetCellMarginLeft(mleft)
      }
      if (mright != undefined) {
        cell.SetCellMarginRight(mright)
      }
      cell.SetCellMarginTop(0)
      cell.SetCellMarginBottom(0)

    }
    function addTextToCell(oTable, nRow, nCell, texts, jc, w, fontSize, mleft, mright) {
      var cell = oTable.GetCell(nRow, nCell)
      var oCellContent = cell.GetContent()
      if (oCellContent) {
        var p = oCellContent.GetElement(0)
        if (p && p.GetClassType() == "paragraph") {
          var oRun = Api.CreateRun();
          for (var i = 0; i < texts.length; ++i) {
            oRun.AddText(texts[i]);
            if (i < texts.length - 1) {
              oRun.AddLineBreak();
            }
          }
          p.AddElement(oRun);
          p.SetFontSize(fontSize)
          p.SetJc(jc)
          p.SetColor(33, 33, 33, false)
        }
      }
      cell.SetVerticalAlign("center")
      cell.SetCellMarginLeft(0)
      cell.SetCellMarginRight(0)
      cell.SetWidth("twips", w / MM2TWIPS)
      if (mleft != undefined) {
        cell.SetCellMarginLeft(mleft)
      }
      if (mright != undefined) {
        cell.SetCellMarginRight(mright)
      }
    }
    feature_wait_handle.forEach(options => {
      var pos_data = pos_list.find(pos => {
        return pos.zone_type == options.zone_type && pos.v == options.v
      })
      var oDrawing = null
      if (pos_data && pos_data.drawing_id) { // 已存在
        var index = objs.findIndex(obj => {
          return obj.Drawing.Id == pos_data.drawing_id
        })
        if (index >= 0) {
          oDrawing = objs[index]
        }
      }
      var result = {
        code: 0,
        zone_type: options.zone_type,
        v: options.v,
        cmd: options.cmd
      }
      if (options.cmd == 'close') { // 关闭
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
            if (options.zone_type == ZONE_TYPE.QRCODE) {
              shapeWidth = 32
              shapeHeight += 2
            }
            // if (options.zone_type == ZONE_TYPE.AGAIN) {
            //   var firstSection = oSections[0]
            //   var oHeader = firstSection.GetHeader("default", true);
            //   var oParagraph = oHeader.GetElement(0);
            //   console.log('oHeader', oHeader)
            //   console.log('oHeader', oParagraph)
            // }
            if (options.zone_type == ZONE_TYPE.QRCODE) {
              var oTable = Api.CreateTable(2, 1)
              addImageToCell(oTable, 0, 0, options.url, options.size.w, options.size.h)
              oTable.GetCell(0, 0).SetCellBorderRight("single", 1, 0.1, 255, 255, 255)
              oTable.SetTableBorderTop("single", 1, 0.1, 255, 255, 255)
              oTable.SetTableBorderBottom("single", 1, 0.1, 255, 255, 255)
              oTable.SetTableBorderLeft("single", 1, 0.1, 255, 255, 255)
              oTable.SetTableBorderRight("single", 1, 0.1, 255, 255, 255)
              
              addTextToCell(oTable, 0, 1, ["编码-未生成", "吕老师"], 'center', (shapeWidth - options.size.w - 2), 14)
              oTable.SetWidth('twips', shapeWidth / MM2TWIPS)

              var oFill = Api.CreateNoFill()
              var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
              oDrawing = Api.CreateShape("rect", shapeWidth * 36E3, shapeHeight * 36E3, oFill, oStroke)
              var oDrawingContent = oDrawing.GetContent()
              oDrawingContent.AddElement(0, oTable)
            } else if (options.zone_type == ZONE_TYPE.THER_EVALUATION) {
              var oTable = Api.CreateTable(13, 1)
              var scale = 0.25
              var flowersize = 24
              var fw = flowersize * scale
              var fh = flowersize * scale
              var textw = 24
              var w = (21.33 * scale + textw + fw * 4) * 2 + 10
              var furl = 'https://eduteacher.xmdas-link.com/online_editor/static/img_20240614165159/flower.png'
              var fmargin = 1 / MM2TWIPS
              addImageToCell(oTable, 0, 0, 'https://eduteacher.xmdas-link.com/online_editor/static/xiaoyue.png', 21.33 * scale, 30 * scale, 0, 0)
              addTextToCell(oTable, 0, 1, ["测试评价:"], 'center', textw, 20, 0, fmargin)
              addImageToCell(oTable, 0, 2, furl, fw, fh, fmargin, fmargin)
              addImageToCell(oTable, 0, 3, furl, fw, fh, fmargin, fmargin)
              addImageToCell(oTable, 0, 4, furl, fw, fh, fmargin, fmargin)
              addImageToCell(oTable, 0, 5, furl, fw, fh, fmargin, fmargin)
              oTable.GetCell(0, 6).SetWidth('twips', 10 / MM2TWIPS)
              addImageToCell(oTable, 0, 7, 'https://eduteacher.xmdas-link.com/online_editor/static/xiaotao.png', 21.33 * scale, 30 * scale, 0, 0)
              addTextToCell(oTable, 0, 8, ["教师评价:"], 'center', textw, 20, 0, fmargin)
              addImageToCell(oTable, 0, 9, furl, fw, fh, fmargin, fmargin)
              addImageToCell(oTable, 0, 10, furl, fw, fh, fmargin, fmargin)
              addImageToCell(oTable, 0, 11, furl, fw, fh, fmargin, fmargin)
              addImageToCell(oTable, 0, 12, furl, fw, fh, fmargin, fmargin)

              oTable.SetWidth('twips', w / MM2TWIPS)
              oTable.GetRow(0).SetHeight("auto", shapeHeight / MM2TWIPS)
              console.log('w', w)
              
              shapeWidth = w + 1
              var oFill = Api.CreateNoFill()
              var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
              oDrawing = Api.CreateShape("rect", shapeWidth * 36E3, shapeHeight * 36E3, oFill, oStroke)
              var oDrawingContent = oDrawing.GetContent()
              oDrawingContent.AddElement(0, oTable)
              oDrawing.SetPaddings(0, 0, 0, 0)
            } else if (options.text) {
              var oFill = Api.CreateNoFill()
              var oStroke = Api.CreateStroke(options.size.stroke_width * 36E3, Api.CreateSolidFill(Api.CreateRGBColor(153,153,153)));
              oDrawing = Api.CreateShape(options.size.shape_type, shapeWidth * 36E3, shapeHeight * 36E3, oFill, oStroke)
              var drawDocument = oDrawing.GetContent()
              var paragraphs = drawDocument.GetAllParagraphs()
              if (paragraphs && paragraphs.length > 0) {
                var oRun = Api.CreateRun();
                oRun.AddText(options.text);
                paragraphs[0].AddElement(oRun)
                paragraphs[0].SetColor(153, 153, 153, false)
                paragraphs[0].SetFontSize(options.size.font_size)
                paragraphs[0].SetJc(options.size.jc || 'center')
              }
              oDrawing.SetVerticalTextAlign("center")
              console.log('oDrawing', oDrawing)
            }
            oDrawing.Drawing.Set_Props({
              title: 'feature'
            })
            oDocument.AddDrawingToPage(oDrawing, options.page_num, options.x * 36E3, options.y * 36E3)
          }
        } else {
          oDrawing.SetHorPosition('page', options.x * 36E3)
          oDrawing.SetVerPosition('page', options.y * 36E3)
        }
        result.code = 1
        result.x = options.x
        result.y = options.y
        result.drawing_id = oDrawing.Drawing.Id
      }
      res.list.push(result)
    });
    return res
  }, false, true, undefined)
}

export {
  handleFeature,
  handleHeader
}