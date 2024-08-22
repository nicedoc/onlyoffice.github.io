// 主要用于处理业务逻辑

import { JSONPath } from '../vendor/jsonpath-plus/dist/index-browser-esm.js';
import { biyueCallCommand } from './command.js';
let paper_info = {} // 从后端返回的试卷信息
const MM2EMU = 36000 // 1mm = 36000EMU
// 根据paper_uuid获取试卷信息
async function initPaperInfo() {
  console.log('============== initPaperInfo')
  await paperOnlineInfo(window.BiyueCustomData.paper_uuid).then(res => {
    paper_info = res.data
    window.BiyueCustomData.exam_title = paper_info.paper.title
    console.log('试卷信息')
    console.log(paper_info)
    updateCustomControls()
  }).catch(res => {
    console.log(res)
    updateCustomControls()
  })
}

function paperOnlineInfo() {}
function structAdd() {} 
function questionCreate() {}
function questionDelete() {}
function questionUpdateContent() {}
function structDelete() {}
function paperCanConfirm() {}
function structRename() {}
function paperSavePosition(){}
function examQuestionsUpdate(){}

function getPaperInfo() {
  return paper_info
}

function p_Twips2MM(twips) {
  return 25.4 / 72 / 20 * twips
}
function p_MM2Twips(mm) {
  return mm / (25.4 / 72 / 20)
}
function p_EMU2MM(EMU) {
  return EMU / 36E3
}
function p_MM2EMU(mm) {
  return mm * 36E3
}

// 更新customData的control_list
function updateCustomControls() {
  Asc.scope.paper_info = paper_info
  return biyueCallCommand(window, function() {
    let paperinfo = Asc.scope.paper_info
    var oDocument = Api.GetDocument();
    let controls = oDocument.GetAllContentControls();
    let ques_no = 1
    let struct_index = 0
    var control_list = []
    controls.forEach(control => {      
      //var rect = Api.asc_GetContentControlBoundingRect(control.Sdt.GetId(), true);
      let tagInfo = JSON.parse(control.GetTag())
      var text = control.GetRange().GetText()
      let obj = {
        control_id: control.Sdt.GetId(),
        regionType: tagInfo.regionType,
        text: text
      }
      if (tagInfo.regionType == 'struct') {
        ++struct_index
        const pattern = /^[一二三四五六七八九十0-9]+.*?(?=[：:])/
        const result = pattern.exec(text);
        obj.name = result ? result[0] : null
        if (paperinfo.ques_struct_list && (struct_index - 1) < paperinfo.ques_struct_list.length) {
          obj.struct_id = paperinfo.ques_struct_list[struct_index - 1].struct_id
        }
      } else if (tagInfo.regionType == 'question') {
        obj.ques_no = ques_no
        const regex = /^([^.．]*)/;
        const match = obj.text.match(regex);
        obj.ques_name = match ? match[1] : ''
        if (paperinfo.info && paperinfo.info.questions && ques_no <= paperinfo.info.questions.length) {
          obj.score = paperinfo.info.questions[ques_no - 1].score
          obj.ques_uuid = paperinfo.info.questions[ques_no - 1].uuid
          obj.ask_controls = []
          if (paperinfo.ques_struct_list && (struct_index - 1) < paperinfo.ques_struct_list.length) {
            obj.struct_id = paperinfo.ques_struct_list[struct_index - 1].struct_id
          }
        }
        ques_no++
      } else if (tagInfo.regionType == 'write' || tagInfo.regionType == 'sub-question') {
        let parentContentControl = control.GetParentContentControl()
        obj.parent_control_id = parentContentControl.Sdt.GetId()
        function getParentQues(id) {
          for (var i = 0, imax = control_list.length; i < imax; ++i) {
            if (control_list[i].control_id == id) {
              if (control_list[i].regionType == 'question') {
                return control_list[i]
              } else {
                return getParentQues(control_list[i].parent_control_id)
              }
            }
          }
          return null
        }
        var parentQues = getParentQues(obj.parent_control_id)
        if (parentQues) {
          obj.ques_uuid = parentQues.ques_uuid
          obj.parent_ques_control_id = parentQues.control_id
          if(tagInfo.regionType == 'write') {
            if (!parentQues.ask_controls) {
              parentQues.ask_controls = []
            }
            parentQues.ask_controls.push({
              control_id: control.Sdt.GetId(),
              v: 0
            })
          } else {
            if (!parentQues.sub_questions) {
              parentQues.sub_questions = []
            }
            parentQues.sub_questions.push({
              control_id: control.Sdt.GetId()
            })
          }
        }
      } else if (tagInfo.regionType == 'feature') {
        obj.zone_type = tagInfo.zone_type
        obj.v = tagInfo.v
      }
      control_list.push(obj)
    })
    return control_list
  }, false, false);
}

// 清除试卷结构和所有题目
async function clearStruct() {
  console.log('开始清除结构')
  if (paper_info.info && paper_info.info.questions) {
    for (const e of paper_info.info.questions) {
      await questionDelete(window.BiyueCustomData.paper_uuid, e.uuid, 1)
    }
  }
  if (paper_info.ques_struct_list) {
    for (const estruct of paper_info.ques_struct_list) {
      await structDelete(window.BiyueCustomData.paper_uuid, estruct.struct_id)
    }
  }
  console.log('清除结构成功')
  await initPaperInfo()
}

let setupPostTask = function(window, task) {
  window.postTask = window.postTask || [];
  window.postTask.push(task);
}

// 获取试卷结构
async function getStruct() {
  if (!window.BiyueCustomData.control_list) {
    return
  }
  console.log('开始获取试卷结构')
  var struct_index = 0
  for (var control of window.BiyueCustomData.control_list) {
    if (control.regionType == 'struct') {
      ++struct_index
      if (paper_info.ques_struct_list && (struct_index - 1) < paper_info.ques_struct_list.length) {
        control.struct_id = paper_info.ques_struct_list[struct_index - 1].struct_id
      }
      if (!control.struct_id) {
        await structAdd({
          paper_uuid: window.BiyueCustomData.paper_uuid,
          name: control.name,
          rich_name: control.text
        }).then(res => {
          if (res.code == 1) {
            control.struct_id = res.data.struct_id
            // control.struct_name = control.name
          }
        })
      } else if (control.name != paper_info.ques_struct_list[struct_index - 1].struct_name) {
        await structRename(window.BiyueCustomData.paper_uuid, control.struct_id, control.name)
      }
    }
  }
  await getQuesUuid()
}

async function getQuesUuid() {
  for (const e of window.BiyueCustomData.control_list) {
    if (e.regionType == 'question') {
      var uuid = ''
      if (paper_info.info && paper_info.info.questions) {
        var ques = paper_info.info.questions.find(item => {
          return item.no == e.ques_no
        })
        if (ques) {
          uuid = ques.uuid
        }
      }
      if (uuid == '') {
        await questionCreate({
          paper_uuid: window.BiyueCustomData.paper_uuid,
          content: encodeURIComponent(e.text), // e.text,
          blank: '',
          type: 1,
          score: 0,
          no: e.ques_no,
          struct_id: e.struct_id
        })
      } else {
        await questionUpdateContent({
          paper_uuid: window.BiyueCustomData.paper_uuid,
          question_uuid: uuid,
          content: encodeURIComponent(e.text), // e.text,
        })
      }
    }
  }
  console.log('所有结构和题目都更新完')
  await initPaperInfo()
}
// 保存位置信息
function savePositons() {
  Asc.scope.customData = window.BiyueCustomData
  window.Asc.plugin.callCommand(function () {
    var oDocument = Api.GetDocument();
    let controls = oDocument.GetAllContentControls();
    const isPageCoord = true;
    let mmToPx = function(mm) {
      // 1 英寸 = 25.4 毫米
      // 1 英寸 = 96 像素（常见的屏幕分辨率）
      // 因此，1 毫米 = (96 / 25.4) 像素
      const pixelsPerMillimeter = 96 / 25.4;
      return Math.floor(mm * pixelsPerMillimeter);
    }
    var list = []
    controls.forEach((control) => {
      let control_id = control.Sdt.GetId()
      let rect_format = {}
      let rect = Api.asc_GetContentControlBoundingRect(control_id, isPageCoord)
      rect_format = {
        page: rect.Page ? rect.Page + 1 : 1,
        x: mmToPx(rect.X0),
        y: mmToPx(rect.Y0),
        w: mmToPx(rect.X1 - rect.X0),
        h: mmToPx(rect.Y1 - rect.Y0)
      }
      list.push({
        rect: rect_format,
        control_id: control.Sdt.GetId()
      })
    })
    return list
    }, false, false, function(list) {
      var customData = window.BiyueCustomData
      var questionPositions = {}
      customData.control_list.forEach((e) => {
        if (e.regionType == 'question') {
          var rectInfo = list.find(item => {
            return item.control_id == e.control_id
          })
          var title_region = [rectInfo.rect]
          var mark_ask_region = {}
          if (e.ask_controls && e.ask_controls.length > 0) {
            e.ask_controls.forEach((ask, askindex) => {
              var askfind = list.find(askItm => () => {
                return askItm.control_id == ask.control_id
              })
              if (askfind) {
                mark_ask_region[(askindex + 1) + ''] = [Object.assign({}, askfind.rect, {
                  v: ask.v + '',
                  order: (askindex + 1) + '' 
                })]
              }
            })
          } else {
            mark_ask_region = {
              "1": [Object.assign({}, rectInfo.rect, {
                v: e.score + '',
                order: "1"
              })]
            }
          }
          var write_ask_region = [Object.assign({}, rectInfo.rect, {
            v: e.score + '',
            order: "1"
          })]


          questionPositions[e.ques_uuid] = {
            ques_type:1,
            itemId:3,
            ques_no: e.ques_no,
            ques_name: e.ques_name,
            content: encodeURIComponent(e.text), // 网络请求时提示（无法将值解码），改成'111'后，接口可以正常调用
            score: e.score,
            ref_id: e.uuid,
            answer: "",
            additional:false,
            mark_method: 1, // 标错还是打分
            title_region: title_region,
            correct_region: {},
            grade_positions:[],
            mark_ask_region: mark_ask_region,
            write_ask_region: write_ask_region,
            ask_num: e.ask_controls ? (e.ask_controls.length || 1) : 1
          }
        }
      })
      console.log('positions:', questionPositions)
      paperSavePosition(customData.paper_uuid, questionPositions, '', '').then(res => {
        console.log('保存位置成功')
      }).catch(error => {
        console.log(error)
      })
    });
}

function showQuestionTree() {
  const questionList = $('#questionList');
  if (questionList) {
    if (questionList.children() && questionList.children().length > 0 && questionList.children().length < 2) {
      questionList.toggle()
    } else {
      initTree()
    }
  }
}

function initTree() {
  const questionList = $('#questionList');
  var needRefreshStruct = false
  if (!paper_info.info) {
    needRefreshStruct = true
  }
  if (!needRefreshStruct && window.BiyueCustomData.control_list && window.BiyueCustomData.control_list.length > 0) {
    var find = window.BiyueCustomData.control_list.find(e => {
      return (e.regionType == 'struct' && e.struct_id == 0) || (e.regionType == 'question' && !e.ques_uuid)
    })
    if (find) {
      
      needRefreshStruct = true
    }
  }
  if (needRefreshStruct) {
    questionList.append('<div style="color:#ff0000">试卷结构有异，请先更新试卷结构</div>')
    return
  }
  if (!window.BiyueCustomData.control_list || window.BiyueCustomData.control_list.length == 0) {
    questionList.append('<div><button id="refreshTree">刷新树</button></div>')
    $('#refreshTree').on('click', () => {
      if (window.BiyueCustomData.control_list && window.BiyueCustomData.control_list.length > 0) {
        initTree()
      }
    })
    return
  }
  questionList.append('<div><button id="addfield">添加作答题分数框</button></div>')
  var question_types = [
    { value: '0', label: '未定义' },
    { value: '1', label: '单选' },
    { value: '2', label: '填空' },
    { value: '3', label: '作答' },
    { value: '4', label: '判断' },
    { value: '5', label: '多选' },
    { value: '6', label: '文本' },
    { value: '7', label: '单选组合' },
    { value: '8', label: '作文' }
  ];
  var structItem
  var structList
  window.BiyueCustomData.control_list.forEach((control, index) => {
    if (control.regionType == 'struct') {
      if (structList) {
        structItem.append(structList);
        questionList.append(structItem);  
      }
      structItem = $('<div>').text(control.name);
      structList = $('<ul>');
      const select = $('<select>');
      question_types.forEach(type => {
        const option = $('<option>').val(type.value).text(type.label);
        select.append(option);
      });

      select.val('0');

      select.on('change', function() {
        changeStructQuesType(control.struct_id, parseInt(select.val()))
      });
      
      structItem.append(select);

    } else if (control.regionType == 'question') {
      const questionItem = $('<div>');
      const select = $('<select>');
      const input = $('<input>');
      const unittext = $('<span>') 

      question_types.forEach(type => {
        const option = $('<option>').val(type.value).text(type.label);
        select.append(option);
      });

      select.attr('id', `select${control.control_id}`);
      select.val('0');
      select.on('change', function() {
        control.ques_type = parseInt(select.val());
        console.log('window.BiyueCustomData.control_list', window.BiyueCustomData.control_list)
      });
      input.attr('id', `score${control.control_id}`);
      input.attr('type', 'text').val(control.score === '' ? '（未设置）' : control.score);
      input.on('input', function() {
        control.score = input.val();
        input.attr('placeholder', control.score === '' ? '（未设置）' : '');
      });
      
      questionItem.text(`${control.ques_name}.`);
      unittext.text('分')
      questionItem.append(select);
      questionItem.append(input);
      questionItem.append(unittext)
      structList.append(questionItem);
      let str = ''
      if (control.sub_questions && control.sub_questions.length) {
        str += `含${control.sub_questions.length}小题，`
      }
      if (control.ask_controls && control.ask_controls.length) {
        str += `${control.ask_controls.length}小问`
      }
      if (str != '') {
        structList.append(`<div style="color:#999">(${str})</div>`)
      }
      
    }
  });
  if (structList) {
    structItem.append(structList);
    questionList.append(structItem);  
  }
  $('#addfield').on('click', addScoreField)
}

function changeStructQuesType(struct_id, v) {
  console.log('changeStructQuesType', struct_id, v)
  window.BiyueCustomData.control_list.forEach(control => {
    if (control.regionType == 'question' && control.struct_id == struct_id) {
      control.ques_type = v
      $(`#select${control.control_id}`).val(v)
    }
  })
}
// 更新题目分数
function updateQuestionScore() {
  window.BiyueCustomData.control_list.forEach(control => {
    if (control.regionType == 'question') {
      console.log('control', control.ques_no, control.score)
      control.score = $(`#score${control.control_id}`).val(control.score)
    }
  })
}

function addDrawingObj() {
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument();
    var oFill = Api.CreateNoFill()
    var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
    var oStroke = Api.CreateStroke(3600, oFill2);
    // 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
    var oDrawing = Api.CreateShape("rect", 36000 * 20, 36000 * 10, oFill, oStroke);
    var drawDocument = oDrawing.GetContent()
    var oParagraph = Api.CreateParagraph();    
    var oMyStyle = oDocument.CreateStyle("field style")    
    var oTextPr = oMyStyle.GetTextPr()
    oTextPr.SetColor(0, 0, 0, false)
    oTextPr.SetFontSize(32)
    var oParaPr = oMyStyle.GetParaPr()
    oParaPr.SetJc("center")
    oParagraph.SetStyle(oMyStyle)
    oParagraph.AddText('再练')
    drawDocument.AddElement(0,oParagraph)
    oDrawing.SetVerticalTextAlign("center")
    oDocument.AddDrawingToPage(oDrawing, 0, 1070821, 963295);
  }, false, true, undefined)
}

// 测试在文档尾部加入表格，显示打分区
function testAddTable() {
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument();
    var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
    var num = 3
    var oTable = Api.CreateTable(num + 2, 1);
    oTable.SetWidth("percent", 20);
    var oTableStylePr = oTableStyle.GetConditionalTableStyle("wholeTable");
    oTable.SetTableLook(true, true, true, true, true, true);
    oTableStylePr.GetTableRowPr().SetHeight("atLeast", 500);
    var oTableCellPr = oTableStyle.GetTableCellPr();
    oTableCellPr.SetVerticalAlign("center");
    //oTable.SetHAlign("center") // 整个表格相对于段落的水平居中对齐
    oTable.SetStyle(oTableStyle);
    var oMyStyle = oDocument.CreateStyle("My style with center")
    var oParaPr = oMyStyle.GetParaPr()
    oParaPr.SetJc("center") // 设置段落内容对齐方式。
    for (var i = 0; i < num + 2; ++i) {
      var oCellPara = oTable.GetCell(0, i).GetContent().GetElement(0)
      if (i == num + 1) {
        oCellPara.AddText('0.5')
      } else {
        oCellPara.AddText(i + '')
      }
      oCellPara.SetStyle(oMyStyle)
    }
    
    oDocument.Push(oTable);
  }, false, true, undefined)
}

function addQuesScore(score = 10) {
  Asc.scope.score = score
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument();
    var controls = oDocument.GetAllContentControls()
    var score = Asc.scope.score
    if (controls) {
      for (var i = 0; i < controls.length; ++i) {
        var control = controls[i]
        var tag = JSON.parse(control.GetTag())
        if (tag.regionType == 'question') {
          var range = control.GetRange()
          var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
          var num = score
          var oTable = Api.CreateTable(num + 2, 1);
          oTable.SetWidth("percent", 20);
          var oTableStylePr = oTableStyle.GetConditionalTableStyle("wholeTable");
          oTable.SetTableLook(true, true, true, true, true, true);
          oTableStylePr.GetTableRowPr().SetHeight("atLeast", 500);
          var oTableCellPr = oTableStyle.GetTableCellPr();
          oTableCellPr.SetVerticalAlign("center");
          //oTable.SetHAlign("center") // 整个表格相对于段落的水平居中对齐
          oTable.SetStyle(oTableStyle);
          var oMyStyle = oDocument.CreateStyle("My style with center")
          var oParaPr = oMyStyle.GetParaPr()
          oParaPr.SetJc("center") // 设置段落内容对齐方式。
          for (var i = 0; i < num + 2; ++i) {
            var oCellPara = oTable.GetCell(0, i).GetContent().GetElement(0)
            if (i == num + 1) {
              oCellPara.AddText('0.5')
            } else {
              oCellPara.AddText(i + '')
            }
            oCellPara.SetStyle(oMyStyle)
          }
          if (range && range.StartPos && range.StartPos.length > 0) {
            var pos = range.StartPos[0].Position
            control.AddElement(oTable, 0)
            // oDocument.AddElement(pos, oTable)
          }
          break
        }
      }
    }
    
  }, false, true, function(res) {

  })
}

// 添加分数框
function addScoreField(score, mode, layout) {
  Asc.scope.control_list = window.BiyueCustomData.control_list
  Asc.scope.params = {
    score: score,
    mode: mode,
    layout: layout,
    scores: getScores(score, mode)
  }
  
  window.Asc.plugin.callCommand(function() {
    var control_list = Asc.scope.control_list
    var controls = Api.GetDocument().GetAllContentControls();
    var params = Asc.scope.params
    var score = params.score
    for (var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'question') {
        var control_id = control.Sdt.GetId()
        var controldata = control_list.find(item => {
          return item.regionType == 'question' && item.control_id == control_id
        })
        var rect = Api.asc_GetContentControlBoundingRect(control_id, true);
        var width = rect.X1 - rect.X0
        var trips_width = width / (25.4 / 72 / 20)
        if (!controldata) {
          console.log(control_id, control_list)
        }
        if (controldata) {
          if (controldata.score_options && controldata.score_options.run_id && controldata.score_options.paragraph_id) {
            var paragraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(controldata.score_options.paragraph_id))
            if (paragraph) {
              for (var i = 0; i < paragraph.Paragraph.Content.length; ++i) {
                if (paragraph.Paragraph.Content[i].Id == controldata.score_options.run_id) {
                  return {
                    control_id: control_id,
                    add: false
                  }
                }
              }
            }
          }
          var scores = params.scores
          var cellcount = scores.length
          var cell_width_mm = 8
          var cell_height_mm = 8
          var cellwidth = cell_width_mm / (25.4 / 72 / 20)
          var cellHeight = cell_height_mm / (25.4 / 72 / 20)
          var maxTableWidth
          if (params.layout == 1) { // 顶部
            maxTableWidth = trips_width
          } else if (params.layout == 2) { // 嵌入式
            maxTableWidth = trips_width / 2
          }
          var rowcount = 1
          var columncount = cellcount
          if (maxTableWidth < cellcount * cellwidth) { // 需要换行
            rowcount = Math.ceil(cellcount * cellwidth / maxTableWidth)
            columncount = Math.ceil(maxTableWidth / cellwidth) 
          }
          var oDocument = Api.GetDocument();
          var oTable = Api.CreateTable(columncount, rowcount);
          oTable.SetWidth('twips', columncount * cellwidth)
          var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
          var oTableStylePr = oTableStyle.GetConditionalTableStyle("wholeTable");
          oTable.SetTableLook(true, true, true, true, true, true);
          oTableStylePr.GetTableRowPr().SetHeight("atLeast", cellHeight); // 高度至少多少trips
          var oTableCellPr = oTableStyle.GetTableCellPr();
          oTableCellPr.SetVerticalAlign("center");
          oTable.SetWrappingStyle(params.layout == 1 ? true : false);
          oTable.SetStyle(oTableStyle);
          var oMyStyle = oDocument.CreateStyle("ques_score_cell")
          var oParaPr = oMyStyle.GetParaPr()
          var oTextPr = oMyStyle.GetTextPr()
          oTextPr.SetColor(0, 0, 0, false)
          oTextPr.SetFontSize(16)
          oParaPr.SetJc("center") // 设置段落内容对齐方式。
          var mergecount = rowcount * columncount - cellcount
          if (mergecount > 0) {
            var cells = []
            for (var k = 0; k < mergecount; ++k) {
              cells.push(oTable.GetRow(rowcount - 1).GetCell(k))
            }
            oTable.MergeCells(cells);
          }
          var scoreindex = -1
          for (var irow = 0; irow < rowcount; ++irow) {
            var cbegin = 0
            var cend = columncount
            if (mergecount > 0 && irow == rowcount - 1) { // 最后一行
              cbegin = 1
              cend = columncount - mergecount + 1
            }
            for (var icolumn = cbegin; icolumn < cend; ++icolumn) {
              var cr = irow
              var cc = icolumn
              scoreindex++
              if (scoreindex >= scores.length) {
                break
              }
              var cell = oTable.GetCell(cr, cc)
              if (cell) {
                var cellcontent = cell.GetContent()
                if (cellcontent) {
                  var oCellPara = cellcontent.GetElement(0)
                  if (oCellPara) {
                    oCellPara.AddText(scores[scoreindex].v)
                    // cell.SetWidth('percent', 100/(num+ 2)) // 设置每个单元格的宽度占比
                    cell.SetWidth("twips", cellwidth);
                    oCellPara.SetStyle(oMyStyle)
                    scores[scoreindex].row = cr
                    scores[scoreindex].column = cc
                  } else {
                    console.log('oCellPra is null')
                  }
                } else {
                  console.log('cellcontent is null')
                }
              } else {
                console.log('cannot get cell', cc, cr)
              }
            }
          }
          // 表格-高级设置 相关参数
          var Props = {
            CellSelect: true,
            Locked: false,
            PositionV: {
              Align: 1,
              RelativeFrom: 2,
              UseAlign:true,
              Value: 0
            },
            PositionH: {
              Align: 4,
              RelativeFrom: 0,
              UseAlign:true,
              Value: 0
            },
            TableDefaultMargins: {
              Bottom: 0,
              Left: 0,
              Right: 0,
              Top: 0
            }
          }
          oTable.Table.Set_Props(Props);
          if (oTable.SetLockValue) {
            oTable.SetLockValue(true)
          } else {
            console.log('not function SetLockValue')
          }
          console.log('add table', oTable)
          var oFill = Api.CreateNoFill()
          var oStroke = Api.CreateStroke(3600, Api.CreateNoFill());
          var oDrawing = Api.CreateShape("rect",  (cell_width_mm * columncount + columncount / 2) * 36E3, (cell_height_mm * rowcount + 4) * 36E3, oFill, oStroke);
          
          oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
          var drawDocument = oDrawing.GetContent()
          drawDocument.AddElement(0, oTable)

          var paragraphs = control.GetContent().GetAllParagraphs()

          var paragraph = paragraphs[0]
          if (params.layout == 2) { // 嵌入式
            oDrawing.SetWrappingStyle("square")
            oDrawing.SetHorPosition("column", (width - ((cell_width_mm + 0.3) * columncount)) * 36E3);
          } else { // 顶部
            oDrawing.SetWrappingStyle("topAndBottom")
            oDrawing.SetHorPosition("column", (width - ((cell_width_mm + 0.3) * columncount)) * 36E3);
            oDrawing.SetVerAlign("paragraph")
          }

          var oRun = Api.CreateRun();          
          oRun.AddDrawing(oDrawing);
          var r = paragraph.AddElement(oRun, 1);
          return {
            control_id: control.Sdt.GetId(),
            add: true,
            score: score,
            score_options: {
              paragraph_id: paragraph.Paragraph.Id,
              table_id: oTable.Table.Id,
              run_id: oRun.Run.Id,
              drawing_id: oDrawing.Drawing.Id,
              mode: params.mode,
              layout: params.layout,
              table_cells: scores
            }
          }
        }
      }
    }
    return {
      add: false
    }
  }, false, true, function(res) {
    console.log(res)
    if (res && res.add) {
      for (var i = 0; i < window.BiyueCustomData.control_list.length; ++i) {
        if (window.BiyueCustomData.control_list[i].control_id == res.control_id) {
          window.BiyueCustomData.control_list[i].score = score
          window.BiyueCustomData.control_list[i].score_options = res.score_options
          break  
        }
      }
    }
  })
}

function selectQues(treeInfo, index) {
  Asc.scope.temp_sel_index = index
  window.Asc.plugin.callCommand(function () {
    var res = Api.GetDocument().GetAllContentControls()
    var index = Asc.scope.temp_sel_index
    if (res && res[index]) {
      res[index].GetRange().Select()
    }
  }, false, false, undefined)
}

function drawPosition2(data) {
  console.log('drawPosition', data)
  // 绘制区域，需要判断原本是否有这个区域，如果有，修改位置，如果没有，添加
  Asc.scope.pos = data
  Asc.scope.MM2EMU = MM2EMU
  console.log('window.BiyueCustomData.control_list', window.BiyueCustomData.control_list)
  var find = window.BiyueCustomData.control_list.find(e => {
    return e.regionType == 'feature' && e.zone_type == data.zone_type && e.v == data.v
  })
  if (find) { // 原本已经有
    Asc.scope.control_id = find.control_id
  } else {
    Asc.scope.control_id = null
  }
  // testAddTable()
  window.Asc.plugin.callCommand(function () {
    var posdata = Asc.scope.pos
    var MM2EMU = Asc.scope.MM2EMU
    var control_id = Asc.scope.control_id
    var oDocument = Api.GetDocument();
    if (control_id) {
      var tag = JSON.stringify({
        regionType: 'feature',
        zone_type: posdata.zone_type,
        mode: 6,
        v: posdata.v
      })
      var oControls = oDocument.GetAllContentControls();
        for (var i = 0; i < oControls.length; i++) {
            var oControl = oControls[i];
            if (oControl.Sdt.GetId() === control_id) {
              console.log('找到control 了', oControl)
              break;
            }
        }
      var controls = oDocument.GetContentControlsByTag(tag)
      return {
        add: false
      }
    } else {
      var oFill = Api.CreateNoFill()
      var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
      var oStroke = Api.CreateStroke(3600, oFill2);
      // 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
      var oDrawing = Api.CreateShape("rect", MM2EMU * posdata.w, MM2EMU * posdata.h, oFill, oStroke);
      var drawDocument = oDrawing.GetContent()
      var oParagraph = Api.CreateParagraph();
      var oMyStyle = oDocument.CreateStyle("field style")
      var oTextPr = oMyStyle.GetTextPr()
      oTextPr.SetColor(0, 0, 0, false)
      oTextPr.SetFontSize(24)
      var oParaPr = oMyStyle.GetParaPr()
      oParaPr.SetJc("center")
      oParagraph.SetStyle(oMyStyle)
      var text = ''
      if (posdata.zone_type == 15) {
        text = '再练'
      } else if (posdata.zone_type == 16) {
        text = '完成'
      }
      oParagraph.AddText(text)
      drawDocument.AddElement(0,oParagraph)
      oDrawing.SetVerticalTextAlign("center")
      oDrawing.SetVerPosition('page', 914400)
      oDrawing.SetSize(914400, 914400)

      oDocument.AddDrawingToPage(oDrawing, 0, MM2EMU * posdata.x, MM2EMU * posdata.y);
      var range = drawDocument.GetRange()
      range.Select()
      var tag = {
        regionType: 'feature',
        zone_type: posdata.zone_type,
        mode: 6,
        v: posdata.v
      }
      var oResult = Api.asc_AddContentControl(range.controlType || 1, {"Tag": JSON.stringify(tag)});
      Api.asc_RemoveSelection();
      return {
        add: true,
        control: {
          control_id: oResult.InternalId,
          regionType: 'feature',
          v: posdata.v,
          zone_type: posdata.zone_type
        }
      }
    }
    
  }, false, true, function(res) {
    if (res && res.add) {
      window.BiyueCustomData.control_list.push(res.control)
    }
  })
}

function drawPosition(data) {
  // 绘制区域，需要判断原本是否有这个区域，如果有，修改位置，如果没有，添加
  Asc.scope.pos = data
  Asc.scope.MM2EMU = MM2EMU
  var find = null
  if (window.BiyueCustomData.pos_list) {
    find = window.BiyueCustomData.pos_list.find(e => {
      return e.zone_type == data.zone_type && e.v == data.v
    })
  }
  if (find) { // 原本已经有
    Asc.scope.drawing_id = find.id
  } else {
    Asc.scope.drawing_id = null
  }
  // testAddTable()
  window.Asc.plugin.callCommand(function () {
    var posdata = Asc.scope.pos
    var MM2EMU = Asc.scope.MM2EMU
    var drawing_id = Asc.scope.drawing_id
    var oDocument = Api.GetDocument();
    if (drawing_id) {
      var objs = oDocument.GetAllDrawingObjects()
      for (var i = 0; i < objs.length; i++) {
          var oDrawing = objs[i];
          if (oDrawing.Drawing.Id == drawing_id) {
            oDrawing.SetSize(MM2EMU * posdata.w, MM2EMU * posdata.h)
            oDrawing.SetVerPosition('page', MM2EMU * posdata.x)
            oDrawing.SetHorPosition('page', MM2EMU * posdata.y)
            break
          }
      }
      return {
        add: false
      }
    } else {
      var oFill = Api.CreateNoFill()
      var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
      var oStroke = Api.CreateStroke(3600, oFill2);
      // 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
      var oDrawing = Api.CreateShape("rect", MM2EMU * posdata.w, MM2EMU * posdata.h, oFill, oStroke);
      
      var drawDocument = oDrawing.GetContent()
      var oParagraph = Api.CreateParagraph();
      var oMyStyle = oDocument.CreateStyle("field style")
      var oTextPr = oMyStyle.GetTextPr()
      oTextPr.SetColor(255, 0, 0, false)
      oTextPr.SetFontSize(24)
      var oParaPr = oMyStyle.GetParaPr()
      oParaPr.SetJc("center")
      oParagraph.SetStyle(oMyStyle)
      var text = ''
      if (posdata.zone_type == 15) {
        text = '再练'
      } else if (posdata.zone_type == 16) {
        text = '完成'
      }
      oParagraph.AddText(text)
      drawDocument.AddElement(0,oParagraph)
      oDrawing.SetVerticalTextAlign("center")

      // debugger
      oDrawing.SetVerPosition('page', 914400)
      oDrawing.SetSize(MM2EMU * posdata.w, MM2EMU * posdata.h)
      oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
      oDocument.AddDrawingToPage(oDrawing, 0, MM2EMU * posdata.x, MM2EMU * posdata.y);
      return {
        add: true,
        id: oDrawing.Drawing.Id,
        zone_type: posdata.zone_type,
        v: posdata.v
      }
    }
    
  }, false, true, function(res) {
    console.log(res)
    if (res && res.add) {
      if (!window.BiyueCustomData.pos_list) {
        window.BiyueCustomData.pos_list = []
      }
      window.BiyueCustomData.pos_list.push({
        id: res.id,
        zone_type: res.zone_type,
        v: res.v
      })
    }
  })
}

function delScoreField() {
  Asc.scope.control_list = window.BiyueCustomData.control_list
  window.Asc.plugin.callCommand(function() {
    var control_list = Asc.scope.control_list
    var oDocument = Api.GetDocument() 
    var controls = oDocument.GetAllContentControls();
    for(var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var strtag = control.GetTag()
      if (!strtag) {
        continue
      }
      var tag = JSON.parse(strtag)
      if (tag.regionType == 'question') {
        var controlid = control.Sdt.GetId()
        var controldata = control_list.find(e => {
          return e.control_id == controlid
        })
        if (controldata && controldata.score_options && controldata.score_options.paragraph_id) {
          var paragraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(controldata.score_options.paragraph_id))
          if (paragraph) {
            for (var i = 0; i < paragraph.Paragraph.Content.length; ++i) {
              if (paragraph.Paragraph.Content[i].Id == controldata.score_options.run_id) {
                paragraph.RemoveElement(i)
                break
              }
            }
            return {
              code: true,
              control_id: controlid
            }
          }
        }
        break
      }
    }
  }, false, true, function(res) {
    if (res && res.code) {
      var list = window.BiyueCustomData.control_list
      if (list) {
        for (var i = 0; i < list.length; ++i) {
          if (list[i].control_id == res.control_id) {
            list[i].score_options = null
            break
          }
        }
      }
    }
  })
}

function getScores(score, mode) {
  var scores = []
  if (mode == 1) { // 普通模式
    for (var i = 0; i <= score; ++i) {
      scores.push({
        v: i + ''
      })
    }
  } else if (mode == 2) { // 大分值模式
    var ten = (score - score % 10) / 10
    for (var i = 0; i <= ten; ++i ) {
      scores.push({
        v: i == 0 ?`${i}` : `${i}0+`
      })
    }
    if (ten >= 1) {
      for (var j = 1; j < 10; ++j) {
        scores.push({
          v: j + ''
        })
      }
    }
  }
  scores.push({
    v: '0.5'
  })
  return scores
}
// 修改分数显示，数值，模式，布局
function changeScoreField(cscore, cmode, clayout) {
  var control_list = window.BiyueCustomData.control_list
  var find = control_list.find(e => {
    return e.regionType == 'question'
  })
  if (!find) {
    return
  }
  if (!find.score_options) {
    return
  }
  if (!find.score_options.paragraph_id) {
    return
  }
  Asc.scope.params = Object.assign({}, find, {
    target_score: cscore == null ? find.score : (find.score == 25 ? 15 : 25),
    target_mode: cmode == null ? find.score_options.mode : (find.score_options.mode == 1 ? 2 : 1),
    target_layout: clayout == null ? find.score_options.layout : (find.score_options.layout == 1 ? 2 : 1),
  })
  Asc.scope.params.target_scores = getScores(Asc.scope.params.target_score, Asc.scope.params.target_mode)
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument() 
    var controls = oDocument.GetAllContentControls();
    var params = Asc.scope.params
    console.log('params', params)
    for(var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var strtag = control.GetTag()
      if (!strtag) {
        continue
      }
      var tag = JSON.parse(strtag)
      if (tag.regionType == 'question') {
        var oRange = control.GetRange()
        var text =  oRange.GetText()
        var controlid = control.Sdt.GetId()
        if (controlid == params.control_id) {
          var run = new Api.private_CreateApiRun(AscCommon.g_oTableId.Get_ById(params.score_options.run_id))
          var paragraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(params.score_options.paragraph_id))
          var drawingObjs = control.GetAllDrawingObjects()
          if (params.target_layout != params.layout) { // 修改布局，需要调整shape的属性，若表格行列数有变，要先调整表格
            var oDrawing = drawingObjs.find(d => {
              return params.score_options.drawing_id == d.Drawing.Id
            })
            if (oDrawing) {
              var rect = Api.asc_GetContentControlBoundingRect(params.control_id, true);
              var tables = control.GetContent().GetAllTables()
              var oTable
              if (tables) {
                oTable = tables.find(e => {
                  return e.Table.Id == params.score_options.table_id
                })
              }
              if (oTable) {
                var width = rect.X1 - rect.X0
                var trips_width = width / (25.4 / 72 / 20)
                var cell_width_mm = 8
                var cell_height_mm = 8
                var cellwidth = cell_width_mm / (25.4 / 72 / 20)
                var cellHeight = cell_height_mm / (25.4 / 72 / 20)
                var maxTableWidth = params.target_layout == 1 ? trips_width : (trips_width / 2)
                var rowcount = 1
                var cellcount = params.target_scores.length
                var columncount = cellcount
                if (maxTableWidth < cellcount * cellwidth) { // 需要换行
                  rowcount = Math.ceil(cellcount * cellwidth / maxTableWidth)
                  columncount = Math.ceil(maxTableWidth / cellwidth) 
                }
                var oldrowcount = oTable.GetRowsCount()
                if (oldrowcount != rowcount) { // 行数不同
                  if (oldrowcount > rowcount) {
                    for (var b = oldrowcount; b > rowcount; --b) {
                      oTable.RemoveRow(oTable.GetRow(oldrowcount - 1).GetCell(0));
                    }
                    var row = oTable.GetRow(rowcount - 1)
                    var oldcolumncount = row.GetCellsCount()
                    oTable.AddColumns(row.GetCell(oldcolumncount - 1), columncount - oldcolumncount, false);
                    for (var cindex = oldcolumncount; cindex < columncount; ++cindex) {
                      var cell = oTable.GetCell(rowcount - 1, cindex)
                      if (cell) {
                        var cellcontent = cell.GetContent()
                        if (cellcontent) {
                          var oCellPara = cellcontent.GetElement(0)
                          if (oCellPara) {
                            var scoreindex = params.target_scores.length - (columncount - cindex)
                            oCellPara.AddText(params.target_scores[scoreindex].v)
                            cell.SetWidth("twips", cellwidth);
                            var cellstyle = oDocument.GetStyle('ques_score_cell')
                            oCellPara.SetStyle(cellstyle)
                            params.score_options.table_cells[scoreindex].row = rowcount - 1
                            params.score_options.table_cells[scoreindex].column = cindex
                          } else {
                            console.log('oCellPra is null')
                          }
                        } else {
                          console.log('cellcontent is null')
                        }
                      }
                    }
                    
                  }
                  oDrawing.SetSize((cell_width_mm * columncount + columncount / 2) * 36E3, (cell_height_mm * rowcount + 4) * 36E3)
                }
                if (params.target_layout == 2) { // 嵌入式
                  oDrawing.SetWrappingStyle("square")
                  oDrawing.SetHorPosition("column", (width - ((cell_width_mm + 0.3) * columncount)) * 36E3);
                } else { // 顶部
                  oDrawing.SetWrappingStyle("topAndBottom")
                  oDrawing.SetHorPosition("column", (width - ((cell_width_mm + 0.3) * columncount)) * 36E3);
                  oDrawing.SetVerAlign("paragraph")
                }
                params.score_options.layout = params.target_layout
              }
            }
          }
          return {
            code: 1,
            control_id: controlid,
            score: params.target_score,
            score_options: params.score_options
          }
        }
          
      }
    }
  }, false, true, function(res) {
    if (res && res.code) {
      var list = window.BiyueCustomData.control_list
      if (list) {
        for (var i = 0; i < list.length; ++i) {
          if (list[i].control_id == res.control_id) {
            list[i].score = res.score
            list[i].score_options = res.score_options
            break
          }
        }
      }
    }
  })
}

function addImage() {
  toggleWeight()
  return
  window.Asc.plugin.callCommand(function () {
    var oDocument = Api.GetDocument();
    let controls = oDocument.GetAllContentControls();
    for (var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'sub-question') {
        var imgurl = 'https://teacher.biyue.tech/teacher/static/img_20240430095417/logo.png'
        var oDrawing = Api.CreateImage(imgurl, 6 * 36000, 5 * 36000);
        var paragraphs = control.GetContent().GetAllParagraphs()
        var paragraph = paragraphs[0]
        oDrawing.SetWrappingStyle("inline")
        oDrawing.SetLockValue("noSelect", true)
        oDrawing.SetLockValue('noResize', true)
        var oRun = Api.CreateRun();          
        oRun.AddDrawing(oDrawing);
        var r = paragraph.AddElement(oRun);
        break
      }
    }
  }, false, true, function(res) {

  })
}
// 添加批改区域
function addMarkField() {
  getPos()
  return
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument()
    let controls = oDocument.GetAllContentControls();
    for (var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'question') {
        var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
        oFill.UniFill.transparent = 20 // 透明度
        var oFill2 = Api.CreateRGBColor(125, 125, 125)
        var oStroke = Api.CreateStroke(3600, oFill2);
        var oDrawing = Api.CreateShape('rect', 20 * 36000, 20 * 36000, oFill, oStroke)
        var paragraphs = control.GetContent().GetAllParagraphs()
        var paragraph = paragraphs[0]
        oDrawing.SetWrappingStyle("inFront")
        var oRun = Api.CreateRun();
        oRun.AddDrawing(oDrawing);
        var r = paragraph.AddElement(oRun);
        var range = oDrawing.GetContent().GetRange()
        range.Select()
        break
      }
    }
  }, false, true, function(res) {

  })
}
// 切换权重显示
function toggleWeight() {
  Asc.scope.control_list = window.BiyueCustomData.control_list
  window.Asc.plugin.callCommand(function() {
    var control_list = Asc.scope.control_list
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    var added = false
    for (var i = 0; i < control_list.length; ++i) {
      var controlData = control_list[i]
      if (controlData.regionType == 'question' && controlData.ask_controls && controlData.ask_controls.length) {
        var controlRect = Api.asc_GetContentControlBoundingRect(controlData.control_id, true)
        for (var iask = 0; iask < controlData.ask_controls.length; ++iask) {
          var askControlId = controlData.ask_controls[iask].control_id
          var askControl = controls.find(e => {
            return e.Sdt.GetId() == askControlId
          })
          if (!controlData.ask_controls[iask].weight_options) { // 还没有权重的id，需要创建
            var rect = Api.asc_GetContentControlBoundingRect(askControlId, true)
            var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
            oFill.UniFill.transparent = 40 // 透明度
            var oStroke = Api.CreateStroke(3600, Api.CreateNoFill());
            var width = rect.X1 - rect.X0
            var height = rect.Y1 - rect.Y0
            var oDrawing = Api.CreateShape("rect", width * 36E3, height * 36E3, oFill, oStroke);
            oDrawing.SetPaddings(0, 0, 0, 0);
            oDrawing.SetWrappingStyle('inFront')
            oDrawing.SetHorPosition("column",(rect.X0 - controlRect.X0) * 36E3)
            oDrawing.Drawing.Set_Props({
              title: 'ask_weight'
            })
            

            var drawDocument = oDrawing.GetContent()
            var oParagraph = Api.CreateParagraph();    
            oParagraph.AddText((iask + 1) + '')
            drawDocument.AddElement(0,oParagraph)
            oParagraph.SetJc("center");
            var oTextPr = Api.CreateTextPr();
            oTextPr.SetColor(255, 111, 61, false)
            oTextPr.SetFontSize(14)
            oParagraph.SetTextPr(oTextPr);
            oDrawing.SetVerticalTextAlign("center")

            var oRun = Api.CreateRun();
            oRun.AddDrawing(oDrawing);
            var asktype = askControl.GetClassType()
            if (asktype == 'inlineLvlSdt') {
              askControl.AddElement(oRun)
              console.log('ask oParagraph', oParagraph, 'run', oRun, 'oDrawing', oDrawing)
              controlData.ask_controls[iask].weight_options = {
                paragraph_id: oParagraph.Paragraph.Id,
                run_id: oRun.Run.Id,
                drawing_id: oDrawing.Drawing.Id,
                show: true
              }
              added = true
            } else if (asktype == 'blockLvlSdt') {
              // todo..
            }
          }
        }
      }
    }
    for (var i = 0; i < control_list.length; ++i) {
      var controlData = control_list[i]
      if (controlData.regionType != 'question') {
        continue
      }
      var pcontrol = controls.find(e => {
        return e.Sdt.GetId() == controlData.control_id
      })
      if (pcontrol) {
        if (pcontrol.GetAllDrawingObjects) {
          var drawingObjs = pcontrol.GetAllDrawingObjects()
          drawingObjs.forEach(e => {
            if (e.Drawing.docPr.title == "ask_weight") {
              if (!added) {
                e.Fill(Api.CreateNoFill())
              } else {
                e.Drawing.Set_Props({
                  title: 'ask_weight',
                  hidden: null
                })
              }
            }
          })
          console.log('drawingObjs', drawingObjs)
        } else {
          console.log('GetAllDrawingObjects is unvalid', pcontrol)
        }
        if (controlData.ask_controls) {
          controlData.ask_controls.forEach(ask => {
            if (ask.weight_options && ask.weight_options.paragraph_id) {
              var paragraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(ask.weight_options.paragraph_id))
              if (!added) {
                paragraph.RemoveElement(0)
              }
            }
          })
        }
      }
    }

    return control_list
  }, false, true, function(res) {
    window.BiyueCustomData.control_list = res
  })
}

function getPos() {
  Asc.scope.control_list = window.BiyueCustomData.control_list
  window.Asc.plugin.callCommand(function () {
    var control_list = Asc.scope.control_list
    console.log('control_list', control_list)
    var oDocument = Api.GetDocument()
    for (var i = 0; i < control_list.length; ++i) {
      if (control_list[i].regionType == 'question') {
        var score_options = control_list[i].score_options
        if (score_options && score_options.table_id) {
          var controls = oDocument.GetAllContentControls()
          var control = controls.find(e => {
            return e.Sdt.GetId() == control_list[i].control_id
          })
          var tables = control.GetContent().GetAllTables()
          var oTable
          if (tables) {
            oTable = tables.find(e => {
              return e.Table.Id == control_list[i].score_options.table_id
            })
          }
          if (oTable) {
            console.log('oTable', oTable)
            var rowCount = oTable.GetRowsCount()
            for (var irow = 0; irow < rowCount; ++irow) {
              var row = oTable.GetRow(irow)
              console.log('row', row)
            }
          }
        }
      }
    }
  }, false, false, function(res) {

  })
}

function handleContentControlChange(params) {
  var controlId = params.InternalId
  var control_list = window.BiyueCustomData.control_list
  var tag = params.Tag
  if (tag) {
    tag = JSON.parse(params.Tag)
    if (tag.regionType == 'question') {
      var find = control_list.find(e => {
        return e.control_id == controlId
      })
      Asc.scope.find_controldata = find   
    } else {
      return
    }
  }
  Asc.scope.params = params
  window.Asc.plugin.callCommand(function () {
    var controldata = Asc.scope.find_controldata
    var params = Asc.scope.params
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    var control = controls.find(e => {
      return e.Sdt.GetId() == params.InternalId
    })
    if (!controldata) {
      if (control && control.GetAllDrawingObjects) {
        var drawingObjs = control.GetAllDrawingObjects()
        for (var i = 0, imax = drawingObjs.length; i < imax; ++i) {
          var oDrawing = drawingObjs[i]
          if (oDrawing.Drawing.docPr.title == 'ask_weight') {
            oDrawing.Delete()
          }
        }
      }
    }
  }, false, true, undefined)

}

export {
  getPaperInfo,
  initPaperInfo,
  updateCustomControls,
  clearStruct,
  getStruct,
  savePositons,
  showQuestionTree,
  updateQuestionScore,
  testAddTable,
  drawPosition,
  addQuesScore,
  addScoreField,
  delScoreField,
  changeScoreField,
  addImage,
  addMarkField,
  handleContentControlChange
}