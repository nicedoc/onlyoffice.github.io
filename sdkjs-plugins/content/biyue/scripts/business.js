// 主要用于处理业务逻辑
import { paperOnlineInfo, structAdd, questionCreate, questionDelete, questionUpdateContent, structDelete, paperCanConfirm, structRename, paperSavePosition, examQuestionsUpdate } from './api/paper.js'
import { JSONPath } from '../vendor/jsonpath-plus/dist/index-browser-esm.js';
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
  })
}

function getPaperInfo() {
  return paper_info
}

// 更新customData的control_list
function updateCustomControls() {
  Asc.scope.paper_info = paper_info
  window.Asc.plugin.callCommand(function() {
    let paperinfo = Asc.scope.paper_info
    var oDocument = Api.GetDocument();
    let controls = oDocument.GetAllContentControls();
    let ques_no = 1
    let struct_index = 0
    var control_list = []
    controls.forEach(control => {
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
        console.log('control', control)
        obj.zone_type = tagInfo.zone_type
        obj.v = tagInfo.v
      }
      control_list.push(obj)
    })
    return control_list
  }, false, false, function(control_list) {
    window.BiyueCustomData.control_list = control_list
    console.log('control_list', control_list)
  })
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

// 添加分数框
function addScoreField() {
  Asc.scope.control_list = window.BiyueCustomData.control_list
  console.log('control_list', window.BiyueCustomData.control_list)
  window.Asc.plugin.callCommand(function() {
    var control_list = Asc.scope.control_list
    var controls = Api.GetDocument().GetAllContentControls();
    controls.forEach(control => {
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'question') {
        var controldata = control_list.find(item => {
          return item.regionType == 'question' && item.control_id == control.Sdt.GetId()
        })
        if (!controldata) {
          console.log(control.Sdt.GetId(), control_list)
        }
        if (controldata && (controldata.ques_type * 1 == 3) && controldata.score > 0) {
          var oDocument = Api.GetDocument();
          var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
          let num = controldata.score * 1 + 1
          var oTable = Api.CreateTable(num + 1, 1);
          // oTable.SetWidth("percent", 100);
          oTable.SetStyle(oTableStyle);
          oTable.SetWrappingStyle(false);
          for (var i = 0; i <= num; i++) {
            oTable.GetCell(0, i).GetContent().GetElement(0).AddText(i== num ? '0.5' : (i + ''))
            oTable.GetCell(0, i).SetWidth('twips', 283)
          }
          
          // oDocument.Push(oTable) // 添加到文档的底部

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
          // Api.tblApply(Props)
          oTable.Table.Set_Props(Props);
          control.AddElement(oTable, 0)
        }
      }
      
    })
  }, false, true, function() {

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

function drawPosition(data) {
  console.log('drawPosition', data)
  // 绘制区域，需要判断原本是否有这个区域，如果有，修改位置，如果没有，添加
  Asc.scope.pos = data
  Asc.scope.MM2EMU = MM2EMU
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
    console.log('posdata', posdata, MM2EMU)
    var oDocument = Api.GetDocument();
    if (control_id) {
      var tag = JSON.stringify({
        regionType: 'feature',
        zone_type: posdata.zone_type,
        mode: 6,
        v: posdata.v
      })
      var oControls = oDocument.GetAllContentControls();
      console.log('control_id', control_id)
        for (var i = 0; i < oControls.length; i++) {
            var oControl = oControls[i];
            if (oControl.Sdt.GetId() === control_id) {
              var oDrawing = oControl.Sdt.LogicDocument.DrawingObjects
              console.log('oDrawing', oDrawing)
              console.log('oControl', oDrawing.drawingObjects)
              var paradrawing = oDrawing.drawingObjects[0]
              console.log('paradrawing', paradrawing)
              paradrawing.setExtent(20, 10)
              console.log('Asc.c_oAscRelativeFromH', Asc.c_oAscRelativeFromH)
              // paradrawing.Set_PositionH(nRelativeFrom, false, 50, false)
              // paradrawing.Set_PositionV(nRelativeFrom, false, 80, false)
              // console.log('CShape', paradrawing.GraphicObj)
              // console.log('setExtX', paradrawing.GraphicObj.spPr.xfrm.setExtX)
              // paradrawing.GraphicObj.spPr.xfrm.setExtX(50)
              // paradrawing.GraphicObj.spPr.xfrm.setExtY(80)
              // this.Drawing.GraphicObj.spPr.xfrm.setExtX(fWidth);
                // this.Drawing.GraphicObj.spPr.xfrm.setExtY(fHeight)
              break;
            }
        }
      var controls = oDocument.GetContentControlsByTag(tag)
      console.log('control', controls)
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
      debugger
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
    console.log(res)
    if (res && res.add) {
      window.BiyueCustomData.control_list.push(res)
    }
  })
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
  drawPosition
}