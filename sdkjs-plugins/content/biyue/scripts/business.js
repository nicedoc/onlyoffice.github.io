// 主要用于处理业务逻辑
import { paperOnlineInfo, structAdd, questionCreate, questionDelete, questionUpdateContent, structDelete, paperCanConfirm, structRename, paperSavePosition, examQuestionsUpdate } from './api/paper.js'
import { getBase64, map_base64 } from '../resources/list_base64.js'
let paper_info = {} // 从后端返回的试卷信息
let select_ques_ids = []
const MM2EMU = 36000 // 1mm = 36000EMU
// 根据paper_uuid获取试卷信息
function initPaperInfo() {
  return new Promise((resolve, reject) => {
    paperOnlineInfo(window.BiyueCustomData.paper_uuid).then(res => {
      paper_info = res.data
      window.BiyueCustomData.exam_title = paper_info.paper.title
      console.log('试卷信息')
      console.log(paper_info)
      updateControls().then(res => {
        resolve('iniPaperInfo ok updateControls ok')
      })
    }).catch(res => {
      updateControls().then(res => {
        resolve('iniPaperInfo ok getinfo fail updateControls ok')
      })
    })
  })
}

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

function updateQeusTree() {
  var listElement = document.getElementById('ques-list')
  if (!listElement) {
    return
  }
  var control_list = window.BiyueCustomData.control_list || []
  var html = ''
  control_list.forEach(e => {
    if (e.regionType == 'struct') {
      html += `<div class="quesitem" id=${e.control_id} draggable="true"><div class="tag-struct"></div><span class="text-struct">${e.name}</span></div>`
    } else if (e.regionType == 'question') {
      html += `<div class="quesitem" id=${e.control_id} draggable="true"><div class="tag-ques"></div><span class="text-ques">${e.text}</span></div></div>`
    } else if (e.regionType == 'sub-question') {
      html += `<div class="quesitem" id=${e.control_id} draggable="true"><div class="tag-sub-ques"></div><span class="text-ques-sub">${e.text}</span></div></div>`
    }
  })
  listElement.innerHTML = html
  $('.quesitem').on('click', onQuesTreeClick)
  $('.quesitem').on('dragstart', onTreeDragStart)
  $('.quesitem').on('dragover', onDragOver)
  $('.quesitem').on('drop', function(e) {
    console.log('drop', e)
    e.preventDefault();
  })
  listElement.addEventListener('drop', (e) => {
    console.log('list drop', e)
    e.preventDefault();
  })
}

function onTreeDragStart(e) {
  console.log('onTreeDragStart', e)
  e.originalEvent.dataTransfer.setData('drag_data', e.target.id);
}

function onDragOver(e) {
  e.preventDefault()
  e.cancelBubble = true
  if (e.dataTransfer) {
    e.dataTransfer.dropEffect = "move"
  } else {
    e.originalEvent.dataTransfer.dropEffect = "move"
  }
}

function onQuesTreeClick(e) {
  console.log('onQuesTreeClick', e)
  var id
  if (e.target && e.target.id && e.target.id != '') {
    id = e.target.id
  } else if (e.currentTarget) {
    id = e.currentTarget.id
  }
  var control_list = window.BiyueCustomData.control_list || []
  var controlData = control_list.find(item => {
    return item.control_id == id
  })
  if (!controlData) {
    console.log('onQuesTreeClick cannot find ', id)
    return
  }
  var ctrlKey = e.ctrlKey
  var newlist = []
  if (ctrlKey) {
    newlist = [].concat(select_ques_ids)
    var index = newlist.indexOf(id)
    if (index >=0) { // 原本已存在，取消选中
      newlist.splice(index, 1)
    } else {
      newlist.push(id)
    }
  } else {
    newlist.push(id)
    var event = new CustomEvent('clickSingleQues', {detail: {
      control_id: id,
      regionType: 'question'
    }})
    document.dispatchEvent(event)
  }
  updateQuesStyle(newlist)
  Asc.scope.click_ids = newlist
  window.Asc.plugin.callCommand(function() {
    var ids = Asc.scope.click_ids
    var oDocument = Api.GetDocument()
    oDocument.RemoveSelection()
    var controls = oDocument.GetAllContentControls()
    var firstRange = null
    ids.forEach((id, index) => {
      var control = controls.find(e => {
        return e.Sdt.GetId() == id
      })
      if (control) {
        if (index == 0) {
          firstRange = control.GetRange()
        } else {
          var oRange = control.GetRange()
          firstRange = firstRange.ExpandTo(oRange)
        }
      }
    })
    firstRange.Select()
  }, false, false, undefined)
}
// 更新题目选中样式
function updateQuesStyle(idList) {
  console.log('updateQuesStyle', idList)
  for (var i = 0; i < select_ques_ids.length; ++i) {
    if (idList.indexOf(select_ques_ids[i]) == -1) {
      $('#'+select_ques_ids[i]).removeClass('selected')
    }
  }
  for (var j = 0; j < idList.length; ++j) {
    $('#'+idList[j]).addClass('selected')
  }
  select_ques_ids = idList
}

function updateControls() {
  return new Promise((resolve, reject) => {
    Asc.scope.paper_info = paper_info
    setupPostTask(window, function(res) {
      console.log('control_list', res)
      window.BiyueCustomData.control_list = res
      updateQeusTree()
      console.log('+++++updateControls callback ')
      resolve()
    })
    window.Asc.plugin.callCommand(function() {
      console.log('+++++++++++++++++++++++')
      let paperinfo = Asc.scope.paper_info
      var oDocument = Api.GetDocument();
      let controls = oDocument.GetAllContentControls() || [];
      let ques_no = 1
      let struct_index = 0
      var control_list = []
      controls.forEach(control => {
        var rect = Api.asc_GetContentControlBoundingRect(control.Sdt.GetId(), true);
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
          const regex = /^([^.．、]*)/;
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
          if (!parentContentControl) {
            console.log('parentContentControl is null')
          }
          if (parentContentControl) {
            obj.parent_control_id = parentContentControl.Sdt.GetId()
          }
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
      console.log(' updatecontrol           control_list', control_list)
      return control_list
    }, false, false, undefined)
  })
}

// 更新customData的control_list
function updateCustomControls() {
  Asc.scope.paper_info = paper_info
  console.log('           updateCustomControls')
  setupPostTask(window, function(res) {
    console.log('control_list', res)
    window.BiyueCustomData.control_list = res
    updateQeusTree()
  })
  console.log('           updateCustomControls============')
  window.Asc.plugin.callCommand(function() {
    console.log('+++++++++++++++++++++++')
    let paperinfo = Asc.scope.paper_info
    var oDocument = Api.GetDocument();
    let controls = oDocument.GetAllContentControls() || [];
    let ques_no = 1
    let struct_index = 0
    var control_list = []
    console.log('===============1')
    controls.forEach(control => {
      console.log('===============2', control)
      var rect = Api.asc_GetContentControlBoundingRect(control.Sdt.GetId(), true);
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
        const regex = /^([^.．、]*)/;
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
        if (!parentContentControl) {
          console.log('parentContentControl is null')
        }
        if (parentContentControl) {
          obj.parent_control_id = parentContentControl.Sdt.GetId()
        }
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
        console.log('===============3')
      } else if (tagInfo.regionType == 'feature') {
        obj.zone_type = tagInfo.zone_type
        obj.v = tagInfo.v
      }
      control_list.push(obj)
    })
    console.log(' updatecontrol           control_list', control_list)
    return control_list
  }, false, false, undefined)
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
  initPaperInfo()
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
  initPaperInfo()
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
function addScoreField(score, mode, layout, posall) {
  Asc.scope.control_list = window.BiyueCustomData.control_list
  Asc.scope.params = {
    score: score,
    mode: mode,
    layout: layout,
    scores: getScores(score, mode),
    posall: posall
  }
  setupPostTask(window, function(r) {
    console.log('callcack addScoreField', r)
  })
  console.log('setup post task for addScoreField')
  window.Asc.plugin.callCommand(function() {
    var control_list = Asc.scope.control_list
    var controls = Api.GetDocument().GetAllContentControls();
    var params = Asc.scope.params
    var score = params.score
    var res = {
      add: false
    }
    for (var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'question') {
        var control_id = control.Sdt.GetId()
        var controlIndex = control_list.findIndex(item => {
          return item.regionType == 'question' && item.control_id == control_id
        })
        var controldata = control_list.find(item => {
          return item.regionType == 'question' && item.control_id == control_id
        })
        var rect = Api.asc_GetContentControlBoundingRect(control_id, true);
        console.log('rect', rect)
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
          var maxTableWidth = trips_width / params.layout
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
          var mergecount = rowcount * columncount - cellcount
          if (mergecount > 0) {
            var cells = []
            for (var k = 0; k < mergecount; ++k) {
              cells.push(oTable.GetRow(rowcount - 1).GetCell(k))
            }
            oTable.MergeCells(cells);
          }
          var scoreindex = -1
          // 设置单元格文本
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
                    cell.SetWidth("twips", cellwidth);
                    oCellPara.SetJc('center')
                    oCellPara.SetColor(0, 0, 0, false)
                    oCellPara.SetFontSize(16)
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
          var wline = params.layout == 2 ? 0.3 : 0.25
          var shapew = (cell_width_mm + wline) * columncount
          var shapeh = (cell_height_mm * rowcount + 4)
          var oDrawing = Api.CreateShape("rect",  shapew * 36E3, shapeh * 36E3, oFill, oStroke);
          
          // oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
          var drawDocument = oDrawing.GetContent()
          drawDocument.AddElement(0, oTable)
          if (params.layout == 2) { // 嵌入式
            oDrawing.SetWrappingStyle("square")
            oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
          } else { // 顶部
            oDrawing.SetWrappingStyle("topAndBottom")
            oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
            oDrawing.SetVerAlign("paragraph")
          }
          var titleobj = {
            type: 'qscore',
            ques_control_id: control_id
          }
          oDrawing.Drawing.Set_Props({
            title: JSON.stringify(titleobj)
          })
          var paragraph
          if (params.posall) {
            // paragraph = Api.CreateParagraph();
            // paragraph.AddElement(oRun, 1);
            control.GetRange().Select()
            oDocument.AddDrawingToPage(oDrawing, 0, (rect.X1 - shapew) * 36E3, rect.Y0 * 36E3); // 用这种方式加入的一定是相对页面的
            oDrawing.SetVerAlign("paragraph", rect.Y0 * 36E3)
            Api.asc_RemoveSelection();
          } else {
            console.log('++++++++++++++++')
            var oRun = Api.CreateRun();          
            oRun.AddDrawing(oDrawing);
            var paragraphs = control.GetContent().GetAllParagraphs()
            paragraph = paragraphs[0]
            paragraph.AddElement(oRun, 1);
          }
          res = {
            control_id: control.Sdt.GetId(),
            add: true,
            score: score,
            score_options: {
              paragraph_id: paragraph ? paragraph.Paragraph.Id : 0,
              table_id: oTable.Table.Id,
              run_id: oRun ?  oRun.Run.Id : 0,
              drawing_id: oDrawing.Drawing.Id,
              mode: params.mode,
              layout: params.layout,
              table_cells: scores,
              pos_all: params.posall
            }
          }
          console.log('======== ', res.score_options)
          break
        }
      }
    }
    console.log('777777777777 ', res)
    // debugger
    return res
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
      var text = ''
      if (posdata.zone_type == 15) {
        text = '再练'
      } else if (posdata.zone_type == 16) {
        text = '完成'
      }
      oParagraph.AddText(text)
      oParagraph.SetColor(125, 125, 125, false)
      oParagraph.SetFontSize(24)
      oParagraph.SetJc('center')
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

function drawPositions(list) {
  Asc.scope.positions_list = list
  Asc.scope.pos_list = window.BiyueCustomData.pos_list
  setupPostTask(window, function(res) {
    console.log('drawPositions result:', res)
    window.BiyueCustomData.pos_list = res
  })
  Asc.scope.MM2EMU = MM2EMU
  Asc.scope.map_base64 = map_base64
  window.Asc.plugin.callCommand(function() {
    var positions_list = Asc.scope.positions_list || []
    var pos_list = Asc.scope.pos_list || []
    var MM2EMU = Asc.scope.MM2EMU
    var oDocument = Api.GetDocument()
    var objs = oDocument.GetAllDrawingObjects()
    var map_base64 = Asc.scope.map_base64
    positions_list.forEach(e => {
      var posdata = pos_list.find(pos => {
        return pos.zone_type == e.zone_type && pos.v == e.v
      })
      var oDrawing = null
      if (posdata && posdata.drawing_id) { // 已存在
        var index = objs.findIndex(obj => {
          return obj.Drawing.Id == posdata.drawing_id
        })
        if (index >= 0) {
          oDrawing = objs[index]
        }
      }
      if (oDrawing) {
        console.log('已存在')
        oDrawing.SetSize(MM2EMU * e.w, MM2EMU * e.h)
        oDrawing.SetVerPosition('page', MM2EMU * e.x)
        oDrawing.SetHorPosition('page', MM2EMU * e.y)
      } else {
        console.log('不存在')
        if (e.draw_type == 'shape') {
          var oFill = Api.CreateNoFill()
          var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
          var oStroke = Api.CreateStroke(3600, oFill2);
          // 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
          oDrawing = Api.CreateShape("rect", MM2EMU * e.w, MM2EMU * e.h, oFill, oStroke);
          
          var drawDocument = oDrawing.GetContent()
          var oParagraph = Api.CreateParagraph();
          var text = ''
          if (e.zone_type == 15) {
            text = '再练'
          } else if (e.zone_type == 16) {
            text = '完成'
          }
          oParagraph.AddText(text)
          oParagraph.SetFontSize(24)
          oParagraph.SetColor(125, 125, 125, false)
          oParagraph.SetJc("center")
          drawDocument.AddElement(0,oParagraph)
          oDrawing.SetVerticalTextAlign("center")

          // debugger
          oDrawing.SetVerPosition('page', 914400)
          oDrawing.SetSize(MM2EMU * e.w, MM2EMU * e.h)
          // oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
          oDocument.AddDrawingToPage(oDrawing, e.page_num, MM2EMU * e.x, MM2EMU * e.y);
        } else if (e.draw_type == 'image') {
          var imgname = ''
          var imgurl = ''
          if (e.zone_type == 16) {
            imgname = 'complete'
          } else if (e.zone_type == 28) {
            imgname = 'check'
          } else if (e.zone_type == 11) {
            imgname = 'pass'
          }
          imgurl = map_base64[imgname] //  getBase64(imgname)
          oDrawing = Api.CreateImage(imgurl, 77 * 0.3 * 36000, 28 * 0.3 * 36000);
          oDocument.AddDrawingToPage(oDrawing, e.page_num, MM2EMU * e.x, MM2EMU * e.y);
        }
      }
      if (!posdata) {
        console.log('add')
        pos_list.push({
          zone_type: e.zone_type,
          v: e.v,
          drawing_id: oDrawing.Drawing.Id
        })
      } else {
        console.log('update')
        posdata.drawing_id = oDrawing.Drawing.Id
      }
    })
    return pos_list
  }, false, true, undefined)
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
    console.log('posdata', posdata)
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
      var text = ''
      if (posdata.zone_type == 15) {
        text = '再练'
      } else if (posdata.zone_type == 16) {
        text = '完成'
      }
      oParagraph.AddText(text)
      oParagraph.SetFontSize(24)
      oParagraph.SetColor(125, 125, 125, false)
      oParagraph.SetJc("center")
      drawDocument.AddElement(0,oParagraph)
      oDrawing.SetVerticalTextAlign("center")

      // debugger
      oDrawing.SetVerPosition('page', 914400)
      oDrawing.SetSize(MM2EMU * posdata.w, MM2EMU * posdata.h)
      // oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
      oDocument.AddDrawingToPage(oDrawing, posdata.page_num, MM2EMU * posdata.x, MM2EMU * posdata.y);
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
// 打分区用base64实现
function handleScoreField4(options) {
  if (!options) {
    return
  }
  var control_list = window.BiyueCustomData.control_list
  var list = []
  options.forEach(e => {
    var index = control_list.findIndex(control => {
      return e.ques_no == control.ques_no && control.regionType == 'question'
    })
    if (index >= 0) {
      var controlData = control_list[index]
      var needHandle = true
      if (controlData.score == e.score) {
        if (e.score) {
          var score_options = controlData.score_options
          if (score_options && score_options.mode == e.mode && score_options.layout == e.layout) {
            needHandle = false
          }
        }
      }
      if (needHandle) {
        var obj = Object.assign({}, e, {
          control_index: index
        })
        if (e.score) {
          obj.scores = getScores(e.score, e.mode)
        }
        list.push(obj)
      }
    }
  })
  if (list.length == 0) {
    console.log('没有要处理的题目')
    return
  }
  Asc.scope.control_list = control_list
  Asc.scope.list = list
  Asc.scope.map_base64 = map_base64
  setupPostTask(window, function(res) {
    console.log('callback for handleScoreField', res)
    if (!res) {
      return
    }
    res.forEach(e => {
      if (e.options && e.options.control_index != undefined) {
        control_list[e.options.control_index].score = e.options.score
        if (e.options.score) {
          control_list[e.options.control_index].score_options = {
            paragraph_id: e.paragraph_id,
            run_id: e.run_id,
            drawing_id: e.drawing_id,
            table_id: e.table_id,
            mode: e.options.mode,
            layout: e.options.layout
          }
        } else {
          control_list[e.options.control_index].score_options = null
        }
      }
    })
  })
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    var control_list = Asc.scope.control_list
    var list = Asc.scope.list
    var map_base64 = Asc.scope.map_base64
    console.log('list', list)
    var resList = []
    var shapes = oDocument.GetAllShapes()
    var drawings = oDocument.GetAllDrawingObjects()
    var cell_width_mm = 6
    var cell_height_mm = 6
    var MM2TWIPS = (25.4 / 72 / 20)
    var cellWidth = cell_width_mm / MM2TWIPS
    var cellHeight = cell_height_mm / MM2TWIPS
    for (var idx = 0, maxidx = list.length; idx < maxidx; ++idx) {
      var options = list[idx]
      var controlData = control_list[options.control_index]
      var control = controls.find(e => {
        return e.Sdt.GetId() == controlData.control_id
      })
      if (!control) {
        resList.push({
          ques_no: options.ques_no,
          message: '题目控件不存在'
        })
        continue
      }
      var oShape = null
      var score_options = controlData.score_options
      if (score_options && score_options.run_id) {
        var scoreParagraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(score_options.paragraph_id))
        for (var i = 0; i < scoreParagraph.Paragraph.Content.length; ++i) {
          if (scoreParagraph.Paragraph.Content[i].Id == score_options.run_id) {
            console.log('remove run')
            scoreParagraph.RemoveElement(i)
            break
          }
        }
      } else if (score_options && score_options.drawing_id) {
        oShape = shapes.find(e => {
          return e.Drawing.Id == controlData.score_options.drawing_id
        })
      } else {
        for (var i = 0, imax = shapes.length; i < imax; ++i) {
          var dtitle = shapes[i].Drawing.docPr.title
          if (dtitle && dtitle != '') {
            var titlejson = JSON.parse(dtitle)
            if (titlejson.type == 'qscore' && titlejson.ques_control_id == controlData.control_id) {
              oShape = shapes[i]
              break // 暂时假设打分区不会重复
            }
          }
        }
      }
      var oDrawing = null
      if (oShape) {
        oDrawing = drawings.find(e => {
          return e.Drawing.Id == oShape.Drawing.Id
        })
      }
      if (oDrawing) {
        oDrawing.Delete()
      }
      if (!options.score || options.score == '') { // 删除
        resList.push({
          ques_no: options.ques_no,
          code: 1,
          options: options
        })
        continue
      }
      var rect = Api.asc_GetContentControlBoundingRect(controlData.control_id, true);
      var newRect = {
        Left: rect.X0,
        Right: rect.X1,
        Top: rect.Y0,
        Bottom: rect.Y1
      }
      var controlContent = control.GetContent()
      if (controlContent) {
        var pageIndex = 0
        if (controlContent.Document && controlContent.Document.Pages && controlContent.Document.Pages.length > 1) {
          for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
            if (!control.Sdt.IsEmptyPage(p)) {
              pageIndex = p
              break
            }
          }
        }
        console.log('controlContent', controlContent)
        console.log('pageIndex', pageIndex)
        var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
        if (pagebounds) {
          newRect.Right = Math.max(pagebounds.Right, newRect.Right) 
        }
      }
      console.log(controlData.ques_no, controlData.control_id, 'rect', rect, 'pagebounds', pagebounds)
      var width = newRect.Right - newRect.Left
      console.log('newRect', newRect, 'width', width)
      var trips_width = width / MM2TWIPS
      console.log('trips_width', trips_width, options)
      var scores = options.scores
      var maxWidth = trips_width / options.layout
      var rowcount = 1
      var cellcount = scores.length
      var columncount = cellcount
      if (maxWidth < cellcount * cellWidth) { // 需要换行
        rowcount = Math.ceil(cellcount * cellWidth / maxWidth)
        columncount = Math.floor(maxWidth / cellWidth)
      }
      console.log('rowcount', rowcount, 'columncount', columncount)
      var oFill = Api.CreateNoFill()
      var oStroke = Api.CreateStroke(3600, Api.CreateNoFill());
      var shapew = columncount * cell_width_mm + 6
      var shapeh = rowcount * cell_height_mm + 3
      oDrawing = Api.CreateShape("rect",  shapew * 36E3, shapeh * 36E3, oFill, oStroke);
      var drawDocument = oDrawing.GetContent()
      var oParagraph = Api.CreateParagraph()
      for (var i = 0; i < scores.length; ++i) {
        var imgurl = map_base64['1']
        console.log(i, scores[i].v)
        var imgDrawing = Api.CreateImage(imgurl, cell_width_mm * 36E3, cell_height_mm * 36E3)
        console.log('imgDrawing width height', imgDrawing.GetWidth(), imgDrawing.GetHeight())
        oParagraph.AddDrawing(imgDrawing)
      }
      oParagraph.SetJc('right')
      oParagraph.SetColor(125, 125, 125, false)
      oParagraph.SetFontSize(24)
      oParagraph.SetFontFamily('黑体')
      console.log('oParagraph', oParagraph)
      drawDocument.AddElement(0, oParagraph)
      if (options.layout == 2) { // 嵌入式
        oDrawing.SetWrappingStyle("square")
        oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
      } else { // 顶部
        oDrawing.SetWrappingStyle("topAndBottom")
        oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
        oDrawing.SetVerPosition("paragraph", 1 * 36E3);
        // oDrawing.SetVerAlign("paragraph")
      }
      var titleobj = {
        type: 'qscore',
        ques_control_id: controlData.control_id
      }
      oDrawing.Drawing.Set_Props({
        title: JSON.stringify(titleobj)
      })
      // 下面是全局插入的
      // control.GetRange().Select()
      // oDocument.AddDrawingToPage(oDrawing, 0, (newRect.Right - shapew) * 36E3, newRect.Top * 36E3); // 用这种方式加入的一定是相对页面的
      // oDrawing.SetVerAlign("paragraph", newRect.Top * 36E3)
      // Api.asc_RemoveSelection();
      // 在题目内插入
      var oRun = Api.CreateRun();
      oRun.AddDrawing(oDrawing);
      console.log('oDrawing', oDrawing)
      var paragraphs = controlContent.GetAllParagraphs()
      console.log('paragraphs', paragraphs)
      if (paragraphs && paragraphs.length > 0) {
        paragraphs[0].AddElement(oRun, 1);
        resList.push({
          code: 1,
          ques_no: options.ques_no,
          options: options,
          paragraph_id: paragraphs[0].Paragraph.Id,
          run_id: oRun.Run.Id,
          drawing_id: oDrawing.Drawing.Id,
          run_id: oRun.Run.Id
        })
      }
    }
    return resList
  }, false, true, undefined)
}
// 打分区用添加表格单元格距离实现
function handleScoreField(options) {
  if (!options) {
    return
  }
  var control_list = window.BiyueCustomData.control_list
  var list = []
  options.forEach(e => {
    var index = control_list.findIndex(control => {
      return e.ques_no == control.ques_no && control.regionType == 'question'
    })
    if (index >= 0) {
      var controlData = control_list[index]
      var needHandle = true
      if (controlData.score == e.score) {
        if (e.score) {
          var score_options = controlData.score_options
          if (score_options && score_options.mode == e.mode && score_options.layout == e.layout) {
            needHandle = false
          }
        }
      }
      if (needHandle) {
        var obj = Object.assign({}, e, {
          control_index: index
        })
        if (e.score) {
          obj.scores = getScores(e.score, e.mode)
        }
        list.push(obj)
      }
    }
  })
  if (list.length == 0) {
    console.log('没有要处理的题目')
    return
  }
  Asc.scope.control_list = control_list
  Asc.scope.list = list
  setupPostTask(window, function(res) {
    console.log('callback for handleScoreField', res)
    if (!res) {
      return
    }
    res.forEach(e => {
      if (e.options && e.options.control_index != undefined) {
        control_list[e.options.control_index].score = e.options.score
        if (e.options.score) {
          control_list[e.options.control_index].score_options = {
            paragraph_id: e.paragraph_id,
            run_id: e.run_id,
            drawing_id: e.drawing_id,
            table_id: e.table_id,
            mode: e.options.mode,
            layout: e.options.layout
          }
        } else {
          control_list[e.options.control_index].score_options = null
        }
      }
    })
  })
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    var control_list = Asc.scope.control_list
    var list = Asc.scope.list
    console.log('list', list)
    var resList = []
    var shapes = oDocument.GetAllShapes()
    var drawings = oDocument.GetAllDrawingObjects()
    var cell_width_mm = 8 + 3
    var cell_height_mm = 8 + 3.5
    var MM2TWIPS = (25.4 / 72 / 20)
    var cellWidth = cell_width_mm / MM2TWIPS
    var cellHeight = cell_height_mm / MM2TWIPS
    for (var idx = 0, maxidx = list.length; idx < maxidx; ++idx) {
      var options = list[idx]
      var controlData = control_list[options.control_index]
      var control = controls.find(e => {
        return e.Sdt.GetId() == controlData.control_id
      })
      if (!control) {
        resList.push({
          ques_no: options.ques_no,
          message: '题目控件不存在'
        })
        continue
      }
      var oShape = null
      var score_options = controlData.score_options
      if (score_options && score_options.run_id) {
        var scoreParagraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(score_options.paragraph_id))
        for (var i = 0; i < scoreParagraph.Paragraph.Content.length; ++i) {
          if (scoreParagraph.Paragraph.Content[i].Id == score_options.run_id) {
            console.log('remove run')
            scoreParagraph.RemoveElement(i)
            break
          }
        }
      } else if (score_options && score_options.drawing_id) {
        oShape = shapes.find(e => {
          return e.Drawing.Id == controlData.score_options.drawing_id
        })
      } else {
        for (var i = 0, imax = shapes.length; i < imax; ++i) {
          var dtitle = shapes[i].Drawing.docPr.title
          if (dtitle && dtitle != '') {
            var titlejson = JSON.parse(dtitle)
            if (titlejson.type == 'qscore' && titlejson.ques_control_id == controlData.control_id) {
              oShape = shapes[i]
              break // 暂时假设打分区不会重复
            }
          }
        }
      }
      var oDrawing = null
      if (oShape) {
        oDrawing = drawings.find(e => {
          return e.Drawing.Id == oShape.Drawing.Id
        })
      }
      if (oDrawing) {
        oDrawing.Delete()
      }
      if (!options.score || options.score == '') { // 删除
        resList.push({
          ques_no: options.ques_no,
          code: 1,
          options: options
        })
        continue
      }
      var rect = Api.asc_GetContentControlBoundingRect(controlData.control_id, true);
      var newRect = {
        Left: rect.X0,
        Right: rect.X1,
        Top: rect.Y0,
        Bottom: rect.Y1
      }
      var controlContent = control.GetContent()
      if (controlContent) {
        var pageIndex = 0
        if (controlContent.Document && controlContent.Document.Pages && controlContent.Document.Pages.length > 1) {
          for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
            if (!control.Sdt.IsEmptyPage(p)) {
              pageIndex = p
              break
            }
          }
        }
        console.log('controlContent', controlContent)
        console.log('pageIndex', pageIndex)
        var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
        if (pagebounds) {
          newRect.Right = Math.max(pagebounds.Right, newRect.Right) 
        }
      }
      console.log(controlData.ques_no, controlData.control_id, 'rect', rect, 'pagebounds', pagebounds)
      var width = newRect.Right - newRect.Left
      console.log('newRect', newRect, 'width', width)
      var trips_width = width / MM2TWIPS
      console.log('trips_width', trips_width, options)
      var maxTableWidth = trips_width / options.layout
      var scores = options.scores
      var cellcount = scores.length
      var rowcount = 1
      var columncount = cellcount
      var maxTableWidth = trips_width / options.layout
      console.log('maxTableWidth', maxTableWidth, cellcount, cellWidth)
      if (maxTableWidth < cellcount * cellWidth) { // 需要换行
        rowcount = Math.ceil(cellcount * cellWidth / maxTableWidth)
        console.log('rowCount', rowcount, 'cellcount', cellcount, 'cellWidth', cellWidth, 'maxTableWidth', maxTableWidth)
        columncount = Math.floor(maxTableWidth / cellWidth)
      } else {
        console.log('无需换行')
      }
      var mergecount = rowcount * columncount - cellcount
      if (rowcount <= 0 || columncount <= 0) {
        console.log('行数或列数异常', list)
        return
      }
      var oTable = Api.CreateTable(columncount, rowcount)
      if (mergecount > 0) {
        var cells = []
        for (var k = 0; k < mergecount; ++k) {
          var cellrow = oTable.GetRow(rowcount - 1)
          if (cellrow) {
            var cell = cellrow.GetCell(k)
            if (cell) {
              cells.push(cell)
            } else {
              console.log(rowcount - 1, k, 'cell is null')
            }
          } else {
            console.log('cellrow is null', rowcount - 1, oTable.GetRowsCount())
          }
        }
        if (cells.length > 0) {
          oTable.MergeCells(cells);
        }
      }
      var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table")
      var oTableStylePr = oTableStyle.GetConditionalTableStyle("wholeTable");
      oTable.SetTableLook(true, true, true, true, true, true);
      oTableStylePr.GetTableRowPr().SetHeight("atLeast", cellHeight); // 高度至少多少trips
      var oTableCellPr = oTableStyle.GetTableCellPr();
      oTableCellPr.SetVerticalAlign("center");
      oTable.SetWrappingStyle(params.layout == 1 ? true : false);
      oTable.SetStyle(oTableStyle);
      oTable.SetCellSpacing(150);
      oTable.SetTableBorderTop("single", 1, 0.1, 255, 255, 255)
      oTable.SetTableBorderBottom("single", 1, 0.1, 255, 255, 255)
      oTable.SetTableBorderLeft("single", 1, 0.1, 255, 255, 255)
      oTable.SetTableBorderRight("single", 1, 0.1, 255, 255, 255)
      var scoreindex = -1
      // 设置单元格文本
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
                cell.SetWidth("twips", cellWidth);
                oCellPara.SetJc('center')
                oCellPara.SetColor(0, 0, 0, false)
                oCellPara.SetFontSize(16)
                scores[scoreindex].row = cr
                scores[scoreindex].column = cc
              } else {
                console.log('oCellPra is null')
              }
            } else {
              console.log('cellcontent is null')
            }
            // console.log('cell', cell)
          } else {
            console.log('cannot get cell', cc, cr)
          }
        }
      }
      oTable.SetWidth('twips', columncount * cellWidth)
      var shapew = cell_width_mm * columncount + 3
      var shapeh = (cell_height_mm * rowcount + 4)
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
      // if (oTable.SetLockValue) {
      //   oTable.SetLockValue(true)
      // }
      var oFill = Api.CreateNoFill()
      var oStroke = Api.CreateStroke(3600, Api.CreateNoFill());
      oDrawing = Api.CreateShape("rect",  shapew * 36E3, shapeh * 36E3, oFill, oStroke);
      var drawDocument = oDrawing.GetContent()
      drawDocument.AddElement(0, oTable)
      if (options.layout == 2) { // 嵌入式
        oDrawing.SetWrappingStyle("square")
        oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
      } else { // 顶部
        oDrawing.SetWrappingStyle("topAndBottom")
        oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
        oDrawing.SetVerPosition("paragraph", 1 * 36E3);
        // oDrawing.SetVerAlign("paragraph")
      }
      var titleobj = {
        type: 'qscore',
        ques_control_id: controlData.control_id
      }
      oDrawing.Drawing.Set_Props({
        title: JSON.stringify(titleobj)
      })
      // 下面是全局插入的
      // control.GetRange().Select()
      // oDocument.AddDrawingToPage(oDrawing, 0, (newRect.Right - shapew) * 36E3, newRect.Top * 36E3); // 用这种方式加入的一定是相对页面的
      // oDrawing.SetVerAlign("paragraph", newRect.Top * 36E3)
      // Api.asc_RemoveSelection();
      // 在题目内插入
      var oRun = Api.CreateRun();
      oRun.AddDrawing(oDrawing);
      var paragraphs = controlContent.GetAllParagraphs()
      console.log('paragraphs', paragraphs)
      if (paragraphs && paragraphs.length > 0) {
        paragraphs[0].AddElement(oRun, 1);
        resList.push({
          code: 1,
          ques_no: options.ques_no,
          options: options,
          paragraph_id: paragraphs[0].Paragraph.Id,
          run_id: oRun.Run.Id,
          drawing_id: oDrawing.Drawing.Id,
          table_id: oTable.Table.Id,
          run_id: oRun.Run.Id
        })
      }
    }
    return resList
  }, false, true, undefined)
}
// 打分区用添加表格单元格分割实现
function handleScoreField2(options) {
  if (!options) {
    return
  }
  var control_list = window.BiyueCustomData.control_list
  var list = []
  options.forEach(e => {
    var index = control_list.findIndex(control => {
      return e.ques_no == control.ques_no && control.regionType == 'question'
    })
    if (index >= 0) {
      var controlData = control_list[index]
      var needHandle = true
      if (controlData.score == e.score) {
        if (e.score) {
          var score_options = controlData.score_options
          if (score_options && score_options.mode == e.mode && score_options.layout == e.layout) {
            needHandle = false
          }
        }
      }
      if (needHandle) {
        var obj = Object.assign({}, e, {
          control_index: index
        })
        if (e.score) {
          obj.scores = getScores(e.score, e.mode)
        }
        list.push(obj)
      }
    }
  })
  if (list.length == 0) {
    console.log('没有要处理的题目')
    return
  }
  Asc.scope.control_list = control_list
  Asc.scope.list = list
  setupPostTask(window, function(res) {
    console.log('callback for handleScoreField', res)
    if (!res) {
      return
    }
    res.forEach(e => {
      if (e.options && e.options.control_index != undefined) {
        control_list[e.options.control_index].score = e.options.score
        if (e.options.score) {
          control_list[e.options.control_index].score_options = {
            paragraph_id: e.paragraph_id,
            run_id: e.run_id,
            drawing_id: e.drawing_id,
            table_id: e.table_id,
            mode: e.options.mode,
            layout: e.options.layout
          }
        } else {
          control_list[e.options.control_index].score_options = null
        }
      }
    })
  })
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    var control_list = Asc.scope.control_list
    var list = Asc.scope.list
    console.log('list', list)
    var resList = []
    var shapes = oDocument.GetAllShapes()
    var drawings = oDocument.GetAllDrawingObjects()
    var cell_width_mm = 8
    var cell_height_mm = 8
    var spacing_width_mm = 4
    var MM2TWIPS = (25.4 / 72 / 20)
    var cellWidth = cell_width_mm / MM2TWIPS
    var spacingWidth = spacing_width_mm / MM2TWIPS
    var cellHeight = cell_height_mm / MM2TWIPS
    for (var idx = 0, maxidx = list.length; idx < maxidx; ++idx) {
      var options = list[idx]
      var controlData = control_list[options.control_index]
      var control = controls.find(e => {
        return e.Sdt.GetId() == controlData.control_id
      })
      if (!control) {
        resList.push({
          ques_no: options.ques_no,
          message: '题目控件不存在'
        })
        continue
      }
      var oShape = null
      var score_options = controlData.score_options
      if (score_options && score_options.run_id) {
        var scoreParagraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(score_options.paragraph_id))
        for (var i = 0; i < scoreParagraph.Paragraph.Content.length; ++i) {
          if (scoreParagraph.Paragraph.Content[i].Id == score_options.run_id) {
            console.log('remove run')
            scoreParagraph.RemoveElement(i)
            break
          }
        }
      } else if (score_options && score_options.drawing_id) {
        oShape = shapes.find(e => {
          return e.Drawing.Id == controlData.score_options.drawing_id
        })
      } else {
        for (var i = 0, imax = shapes.length; i < imax; ++i) {
          var dtitle = shapes[i].Drawing.docPr.title
          if (dtitle && dtitle != '') {
            var titlejson = JSON.parse(dtitle)
            if (titlejson.type == 'qscore' && titlejson.ques_control_id == controlData.control_id) {
              oShape = shapes[i]
              break // 暂时假设打分区不会重复
            }
          }
        }
      }
      var oDrawing = null
      if (oShape) {
        oDrawing = drawings.find(e => {
          return e.Drawing.Id == oShape.Drawing.Id
        })
      }
      if (oDrawing) {
        oDrawing.Delete()
      }
      if (!options.score || options.score == '') { // 删除
        resList.push({
          ques_no: options.ques_no,
          code: 1,
          options: options
        })
        continue
      }
      var rect = Api.asc_GetContentControlBoundingRect(controlData.control_id, true);
      var newRect = {
        Left: rect.X0,
        Right: rect.X1,
        Top: rect.Y0,
        Bottom: rect.Y1
      }
      var controlContent = control.GetContent()
      if (controlContent) {
        var pageIndex = 0
        if (controlContent.Document && controlContent.Document.Pages && controlContent.Document.Pages.length > 1) {
          for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
            if (!control.Sdt.IsEmptyPage(p)) {
              pageIndex = p
              break
            }
          }
        }
        console.log('controlContent', controlContent)
        console.log('pageIndex', pageIndex)
        var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
        if (pagebounds) {
          newRect.Right = Math.max(pagebounds.Right, newRect.Right) 
        }
      }
      console.log(controlData.ques_no, controlData.control_id, 'rect', rect, 'pagebounds', pagebounds)
      var width = newRect.Right - newRect.Left
      console.log('newRect', newRect, 'width', width)
      var trips_width = width / MM2TWIPS
      console.log('trips_width', trips_width, options)
      var scores = options.scores
      var scoreCount = scores.length
      var maxTableWidth = trips_width / options.layout
      var rowcount = 1
      var columncount = scoreCount * 2 - 1
      console.log('maxTableWidth', maxTableWidth, (scoreCount * (cellWidth + spacingWidth)))
      var fillCountARow = scoreCount // 每行真正填充分数的格子数
      if (maxTableWidth < (scoreCount * (cellWidth + spacingWidth))) { // 需要换行
        rowcount = Math.ceil(scoreCount * (cellWidth + spacingWidth) / maxTableWidth)
        var x = Math.floor((maxTableWidth + spacingWidth) / (cellWidth + spacingWidth))
        columncount = 2 * x - 1
        fillCountARow = x
      }
      console.log('rowcount', rowcount, 'columncount', columncount)
      if (rowcount <= 0 || columncount <= 0) {
        console.log('行数或列数异常', list)
        return
      }
      var fillRowCount = rowcount // 有填充分数的行数
      rowcount = fillRowCount + Math.floor(fillRowCount / 2)
      console.log('fillCountARow', fillCountARow, 'fillRowCount', fillRowCount, 'rowcount', rowcount, 'columncount', columncount)
      var oTable = Api.CreateTable(columncount, rowcount)
      var mergecount = (fillCountARow * fillRowCount - scoreCount) * 2
      for (var r = 0; r < rowcount; ++r) {
        var orow = oTable.GetRow(r)
        if (r % 2 > 0) {
          var mcell = orow.MergeCells()
          if (mcell) {
            mcell.SetCellBorderLeft('single', 1, 0.1, 255, 255, 255)
            mcell.SetCellBorderRight('single', 1, 0.1, 255, 255, 255)
            mcell.SetCellBorderTop('single', 1, 0.1, 255, 255, 255)
            mcell.SetCellBorderBottom('single', 1, 0.1, 255, 255, 255)
          }
          orow.SetHeight("atLeast", 4 / MM2TWIPS)
        } else if (r == rowcount - 1 && mergecount > 0) {
          var cells = []
          for (var k = 0; k < mergecount; ++k) {
            var cellrow = oTable.GetRow(rowcount - 1)
            if (cellrow) {
              var cell = cellrow.GetCell(k)
              if (cell) {
                cells.push(cell)
              } else {
                console.log(rowcount - 1, k, 'cell is null')
              }
            } else {
              console.log('cellrow is null', rowcount - 1, oTable.GetRowsCount())
            }
          }
          if (cells.length > 0) {
            var mcell = oTable.MergeCells(cells);
            if (mcell) {
              mcell.SetCellBorderLeft('single', 1, 0.1, 255, 255, 255)
              mcell.SetCellBorderRight('single', 1, 0.1, 255, 255, 255)
              mcell.SetCellBorderTop('single', 1, 0.1, 255, 255, 255)
              mcell.SetCellBorderBottom('single', 1, 0.1, 255, 255, 255)
            }
          }  
        }
      }
      var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table")
      var oTableStylePr = oTableStyle.GetConditionalTableStyle("wholeTable");
      oTable.SetTableLook(true, true, true, true, true, true);
      oTableStylePr.GetTableRowPr().SetHeight("atLeast", cellHeight); // 高度至少多少trips
      var oTableCellPr = oTableStyle.GetTableCellPr();
      oTableCellPr.SetVerticalAlign("center");
      oTable.SetWrappingStyle(params.layout == 1 ? true : false);
      oTable.SetStyle(oTableStyle);
      // oTable.SetCellSpacing(150);
      oTable.SetTableBorderTop("single", 0.1, 0.1, 255, 255, 255)
      oTable.SetTableBorderBottom("single", 0.1, 0.1, 255, 255, 255)
      oTable.SetTableBorderLeft("single", 0.1, 0.1, 255, 255, 255)
      oTable.SetTableBorderRight("single", 0.1, 0.1, 255, 255, 255)
      var scoreindex = -1
      // 设置单元格文本
      for (var irow = 0; irow < rowcount; ++irow) {
        var cbegin = 0
        var cend = columncount
        if (mergecount > 0 && irow == rowcount - 1) { // 最后一行
          cbegin = 1
          cend = columncount - mergecount + 1
        }
        console.log('irow', irow, 'cbegin', cbegin, 'cend', cend)
        var roww = 0
        var ww = 0
        var sw = 0
        var w = 0
        if (irow % 2 > 0) {
          continue
        }
        for (var icolumn = cbegin; icolumn < cend; ++icolumn) {
          var cr = irow
          var cc = icolumn
          var cell = oTable.GetCell(cr, cc)
          if (!cell) {
            break
          }
          var fillScore = irow % 2 == 0 && cc % 2 == 0
          if (irow == rowcount - 1 && mergecount > 0) {
            fillScore = cc % 2 > 0
          } 
          var cellcontent = cell.GetContent()
          if (fillScore) {
            scoreindex++
            if (scoreindex >= scores.length) {
              break
            }
          }
          console.log('cell', cr, cc, fillScore)
          var oCellPara = cellcontent.GetElement(0)
          if (oCellPara) {
            oCellPara.AddText(fillScore ? scores[scoreindex].v : '')
            if (fillScore) {
              cell.SetCellBorderLeft('single', 1, 0.1, 212, 212, 212)
              cell.SetCellBorderRight('single', 1, 0.1, 212, 212, 212)
              cell.SetCellBorderTop('single', 1, 0.1, 212, 212, 212)
              cell.SetCellBorderBottom('single', 1, 0.1, 212, 212, 212)
              cell.SetWidth("twips", cellWidth);
              oCellPara.SetJc('center')
              oCellPara.SetColor(0, 0, 0, false)
              oCellPara.SetFontSize(16)
              scores[scoreindex].row = cr
              scores[scoreindex].column = cc
              console.log('cr', cr, 'cc', cc, scores[scoreindex].v, cellWidth)
              roww += cell_width_mm
              ww++
            } else {
              cell.SetWidth("twips", spacingWidth);
              roww += spacing_width_mm
              sw++
            }
            w += cell.CellPr.TableCellW.W
          }
        }
        console.log('roww', roww, 'ww', ww, 'sw', sw, w)
      }
      var tablew = cell_width_mm * (columncount + 1) / 2 + spacing_width_mm * (columncount - 1) / 2
      oTable.SetWidth('twips', tablew / MM2TWIPS)
      console.log('tablewidth', columncount * (cellWidth + spacingWidth) - spacingWidth, 100 / MM2TWIPS)
      var shapew = tablew + 3
      console.log('shapew', shapew - 3)
      var shapeh = (cell_height_mm * fillRowCount + (fillRowCount - 1) * 4 + 4)
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
      // if (oTable.SetLockValue) {
      //   oTable.SetLockValue(true)
      // }
      var oFill = Api.CreateNoFill()
      var oStroke = Api.CreateStroke(3600, Api.CreateNoFill());
      oDrawing = Api.CreateShape("rect",  shapew * 36E3, shapeh * 36E3, oFill, oStroke);
      var drawDocument = oDrawing.GetContent()
      drawDocument.AddElement(0, oTable)
      if (options.layout == 2) { // 嵌入式
        oDrawing.SetWrappingStyle("square")
        oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
      } else { // 顶部
        oDrawing.SetWrappingStyle("topAndBottom")
        oDrawing.SetHorPosition("column", (width - shapew) * 36E3);
        oDrawing.SetVerPosition("paragraph", 1 * 36E3);
        // oDrawing.SetVerAlign("paragraph")
      }
      var titleobj = {
        type: 'qscore',
        ques_control_id: controlData.control_id
      }
      oDrawing.Drawing.Set_Props({
        title: JSON.stringify(titleobj)
      })
      // 下面是全局插入的
      // control.GetRange().Select()
      // oDocument.AddDrawingToPage(oDrawing, 0, (newRect.Right - shapew) * 36E3, newRect.Top * 36E3); // 用这种方式加入的一定是相对页面的
      // oDrawing.SetVerAlign("paragraph", newRect.Top * 36E3)
      // Api.asc_RemoveSelection();
      // 在题目内插入
      var oRun = Api.CreateRun();
      oRun.AddDrawing(oDrawing);
      var paragraphs = controlContent.GetAllParagraphs()
      console.log('paragraphs', paragraphs)
      if (paragraphs && paragraphs.length > 0) {
        paragraphs[0].AddElement(oRun, 1);
        resList.push({
          code: 1,
          ques_no: options.ques_no,
          options: options,
          paragraph_id: paragraphs[0].Paragraph.Id,
          run_id: oRun.Run.Id,
          drawing_id: oDrawing.Drawing.Id,
          table_id: oTable.Table.Id,
          run_id: oRun.Run.Id
        })
      }
    }
    return resList
  }, false, true, undefined)
}

function addImage() {
  // toggleWeight()
  // return
  Asc.scope.map_base64 = map_base64
  window.Asc.plugin.callCommand(function () {
    var oDocument = Api.GetDocument();
    let controls = oDocument.GetAllContentControls();
    var map_base64 = Asc.scope.map_base64
    for (var i = 0; i < controls.length; ++i) {
      var control = controls[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'question') {
        var imgurl = map_base64['1'] //   'https://by-base-cdn.biyue.tech/check.svg'
        var oDrawing = Api.CreateImage(imgurl, 8 * 36000, 8 * 36000);
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
// 显示小问序号
function showAskIndex() {
  var control_list = window.BiyueCustomData.control_list
  Asc.scope.control_list = control_list
  setupPostTask(window, function(res) {
    console.log('result of showAskIndex:', res)
  })
  window.Asc.plugin.callCommand(function() {
    var control_list = Asc.scope.control_list
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    for (var i = 0, imax = control_list.length; i < imax; ++i) {
      var control = control_list[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'question') {

      }
    }
  }, false, true, undefined)
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
      if (control_list) {
        var find = control_list.find(e => {
          return e.control_id == controlId
        })
        Asc.scope.find_controldata = find
      }
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
    if (controldata) {
      // if (control && control.GetAllDrawingObjects) {
      //   var drawingObjs = control.GetAllDrawingObjects()
      //   for (var i = 0, imax = drawingObjs.length; i < imax; ++i) {
      //     var oDrawing = drawingObjs[i]
      //     if (oDrawing.Drawing.docPr.title == 'ask_weight') {
      //       oDrawing.Delete()
      //     }
      //   }
      // }
    }
  }, false, true, undefined)

}
// 添加或删除标识
function handleIdentifyBox(add) {
  setupPostTask(window, function(res) {
    console.log('handleIdentifyBox result', res)
    if (!res.code) {
      return
    }
    var control_list = window.BiyueCustomData.control_list
    var control = control_list.find(e => {
      return e.control_id == res.control_id
    })
    if (!control) {
      return
    }
    if (add) {
      if (!control.identify_list) {
        control.identify_list = []
      }
      control.identify_list.push({
        add_pos: res.add_pos,
        run_id: res.run_id,
        paragraph_id: res.paragraph_id,
        drawing_id: res.drawing_id
      })
    } else if (res.remove_ids && res.remove_ids.length > 0) {
      for (var i = 0; i < res.remove_ids.length; ++i) {
        var index = control.identify_list.findIndex(e => {
          return e.drawing_id == res.remove_ids[i]
        })
        if (index >= 0) {
          control.identify_list.splice(index, 1)
        }
      }
    }
  })
  Asc.scope.add = add
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument()
    var curPosInfo = oDocument.Document.GetContentPosition()
    var add = Asc.scope.add
    console.log('curPosInfo', curPosInfo)
    var res = {
      code: 0
    }
    if (curPosInfo) {
      var runIdx  = -1
      var paragraphIdx = -1
      for (var i = curPosInfo.length - 1; i >= 0; --i) {
        if (curPosInfo[i].Class.GetType) {
          var t = curPosInfo[i].Class.GetType()
          if (t == 1) {
            paragraphIdx = i
            break
          } else if (t == 39) {
            runIdx = i
          }
        }
      }
      function createDrawing() {
        var oFill = Api.CreateNoFill()
        var oStroke = Api.CreateStroke(3600, Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125)));
        var oDrawing = Api.CreateShape("rect",  8 * 36E3, 5 * 36E3, oFill, oStroke);
        var drawDocument = oDrawing.GetContent()
        var oParagraph = Api.CreateParagraph()
        oParagraph.AddText("×");
        oParagraph.SetColor(125, 125, 125, false)
        oParagraph.SetFontSize(24)
        oParagraph.SetFontFamily('黑体')
        oParagraph.SetJc('center')
        console.log('oParagraph', oParagraph.Paragraph.Id, oDrawing.Drawing.Id)
        drawDocument.AddElement(0, oParagraph)
        oDrawing.SetPaddings(0, 0, 0, 0.5 * 36E3)
        var titleobj = {
          type: 'quesIdentify',
          control_id: paraentControl.Sdt.GetId()
        }
        oDrawing.Drawing.Set_Props({
          title: JSON.stringify(titleobj)
        })
        oDrawing.SetWrappingStyle("inline")
        return oDrawing
      }
      if (paragraphIdx >= 0) {
        var pParagraph = new Api.private_CreateApiParagraph(AscCommon.g_oTableId.Get_ById(curPosInfo[paragraphIdx].Class.Id))
        var paraentControl = pParagraph.GetParentContentControl()
        if (paraentControl) {
          var tag = JSON.parse(paraentControl.GetTag())
          if (tag.regionType == 'question' || tag.regionType == 'sub-question') {
            res.control_id = paraentControl.Sdt.GetId()
            if (add) {
              var oDrawing = createDrawing()
              res.drawing_id = oDrawing.Drawing.Id
              res.paragraph_id = pParagraph.Paragraph.Id
              res.code = 1
              if (runIdx >= 0) {
                res.run_id = curPosInfo[runIdx].Class.Id
                curPosInfo[runIdx].Class.Add_ToContent(curPosInfo[runIdx].Position, oDrawing.Drawing);
                res.add_pos = 'run'
              } else {
                var oRun = Api.CreateRun();
                oRun.AddDrawing(oDrawing);
                res.run_id = oRun.Run.Id
                res.add_pos = 'paragraph'
                pParagraph.AddElement(oRun, curPosInfo[paragraphIdx].Position + 1);
              }
            } else {
              res.code = 1
              res.remove_ids = []
              var drawings = paraentControl.GetAllDrawingObjects()
              for (var sidx = 0; sidx < drawings.length; ++sidx) {
                if (drawings[sidx].Drawing && drawings[sidx].Drawing.docPr && 
                  drawings[sidx].Drawing.docPr.title && drawings[sidx].Drawing.docPr.title != '') {
                  try {
                    var dtitle = JSON.parse(drawings[sidx].Drawing.docPr.title)
                    if (dtitle.type == 'quesIdentify') {
                      res.remove_ids.push(drawings[sidx].Drawing.Id)
                      drawings[sidx].Delete()
                    }
                  } catch (error) {
                    
                  }
                }
              }
            }
          }
        }
      }
    }
    return res
  }, false, true, undefined)
}

function showIdentifyIndex(show) {
  Asc.scope.show = show
  Asc.scope.control_list = window.BiyueCustomData.control_list
  window.Asc.plugin.callCommand(function() {
    var show = Asc.scope.show
    var oDocument = Api.GetDocument()
    var shapes = oDocument.GetAllShapes()
    var control_list = Asc.scope.control_list
    console.log('00000')
    function updateFill(drawing, oFill) {
      if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill") {
        return false;
      }
      drawing.GraphicObj.spPr.setFill(oFill.UniFill);
    }
    var quesnum = 0
    var quescontrol = null
    for (var i = 0, imax = shapes.length; i < imax; ++i) {
      var drawingObj = shapes[i]
      if (drawingObj.Drawing && drawingObj.Drawing.docPr && drawingObj.Drawing.docPr.title && drawingObj.Drawing.docPr.title != '') {
        try {
          var dtitle = JSON.parse(drawingObj.Drawing.docPr.title)
          if (dtitle.type == 'quesIdentify') {
            console.log(i, 'drawingObj', drawingObj)
            var drawDocument = drawingObj.GetContent()
            var oParagraph = drawDocument.GetElement(0)
            if (oParagraph && oParagraph.GetClassType() == 'paragraph') {
              oParagraph.RemoveAllElements()
              if (show) {
                var cindex = control_list.findIndex(e => {
                  return e.control_id == dtitle.control_id
                })
                var qid = 0
                if (cindex >= 0) {
                  if (control_list[cindex].regionType == 'question') {
                    qid = control_list[cindex].control_id
                  } else {
                    qid = control_list[cindex].parent_ques_control_id
                  }
                }
                if (qid == quescontrol) {
                  quesnum += 1
                } else {
                  quescontrol = qid
                  quesnum = 1
                }
                oParagraph.AddText(quesnum + '');
                oParagraph.SetColor(255, 0, 0, false)
              } else {
                oParagraph.AddText("×");
                oParagraph.SetColor(125, 125, 125, false)
              }
              oParagraph.SetFontFamily('黑体')
              oParagraph.SetFontSize(24)
              oParagraph.SetJc('center')
            }
            if (show) {
              var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
              oFill.UniFill.transparent = 255 * 0.2 // 透明度
              updateFill(drawingObj.Drawing, oFill)
            } else {
              updateFill(drawingObj.Drawing, Api.CreateNoFill())
           }
          }
        } catch (error) {

        }
      }
    }
  }, false, true, undefined)
}

function removeAllIdentify() {
  var control_list = window.BiyueCustomData.control_list
  control_list.forEach(control => {
    if (control.identify_list) {
      control.identify_list = []
    }
  })
  window.Asc.plugin.callCommand(function() {
    var oDocument = Api.GetDocument()
    var drawingObjs = oDocument.GetAllDrawingObjects()
    for (var i = 0, imax = drawingObjs.length; i < imax; ++i) {
      var drawingObj = drawingObjs[i]
      if (drawingObj.Drawing && drawingObj.Drawing.docPr && drawingObj.Drawing.docPr.title && drawingObj.Drawing.docPr.title != '') {
        try {
          var dtitle = JSON.parse(drawingObj.Drawing.docPr.title)
          if (dtitle.type == 'quesIdentify') {
            drawingObj.Delete()
          }
        } catch (error) {
          
        }
      }
    }
  }, false, true, undefined)
}

function showWriteIdentifyIndex(show) {
  Asc.scope.show = show
  setupPostTask(window, function(res) {
    console.log('showWriteIdentifyIndex result:', res)
    if (!res) {
      return
    }
    var control_list = window.BiyueCustomData.control_list
    control_list.forEach(control => {
      if (control.regionType == 'write') {
        control.identify_list = res[control.control_id] || null
      }
    })
  })
  window.Asc.plugin.callCommand(function() {
    var show = Asc.scope.show
    var oDocument = Api.GetDocument()
    var controls = oDocument.GetAllContentControls()
    var list = {}
    function getParentQuesControl(control) {
      var parentControl = control.GetParentContentControl()
      if (parentControl) {
        try {
          var tag = JSON.parse(parentControl.GetTag())
          if (tag.regionType == 'question') {
            return parentControl
          } else if (tag.regionType == 'write' || tag.regionType == 'sub-question') {
            return getParentQuesControl(parentControl)
          }
        } catch (error) {}
      }
      return null
    }
    var quesnum = 0
    var quesControlId = 0
    for (var i = 0, imax = controls.length; i < imax; ++i) {
      var control = controls[i]
      var tag = JSON.parse(control.GetTag())
      if (tag.regionType == 'write' && tag.mode == 3) {
        var parentControl = getParentQuesControl(control)
        if (!parentControl) {
          continue
        }
        // if (parentControl.Sdt.GetId() != '6_2663') {
        //   continue
        // }
        if (show) {
          var rect = Api.asc_GetContentControlBoundingRect(control.Sdt.GetId(), true)
          console.log('rect', rect)
          console.log('control', control)
          var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
          oFill.UniFill.transparent = 255 * 0.2 // 透明度
          var oStroke = Api.CreateStroke(3600, Api.CreateNoFill());
          var width = rect.X1 - rect.X0
          var height = rect.Y1 - rect.Y0
          var oDrawing = Api.CreateShape("rect", width * 36E3, height * 36E3, oFill, oStroke);
          oDrawing.SetPaddings(0, 0, 0, 0);
          oDrawing.SetWrappingStyle('inFront')
          var x = rect.X0
          if (parentControl) {
            var prarentRect = Api.asc_GetContentControlBoundingRect(parentControl.Sdt.GetId(), true)
            x = (rect.X0 - prarentRect.X0)
          }
          console.log('x', x, prarentRect, rect)
          oDrawing.SetHorPosition("column", x * 36E3)
          oDrawing.SetVerPosition("paragraph", 1 * 36E3);
          oDrawing.Drawing.Set_Props({
            title: JSON.stringify({
              type: 'askIdentify',
              control_id: control.Sdt.GetId()
            })
          })
          var drawDocument = oDrawing.GetContent()
          var oParagraph = Api.CreateParagraph();
          if (quesControlId != parentControl.Sdt.GetId()) {
            quesnum = 1
          } else {
            quesnum++
          }
          oParagraph.AddText(quesnum + '')
          drawDocument.AddElement(0,oParagraph)
          oParagraph.SetJc("center");
          var oTextPr = Api.CreateTextPr();
          oTextPr.SetColor(255, 111, 61, false)
          oTextPr.SetFontSize(20)
          oParagraph.SetTextPr(oTextPr);
          oDrawing.SetVerticalTextAlign("center")

          var oRun = Api.CreateRun();
          oRun.AddDrawing(oDrawing);
          control.AddElement(oRun)
          list[control.Sdt.GetId()] = {
            run_id: oRun.Run.Id,
            drawing_id: oDrawing.Drawing.Id
          }
        } else {
          var drawingObjs = parentControl.GetAllDrawingObjects()
          for (var j = 0, jmax = drawingObjs.length; j < jmax; ++j) {
            var drawingObj = drawingObjs[j]
            if (drawingObj.Drawing && drawingObj.Drawing.docPr && drawingObj.Drawing.docPr.title && drawingObj.Drawing.docPr.title != '') {
              try {
                var dtitle = JSON.parse(drawingObj.Drawing.docPr.title)
                if (dtitle.type == 'askIdentify') {
                  drawingObj.Delete()
                  list[control.Sdt.GetId()] = null
                }
              } catch (error) {
                
              }
            }
          }
        }
        quesControlId = parentControl.Sdt.GetId()
        // return list
      }
    }
    return list
  }, false, true, undefined)
}

function deletePositions(list) {
  Asc.scope.pos_delete_list = list
  var pos_list = window.BiyueCustomData.pos_list
  Asc.scope.pos_list = pos_list
  setupPostTask(window, function(res) {
    console.log('deletePositions result:', res)
    if (res) {
      res.forEach(e => {
        var index = pos_list.findIndex(pos => {
          return pos.zone_type == e.zone_type && pos.v == e.v
        })
        if (index >= 0) {
          pos_list.splice(index, 1)
        }
      })
    }
  })
  window.Asc.plugin.callCommand(function() {
    var delete_list = Asc.scope.pos_delete_list || []
    var pos_list = Asc.scope.pos_list || []
    console.log('delete_list', delete_list)
    console.log('pos_list', pos_list)
    var oDocument = Api.GetDocument()
    var objs = oDocument.GetAllDrawingObjects()
    delete_list.forEach(e => {
      var posdata = pos_list.find(pos => {
        return pos.zone_type == e.zone_type && pos.v == e.v
      })
      if (posdata && posdata.drawing_id) {
        var oDrawing = objs.find(obj => {
          return obj.Drawing.Id == posdata.drawing_id
        })
        if (oDrawing) {
          oDrawing.Delete()
        } else {
          console.log('cannot find oDrawing')
        }
      }
    })
    return delete_list
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
  addImage,
  addMarkField,
  handleContentControlChange,
  handleScoreField,
  handleIdentifyBox,
  showIdentifyIndex,
  removeAllIdentify,
  showWriteIdentifyIndex,
  drawPositions,
  deletePositions
}