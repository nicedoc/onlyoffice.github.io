// 单题详情
var questionTypes = [
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
var proportionTypes = [
  { value: '0', label: '默认' },
  { value: '1', label: '1/2' },
  { value: '2', label: '1/3' },
  { value: '3', label: '1/4' },
  { value: '4', label: '1/5' },
  { value: '5', label: '1/6' },
  { value: '6', label: '1/7' },
  { value: '7', label: '1/8' }
]
var ques_control_id = null
var inited = false
var ques_control_index = null

function init() {
  if (!inited) {
    var selectElement = document.getElementById("questionType");
    questionTypes.forEach(q => {
      var option = document.createElement("option");
      if (q) {
        option.text = q.label || '';
        option.value = q.value || '';
        selectElement.add(option);
      }
    })
    var proportionElement = document.getElementById('proportion')
    proportionTypes.forEach(p => {
      var option = document.createElement("option");
      if (p) {
        option.text = p.label || '';
        option.value = p.value || '';
        proportionElement.add(option);
      }
    })
    addInputEvent('#ques_weight')
    selectElement.selectedIndex = 0
    selectElement.addEventListener('change', function() {
      const selectedValue = selectElement.options[selectElement.selectedIndex].value
      window.BiyueCustomData.control_list[ques_control_index].ques_type = selectedValue
    })
    proportionElement.selectedIndex = 0
    proportionElement.addEventListener('change', function() {
      // todo..
    })
    inited = true
  }
}

function addInputEvent(inputid, index) {
  $(inputid).on('input', function () {
    if (inputid == '#ques_weight') {
      window.BiyueCustomData.control_list[ques_control_index].score = getScore($(inputid).val())
    } else if (index != undefined) {
      window.BiyueCustomData.control_list[ques_control_index].ask_controls[index].score = getScore($(inputid).val())
    }
  })
}

function getScore(str) {
  if (!str || str == '') {
    return 0
  } else {
    return str * 1
  }
}
function showQuesData(control_id, reginType) {
  ques_control_id = control_id
  init()
  var control_list = window.BiyueCustomData.control_list || []
  var control_index = control_list.findIndex(e => {
    return e.control_id == control_id
  })
  if (control_index < 0) {
    return
  }
  if (reginType == 'sub-question' || reginType == 'write') {
    var quesCtrlId = control_list[control_index].parent_ques_control_id
    control_index = control_list.findIndex(e => {
      return e.control_id == quesCtrlId
    })
    if (control_index < 0) {
      return
    }
  }
  ques_control_index = control_index
  var controlData = control_list[control_index]
  console.log('singleQues controlData', controlData)
  var selectElement = document.getElementById("questionType");
  if (selectElement) {
    var qtype = controlData.ques_type
    var qindex = questionTypes.findIndex(e => {
      return e.value == qtype
    })
    selectElement.selectedIndex = qindex
    
  }
  $(`#ques_weight`).val(controlData.score ? controlData.score : '')
  if (controlData.ask_controls && controlData.ask_controls.length) {
    $('#askWeightBox').show()
    var content = ''
    controlData.ask_controls.forEach((ask, index) => {
      content += `<div class="item"><span class="asklabel">(${index + 1})</span><span class="score"><div class="el-input"><input id=askinput${ask.control_id} value="${ask.v || ''}"  type="text" autocomplete="off" class="el-input__inner"></div></span></div>`
    })
    $('.asks').html(content)
    controlData.ask_controls.forEach((ask, index) => {
      $(`#askinput${ask.control_id}`).val(ask.score ? ask.score : '')
      addInputEvent(`#askinput${ask.control_id}`, index)
    })
  } else {
    $('#askWeightBox').hide()
  }
}

function initListener() {
  document.addEventListener('clickSingleQues', function(event) {
    console.log('receive clickSingleQues', event.detail)
    showQuesData(event.detail.control_id, event.detail.regionType)
  })
}

export {
  initListener,
  showQuesData
}