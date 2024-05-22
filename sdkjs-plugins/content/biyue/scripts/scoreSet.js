import { setXToken } from "./auth.js";
import { questionBatchEditScore } from './api/paper.js'
(function (window, undefined) {
  var layout = 'exam_exercise';
  var sumScore = 0;
  var bonusScore = 0;
  var NoScoreSet = false;
  var batchValue = '';
  var tree = []
  var source_data = {}
  window.Asc.plugin.init = function() {
    console.log("scoreSet init");
    window.Asc.plugin.sendToPlugin("onWindowMessage", {type: 'BiyueMessage'});
  }

  function init() {
    var headerEl = document.getElementById('scoreSetHeader');
    if (headerEl) {
      var content = '当前试卷总';
      content += layout === 'exam_exercise' ? '权重' : '分'
      content += `<span id="sumScore" style="font-weight: 700;font-size: 20px;color: rgb(43, 132, 221);padding:0 8px">`;
      content += sumScore ? (sumScore % 1 > 0 ? sumScore.toFixed(1) : sumScore) : 0;
      content += '</span>';
      if (layout !== 'exam_exercise') {
        content += '分';
      }

      if (bonusScore > 0) {
        content += '+<span style="color: #f2a839;">' + bonusScore + '</span>';
        if (layout !== 'exam_exercise') {
          content += '分';
        }
        content += '(附加题)';
      }

      if (NoScoreSet) {
        content += '设置';
        content += layout === 'exam_exercise' ? '0权重' : '0分';
        content += '题为';
        content += '<span class="score"><input type="text" value="' + batchValue + '" onchange="checkBatchValue()" /></span>';
        content += layout === 'exam_exercise' ? '' : '分';
        content += '<button onclick="batchBtn()">设置</button>';
      }

      if (checkScoreTemplate() && checkScoreTemplate().value && layout !== 'exam_exercise') {
        content += '<button onclick="setScoresTemplate()">' + checkScoreTemplate().label + '分数模板</button>';
      }

      headerEl.innerHTML = content;
    }
    var tipEl = document.getElementById('scoreSetTip')
    if (tipEl) {
      tipEl.innerHTML = `提示：点击即可修改单题${layout === 'exam_exercise'?'权重':'分数'}，橙色虚线为附加题。`
    }
    var treeEl = document.getElementById('tree');
    if (treeEl) {
      var treeContent = ''
      if (tree) {
        for (var i = 0; i < tree.length; i++) {
          treeContent += '<div class="questions-title">'
          treeContent += `<span>${tree[i].struct_name}</span>`
          treeContent += '<span class="batch-set-score" style="display: inline-block;">'
          treeContent += '批量设置'
          treeContent += `<span class="score"><div class="el-input"><input id=batchinput${i} type="text" autocomplete="off" class="el-input__inner"></div></span>分<button id=batchset${i} type="button" class="el-button el-button--primary el-button--mini is-plain" style="margin-left:8px"><span>设置</span></button></span>`
          treeContent += '</div>'
          treeContent += '<div class="qustions">'
          for (var j = 0; j < tree[i].question_list.length; ++j) {
            treeContent += '<div class="ques">'
            treeContent += `<span class="no">${tree[i].question_list[j].ques_no}.</span>`
            tree[i].question_list[j].asks.forEach(ask => {
              treeContent += `<span class="score"><div class="el-input"><input id=askinput${ask.control_id} value="${ask.v || ''}"  type="text" autocomplete="off" class="el-input__inner"></div></span>`
                treeContent += `<span>分</span>`
            })
            treeContent += '</div>'
          }
          treeContent += "</div>"
        }
      }
      
      treeEl.innerHTML = treeContent
      $('#sumScore').html(sumScore)
      $('#confirm').on('click', onConfirm)
      setTimeout(() => {
        addEventListener()
        $('.title').css('color', 'red')
      }, 200)
    } else {
      console.log('tree is not find')
    }
  }  

  function addEventListener() {
    if (!tree) return
    tree.forEach((struct, istruct) => {
      addSetBtnEvent(istruct)
      struct.question_list.forEach((ques, jques) => {
        ques.asks.forEach((ask, kask) => {
          addInputEvent(`#askinput${ask.control_id}`, istruct, jques, kask)
        })
      })
    })
  }

  function addSetBtnEvent(index) {
    $(`#batchset${index}`).on('click', function () {
      var value = $(`#batchinput${index}`).val()
      if (!value) return
      if (tree && tree[index] && tree[index].question_list) {
        tree[index].question_list.forEach((ques) => {
          ques.asks.forEach((ask) => {
            ask.v = value
            $(`#askinput${ask.control_id}`).val(value)
          })
        })
      }
      updateSum()
    })
  }

  function getScore(str) {
    if (!str || str == '') {
      return 0
    } else {
      return str * 1
    }
  }

  function addInputEvent(inputid, istruct, jques, kask) {
    $(inputid).on('input', function () {
      if (kask != undefined) {
        tree[istruct].question_list[jques].asks[kask].v = getScore($(inputid).val())
      }
      updateSum()
    })
  }

  function updateSum() {
    console.log('updateSum')
    var sum = 0
    if (tree) {
      tree.forEach((struct) => {
        struct.question_list.forEach((ques) => {
          if (ques.asks) {
            ques.asks.forEach((ask) => {
              if (ask.v) {
                sum += (ask.v * 1)
              }
            })
          }
        })
      })
      sumScore = sum
      $('#sumScore').html(sum)
    }
  }

  function checkBatchValue() {
    // Add your logic here
  }
    
  function batchBtn() {
  // Add your logic here
  }
    
  function checkScoreTemplate() {
  // Add your logic here
  }
  
  function setScoresTemplate() {
  // Add your logic here
  }

  function setText(id, text) {
    var element = document.getElementById(id);
    if (element) {
      element.innerHTML = text
    }    
  }

  function onConfirm() {
    var scores = {}
    tree.forEach(struct => {
      if (struct.question_list) {
        struct.question_list.forEach(e => {
          let score = 0
          e.asks.forEach((ask, index) => {
            score += getScore(ask.v)
            if (ask.is_ask) {
              source_data.control_list[e.control_index].ask_controls[index].v = ask.v
            }
          })
          scores[e.uuid] = score + ''
          source_data.control_list[e.control_index].score = score
        })
      }
    })
    console.log('queding', scores)
    questionBatchEditScore({ paper_uuid: source_data.paper_uuid, ques_score: scores }).then(res => {
      console.log('设置分数成功')
      // 将窗口的信息传递出去
      window.Asc.plugin.sendToPlugin("onWindowMessage", {type: 'scoreSetSuccess', data: source_data});
    }).catch(error => {
      console.log(error)
    })
  }

window.Asc.plugin.attachEvent("initInfo", function(message) {
  console.log('接收的消息', message);
  var treeInfo = []
  var sum = 0
  setXToken(message.xtoken)
  source_data = message
  message.control_list.forEach((e, index) => {
    if (e.regionType == 'struct') {
      treeInfo.push({
        regionType: e.regionType,
        struct_name: e.name,
        struct_id: e.struct_id,
        score: 0,
        question_list: []
      })
    } else if (e.regionType == 'question') {
      let asks = []
      if (e.ask_controls && e.ask_controls.length > 0) {
        asks = e.ask_controls.map(ask => {
          return {
            is_ask: true,
            control_id: ask.control_id,
            v: e.ask_controls.length == 1 ? e.score : ask.v
          }
        })
      } else {
        asks.push({
          is_ask: false,
          control_id: e.control_id,
          v: e.score
        })
      }

      treeInfo[treeInfo.length - 1].question_list.push({
        uuid: e.ques_uuid,
        control_id: e.control_id,
        ques_no: e.ques_no,
        ques_name: e.ques_name,
        score: e.score,
        asks: asks,
        control_index: index
      })
      sum += e.score * 1
      treeInfo[treeInfo.length - 1].score += e.score * 1
    }
  })
  sumScore = sum
  tree = treeInfo
  console.log('treeInfo', treeInfo)
  init()
});
})(window, undefined);

