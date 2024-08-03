class ComponentSelect {
  constructor(params = {}) {
    this.id = params.id
    this.callback_item = params.callback_item
	if (params.force_click_notify == undefined) {
		this.force_click_notify = true
	} else {
		this.force_click_notify = params.force_click_notify
	}
    this.updateOptions(params)
  }

  updateOptions(params = {}) {
    this.options = params.options
    this.value_select = params.value_select
    this.width = params.width || '85px'
    this.render()
  }

  render() {
    $(`#${this.id}`).empty()
    var content = ''
    var strOptions = ''
    this.options.forEach(item => {
      strOptions += `<li id="${this.id}_${item.value}" data-value=${item.value} class="${item.value==this.value_select?'selected':''}"><a tabindex="-1" type="menuitem">${item.label}</a></li>`
    })
    content = `
      <span id="${this.id}_span" class="input-group combobox input-group-nr">
        <input id="${this.id}_input" type="text" class="form-control" spellcheck="false" placeholder="" data-hint="1" data-hint-direction="bottom" data-hint-offset="big" readonly="readonly" data-can-copy="false">
        <button id="${this.id}_button" type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown" aria-expanded="true">
          <span class="caret"></span>
        </button>
        <ul id="${this.id}_ul" class="dropdown-menu ps-container oo" style="min-width: ${this.width}; max-height: 774px;" role="menu">${strOptions}</ul>
      </span>
    `
	$(`#${this.id}`).css('width', `${this.width}`)
    $(`#${this.id}`).html(content)

    $(`#${this.id}_span`).on('click', () => {
      this.toggleOptions()
    })
    this.options.forEach(item => {
      $(`#${this.id}_${item.value}`).on('click', (e) => {
        this.clickOption(e, item)
      })
    })
    var data = this.options.find(item => item.value == this.value_select)
    if (data) {
      $(`#${this.id}_input`).val(data.label)
    }
  }

  toggleOptions() {
    if ($(`#${this.id}_span`).hasClass('open')) {
      $(`#${this.id}_span`).removeClass('open')
    } else {
      $(`#${this.id}_span`).addClass('open')
    }
  }

  clickOption(e, data) {
    if (e) {
      e.cancelBubble = true
      e.preventDefault()
      // 检查点击的元素是否是下拉菜单中的选项
      if ($(e.target).closest(`#${this.id}_ul`).length > 0) {
      } else {
        // 如果不是选项，则执行关闭下拉菜单的逻辑
        $(`#qrCode_span_span`).removeClass('open');
      }
    }
    if (this.value_select == data.value) {
		if (this.force_click_notify && this.callback_item) {
			this.callback_item(data)
		}
      return
    }
    $(`#${this.id}_${this.value_select}`).removeClass('selected')
    $(`#${this.id}_${data.value}`).addClass('selected')
    this.value_select = data.value
    $(`#${this.id}_input`).val(data.label)
    if (this.callback_item) {
      this.callback_item(data)
    }
  }

  setSelect(value) {
    if (this.value_select == value) {
      return
    }
    $(`#${this.id}_${this.value_select}`).removeClass('selected')
    $(`#${this.id}_${value}`).addClass('selected')
    this.value_select = value
    var data = this.options.find(item => item.value == value)
    $(`#${this.id}_span`).removeClass('open')
	if (data) {
		$(`#${this.id}_input`).val(data.label)
	}
  }

  getValue() {
    return this.value_select
  }
}

export default ComponentSelect