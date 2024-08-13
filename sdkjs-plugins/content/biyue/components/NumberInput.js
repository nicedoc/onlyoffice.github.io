import { closeOtherSelect } from '../scripts/model/util.js'
class NumberInput {
	constructor(id, options = {}) {
		this.id = id
		this.options = options
		this.width = options.width || '85px'
		this.render()
	}

	removeEvent() {
		var btnup = $(`#${this.id} .spinner-up`)
		if (btnup) {
			btnup.off('click')
		}
		var btndown = $(`#${this.id} .spinner-down`)
		if (btndown) {
			btndown.off('click')
		}
		var input = $(`#${this.id} input`)
		if (input) {
			input.off('input')
			input.off('focus')
		}
	}
	render() {
		this.removeEvent()
		var content = ''
		content = `
    <div id=${this.id} class="spinner" style="width: ${this.width};">
      <input type="text" class="form-control" spellcheck="false" data-hint="1" data-hint-direction="bottom" data-hint-offset="big">
      <div class="spinner-buttons">
        <button type="button" class="spinner-up">
          <i class="arrow"></i>
        </button>
        <button type="button" class="spinner-down">
          <i class="arrow"></i>
        </button>
      </div>
    </div>`
		$(`#${this.id}`).replaceWith(content)
		$(`#${this.id} .spinner-up`).on('click', () => {
			this.onInputBtn(1)
		})
		$(`#${this.id} .spinner-down`).on('click', () => {
			this.onInputBtn(-1)
		})
		$(`#${this.id} input`).on('input', () => {
			this.valueChange($(`#${this.id} input`).val())
		})
		$(`#${this.id} input`).on('focus', () => {
			closeOtherSelect()
			if (this.options.focus) {
				this.options.focus(this.id)
			}
		})
	}

	hide() {
		$(`#${this.id}`).hide()
	}

	show() {
		$(`#${this.id}`).show()
	}

	setValue(value) {
		$(`#${this.id} input`).val(value)
	}

	getValue() {
		return $(`#${this.id} input`).val()
	}

	onInputBtn(offset) {
		var inputEl = $(`#${this.id} input`)
		var v = inputEl.val()
		var vv = parseInt(v)
		if (isNaN(vv)) {
			vv = 1
		}
		var targetv = offset + vv
		if (this.options.min != undefined) {
			if (targetv < this.options.min) {
				targetv = this.options.min
			}
		}
		inputEl.val(targetv)
		this.valueChange(inputEl.val())
	}

	valueChange(value) {
		if (this.options.change) {
			this.options.change(this.id, value)
		}
	}
}
export default NumberInput
