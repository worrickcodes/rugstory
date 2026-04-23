import { LitElement, css, html } from 'lit'
import { customElement, state } from 'lit/decorators.js'
import { PDFDocument } from 'pdf-lib'
import dynamicFields from './rugstoredynamicinfo.json'

export interface TextField {
  fieldName: string
  value: string
}

const TEMPLATE_URL = '/rugstoretemplate.pdf'

@customElement('rug-story')
export class RugStory extends LitElement {
  private fields: TextField[] = dynamicFields as TextField[]

  @state()
  private _pdfObjectUrl = ''

  override connectedCallback() {
    super.connectedCallback()
    this._buildPdf()
  }

  override disconnectedCallback() {
    super.disconnectedCallback()
    this._revoke()
  }

  private _revoke() {
    if (this._pdfObjectUrl) {
      URL.revokeObjectURL(this._pdfObjectUrl)
      this._pdfObjectUrl = ''
    }
  }

  private async _buildPdf() {
    this._revoke()

    try {
      const templateResponse = await fetch(TEMPLATE_URL)
      const templateBytes = await templateResponse.arrayBuffer()
      const pdfDoc = await PDFDocument.load(templateBytes)
      const form = pdfDoc.getForm()

      for (const field of this.fields) {
        const textField = form.getTextField(field.fieldName)
        textField.setText(field.value)
      }

      const finalBytes = await pdfDoc.save()
      const blob = new Blob([finalBytes.buffer as ArrayBuffer], { type: 'application/pdf' })
      this._pdfObjectUrl = URL.createObjectURL(blob)
    } catch (err) {
      console.error('PDF build failed:', err)
    }
  }

  render() {
    if (!this._pdfObjectUrl) return html``
    return html`<iframe .src=${this._pdfObjectUrl}></iframe>`
  }

  static styles = css`
    :host {
      display: block;
      height: 100vh;
      width: 100vw;
    }

    iframe {
      width: 100%;
      height: 100%;
      border: none;
    }
  `
}

declare global {
  interface HTMLElementTagNameMap {
    'rug-story': RugStory
  }
}
