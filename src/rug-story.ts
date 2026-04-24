import { LitElement, css, html } from "lit";
import { customElement, state } from "lit/decorators.js";
import { PDFDocument, PDFName, PDFArray, PDFPage, PDFForm } from "pdf-lib";
import type { PDFTextField } from "pdf-lib";
import dynamicFields from "./rugstorydynamicinfo.json";

export interface TextEntry {
  type: "text";
  fieldName: string;
  value: string;
}

export interface ImageEntry {
  type: "image";
  fieldName: string;
  src: string;
}

export interface Profile {
  name: string;
  role: string;
  description: string;
  src: string;
}

export interface Weavemasters {
  type: "profiles";
  fieldName: string;
  templateUrl: string;
  profiles: Profile[];
}

export type FieldEntry = TextEntry | ImageEntry | Weavemasters;

const TEMPLATE_URL = "./rugstorytemplate1.pdf";

@customElement("rug-story")
export class RugStory extends LitElement {
  private fields: FieldEntry[] = dynamicFields as FieldEntry[];

  @state()
  private _pdfObjectUrl = "";

  override connectedCallback() {
    super.connectedCallback();
    this.generateRugStory();
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    this._revoke();
  }

  render() {
    if (!this._pdfObjectUrl) return html``;
    return html`<iframe .src=${this._pdfObjectUrl}></iframe>`;
  }

  private async generateRugStory() {
    this._revoke();

    try {
      const templateBytes = await fetch(TEMPLATE_URL).then((r) => r.arrayBuffer());
      const pdfDoc = await PDFDocument.load(templateBytes);
      const form = pdfDoc.getForm();
      const pages = pdfDoc.getPages();

      for (const entry of this.fields) {
        try {
          if (entry.type === "text") {
            this._fillTextField(form, entry);
          } else if (entry.type === "image") {
            await this._fillImageField(pdfDoc, form, pages, entry);
          } else if (entry.type === "profiles") {
            await this._fillProfilesField(pdfDoc, form, pages, entry);
          }
        } catch (fieldErr) {
          console.error(`Failed to process field "${(entry as { fieldName: string }).fieldName}":`, fieldErr);
        }
      }

      const finalBytes = await pdfDoc.save();
      const blob = new Blob([finalBytes.buffer as ArrayBuffer], { type: "application/pdf" });
      this._pdfObjectUrl = URL.createObjectURL(blob);
    } catch (err) {
      console.error("PDF build failed:", err);
    }
  }

  // ── Field handlers
  private _fillTextField(form: PDFForm, entry: TextEntry) {
    form.getTextField(entry.fieldName).setText(entry.value);
  }

  private async _fillImageField(pdfDoc: PDFDocument, form: PDFForm, pages: PDFPage[], entry: ImageEntry) {
    const { field, rect, page } = this._resolveField(form, pages, entry.fieldName);

    const embedded = await this._embedImage(pdfDoc, entry.src);
    this._removeField(pdfDoc, form, field, page);

    page.drawImage(embedded, { x: rect.x, y: rect.y, width: rect.width, height: rect.height });

    const contentsKey = PDFName.of("Contents");
    const contents = page.node.get(contentsKey);
    if (contents instanceof PDFArray && contents.size() > 1) {
      const lastIdx = contents.size() - 1;
      const imageRef = contents.get(lastIdx);
      contents.remove(lastIdx);
      contents.insert(0, imageRef);
    }
  }

  private async _fillProfilesField(pdfDoc: PDFDocument, form: PDFForm, pages: PDFPage[], entry: Weavemasters) {
    const { field, rect, page } = this._resolveField(form, pages, entry.fieldName);
    this._removeField(pdfDoc, form, field, page);

    const weaverBytes = await fetch(entry.templateUrl).then((r) => r.arrayBuffer());
    const imageDataList = await Promise.all(
      entry.profiles.map((p) => fetch(p.src).then((r) => r.arrayBuffer()).then((b) => new Uint8Array(b))),
    );

    const cardBounds = { x: 50, y: 420, width: 325, height: 80 };
    const sampleDoc = await PDFDocument.load(weaverBytes);
    const wpWidth = sampleDoc.getPage(0).getWidth();
    const wpHeight = sampleDoc.getPage(0).getHeight();
    const scale = rect.width / cardBounds.width;
    const scaledCardH = cardBounds.height * scale;
    const gap = 4;

    for (let i = 0; i < entry.profiles.length; i++) {
      const filledPage = await this.generateWeavemasterPage(weaverBytes, entry.profiles[i], imageDataList[i]);

      const [embeddedPage] = await pdfDoc.embedPdf(filledPage, [0]);

      const cardY = rect.y + rect.height - (i + 1) * (scaledCardH + gap) + gap;
      page.drawPage(embeddedPage, {
        x: rect.x - cardBounds.x * scale,
        y: cardY - cardBounds.y * scale,
        width: wpWidth * scale,
        height: wpHeight * scale,
      });
    }
  }

  private async generateWeavemasterPage(templateBytes: ArrayBuffer, profile: Profile, imgBytes: Uint8Array) {
    const doc = await PDFDocument.load(templateBytes);
    const form = doc.getForm();

    form.getTextField("Name").setText(`${profile.name} (${profile.role})`);
    const descField = form.getTextField("Description");
    descField.enableMultiline();
    descField.setText(profile.description);

    const photoField = form.getTextField("Photo");
    const photoWidgets = photoField.acroField.getWidgets();
    if (photoWidgets.length) {
      const photoRect = photoWidgets[0].getRectangle();
      const pages = doc.getPages();
      const photoPage = this._getWidgetPage(photoWidgets[0], pages);

      const isJpg = imgBytes[0] === 0xff && imgBytes[1] === 0xd8;
      const embedded = isJpg ? await doc.embedJpg(imgBytes) : await doc.embedPng(imgBytes);

      const dims = this._fitImageInRect(embedded.width, embedded.height, photoRect);
      this._removeField(doc, form, photoField, photoPage);
      photoPage.drawImage(embedded, dims);
    }

    form.flatten();
    const savedBytes = await doc.save();
    return PDFDocument.load(savedBytes);
  }

  // ── PDF utilities
  private _resolveField(form: PDFForm, pages: PDFPage[], fieldName: string) {
    const field = form.getTextField(fieldName);
    const widgets = field.acroField.getWidgets();
    if (!widgets.length) throw new Error(`Field "${fieldName}" has no widgets`);

    const widget = widgets[0];
    return {
      field,
      rect: widget.getRectangle(),
      page: this._getWidgetPage(widget, pages),
    };
  }

  private async _embedImage(pdfDoc: PDFDocument, src: string) {
    const response = await fetch(src);
    const bytes = new Uint8Array(await response.arrayBuffer());
    const isJpg = bytes[0] === 0xff && bytes[1] === 0xd8;
    return isJpg ? pdfDoc.embedJpg(bytes) : pdfDoc.embedPng(bytes);
  }

  private _getWidgetPage(widget: ReturnType<PDFTextField["acroField"]["getWidgets"]>[0], pages: PDFPage[]) {
    const pageRef = widget.P();
    return pageRef ? (pages.find((p) => p.ref === pageRef) ?? pages[0]) : pages[0];
  }

  private _removeField(pdfDoc: PDFDocument, form: PDFForm, field: PDFTextField, page: PDFPage) {
    field.acroField.getWidgets().forEach((w) => {
      const ref = pdfDoc.context.getObjectRef(w.dict);
      if (ref) page.node.removeAnnot(ref);
    });
    form.acroForm.removeField(field.acroField);
  }

  private _fitImageInRect(imgWidth: number, imgHeight: number, rect: { x: number; y: number; width: number; height: number }) {
    const imgAspect = imgWidth / imgHeight;
    const rectAspect = rect.width / rect.height;
    const [drawW, drawH] =
      imgAspect > rectAspect
        ? [rect.width, rect.width / imgAspect]
        : [rect.height * imgAspect, rect.height];
    return {
      x: rect.x + (rect.width - drawW) / 2,
      y: rect.y + (rect.height - drawH) / 2,
      width: drawW,
      height: drawH,
    };
  }

  private _revoke() {
    if (this._pdfObjectUrl) {
      URL.revokeObjectURL(this._pdfObjectUrl);
      this._pdfObjectUrl = "";
    }
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
  `;
}

declare global {
  interface HTMLElementTagNameMap {
    "rug-story": RugStory;
  }
}
