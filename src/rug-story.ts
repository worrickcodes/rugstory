import { LitElement, css, html } from "lit";
import { customElement, state } from "lit/decorators.js";
import { PDFDocument, PDFName, PDFArray, PDFPage, PDFForm, StandardFonts, rgb } from "pdf-lib";
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

export interface Material {
  color: string;
  name: string;
  kg: string;
}

export interface MaterialsEntry {
  type: "materials";
  fieldName: string;
  materials: Material[];
}

export type FieldEntry = TextEntry | ImageEntry | Weavemasters | MaterialsEntry;
const TEMPLATE_URL = "./rugstorytemplate1.pdf";

@customElement("rug-story")
export class RugStory extends LitElement {
  private fields: FieldEntry[] = dynamicFields as FieldEntry[];

  @state()
  private _pdfObjectUrl = "";

  connectedCallback() {
    super.connectedCallback();
    this.generateRugStory();
  }

  disconnectedCallback() {
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
          } else if (entry.type === "materials") {
            await this._fillMaterialsField(pdfDoc, form, pages, entry);
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

  private async _fillMaterialsField(pdfDoc: PDFDocument, form: PDFForm, pages: PDFPage[], entry: MaterialsEntry) {
    const { field, rect, page } = this._resolveField(form, pages, entry.fieldName);
    this._removeField(pdfDoc, form, field, page);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    // Find the largest swatch size that fits, starting from default
    const defaultSwatch = 57;
    const colGap = 8;
    const rowGap = 10;
    const count = entry.materials.length;

    let swatchSize = defaultSwatch;
    let cols = 1;

    // Try default size first; if it doesn't fit, find the largest that does
    for (let c = count; c >= 1; c--) {
      const rows = Math.ceil(count / c);
      const fitW = (rect.width - (c - 1) * colGap) / c;
      const fitH = (rect.height - (rows - 1) * rowGap) / rows;
      const size = Math.min(fitW, fitH);

      if (size >= defaultSwatch) {
        swatchSize = defaultSwatch;
        cols = c;
        break;
      }
      if (size > swatchSize || swatchSize === defaultSwatch) {
        swatchSize = size;
        cols = c;
      }
    }

    // Scale text proportionally to swatch size
    const scale = swatchSize / defaultSwatch;
    const kgFontSize = 9 * scale;
    const nameFontSize = 8 * scale;
    const padding = 5.5 * scale;

    for (let i = 0; i < entry.materials.length; i++) {
      const mat = entry.materials[i];
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = rect.x + col * (swatchSize + colGap);
      const y = rect.y + rect.height - (row + 1) * swatchSize - row * rowGap;
      const { r, g, b } = this._parseColor(mat.color);

      page.drawRectangle({ x, y, width: swatchSize, height: swatchSize, color: rgb(r, g, b) });

      const brightness = r * 0.299 + g * 0.587 + b * 0.114;
      const textColor = brightness > 0.5 ? rgb(0, 0, 0) : rgb(1, 1, 1);

      const kgText = `${mat.kg} Kg`;
      const kgTextWidth = font.widthOfTextAtSize(kgText, kgFontSize);
      page.drawText(kgText, {
        x: x + swatchSize - padding - kgTextWidth,
        y: y + swatchSize - padding - kgFontSize,
        size: kgFontSize,
        font,
        color: textColor,
      });

      // Available height for name: from bottom padding to just below kg text
      const nameAreaH = swatchSize * 0.5 - padding;
      let nfs = nameFontSize;
      let nlh = nfs + 2 * scale;
      let nameLines: string[] = [];

      // Shrink font until all wrapped lines fit in the available area
      while (nfs >= 4 * scale) {
        nameLines = [];
        const words = mat.name.split(" ");
        let cur = words[0];
        for (let w = 1; w < words.length; w++) {
          const test = cur + " " + words[w];
          if (font.widthOfTextAtSize(test, nfs) <= swatchSize - padding * 2) {
            cur = test;
          } else {
            nameLines.push(cur);
            cur = words[w];
          }
        }
        nameLines.push(cur);
        nlh = nfs + 2 * scale;
        if (nameLines.length * nlh <= nameAreaH) break;
        nfs -= 0.5;
      }

      for (let l = 0; l < nameLines.length; l++) {
        page.drawText(nameLines[l], {
          x: x + padding,
          y: y + padding + (nameLines.length - 1 - l) * nlh,
          size: nfs,
          font,
          color: textColor,
        });
      }
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

  private _parseColor(color: string): { r: number; g: number; b: number } {
    const rgbMatch = color.match(/R(\d+)\s*G(\d+)\s*B(\d+)/i);
    if (rgbMatch) {
      return {
        r: parseInt(rgbMatch[1], 10) / 255,
        g: parseInt(rgbMatch[2], 10) / 255,
        b: parseInt(rgbMatch[3], 10) / 255,
      };
    }
    const hex = color.replace("#", "");
    return {
      r: parseInt(hex.substring(0, 2), 16) / 255,
      g: parseInt(hex.substring(2, 4), 16) / 255,
      b: parseInt(hex.substring(4, 6), 16) / 255,
    };
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
