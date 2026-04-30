import { LitElement, css, html } from "lit";
import { customElement, state } from "lit/decorators.js";
import { PDFDocument, PDFName, PDFArray, PDFPage, PDFForm, rgb } from "pdf-lib";
import type { PDFTextField, PDFFont } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
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
  maintainAspectRatio?: boolean;
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
  material: string;
  kg: string;
}

export interface DesignInfo {
  physicalWidth: number;
  physicalHeight: number;
  pixelWidth: number;
  pixelHeight: number;
  kpu: number;
  lpu: number;
  eachCutsCount: number[];
  eachKnotCount: number[];
  colorsUsed: string[];
  materials: string[];
  unit: string;
}

export interface MaterialsEntry {
  type: "materials";
  fieldName: string;
  designInfo?: DesignInfo;
  materials?: Material[];
}

export type FieldEntry = TextEntry | ImageEntry | Weavemasters | MaterialsEntry;

const COLOR_BANK_URL = "./ColorBank.json";
const TEMPLATE_URL = "./rugstorytemplate2.pdf";
const FONT_URL = "./fonts/fonts/fonnts.com-AcuminPro-Regular.ttf";
const FONT_LIGHT_URL = "./fonts/fonts/fonnts.com-AcuminPro-Light.ttf";
const DARK_COLOR = rgb(46 / 255, 46 / 255, 45 / 255);
const DARK_DA = `${(46 / 255).toFixed(4)} ${(46 / 255).toFixed(4)} ${(45 / 255).toFixed(4)} rg`;

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
    return html`<iframe .src=${this._pdfObjectUrl + `#toolbar=0&view=Fit100%`}></iframe>`;
  }

  private async generateRugStory() {
    this._revoke();

    try {
      // Prefetch all resources in parallel
      const fetchUrls: string[] = [TEMPLATE_URL, FONT_URL, FONT_LIGHT_URL];
      for (const entry of this.fields) {
        if (entry.type === "image") fetchUrls.push(entry.src);
        if (entry.type === "profiles") {
          fetchUrls.push(entry.templateUrl);
          entry.profiles.forEach((p) => fetchUrls.push(p.src));
        }
      }
      const fetched = new Map<string, ArrayBuffer>();
      await Promise.all(
        fetchUrls.map((url) =>
          fetch(url)
            .then((r) => r.arrayBuffer())
            .then((buf) => fetched.set(url, buf)),
        ),
      );

      const colorBankRes = await fetch(COLOR_BANK_URL);
      const colorBankData = await colorBankRes.json() as { ColorRows: { R: number; G: number; B: number; ColorName: string }[] }[];
      const colorBank = new Map<string, string>();
      for (const group of colorBankData) {
        for (const row of group.ColorRows) {
          colorBank.set(`${row.R},${row.G},${row.B}`, row.ColorName);
        }
      }

      const templateBytes = fetched.get(TEMPLATE_URL)!;
      const pdfDoc = await PDFDocument.load(templateBytes);
      pdfDoc.registerFontkit(fontkit);
      const font = await pdfDoc.embedFont(fetched.get(FONT_URL)!);
      const form = pdfDoc.getForm();
      const pages = pdfDoc.getPages();

      // Auto-fill calculated fields from designInfo
      const materialsEntry = this.fields.find((e): e is MaterialsEntry => e.type === "materials" && !!e.designInfo);
      if (materialsEntry?.designInfo) {
        const di = materialsEntry.designInfo;
        const totalKnots = di.eachKnotCount.reduce((sum, k) => sum + k, 0);
        const sqm = (di.physicalWidth / 100) * (di.physicalHeight / 100);
        const weightPerSqm = 4;
        const totalConsumption = sqm * weightPerSqm;

        this._fillTextField(form, { type: "text", fieldName: "p22knots", value: totalKnots.toLocaleString() }, font, 14);
        this._fillTextField(form, { type: "text", fieldName: "p22weight", value: `${totalConsumption.toFixed(2)} Kgs` }, font, 14);
        this._fillTextField(form, { type: "text", fieldName: "Text3", value: `${Math.floor(di.physicalWidth)} x ${Math.floor(di.physicalHeight)} ${di.unit}` }, font, 14);
      }

      for (const entry of this.fields) {
        try {
          if (entry.type === "text") {
            this._fillTextField(form, entry, font);
          } else if (entry.type === "image") {
            await this._fillImageField(pdfDoc, form, pages, entry, fetched.get(entry.src)!);
          } else if (entry.type === "profiles") {
            await this._fillProfilesField(pdfDoc, form, pages, entry, fetched);
          } else if (entry.type === "materials") {
            await this._fillMaterialsField(pdfDoc, form, pages, entry, font, colorBank);
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
  private _fillTextField(form: PDFForm, entry: TextEntry, font: PDFFont, fontSize = 0) {
    const field = form.getTextField(entry.fieldName);
    field.acroField.setDefaultAppearance(`${DARK_DA} /Helv ${fontSize} Tf`);
    field.setText(entry.value);
    field.updateAppearances(font);
  }

  private async _fillImageField(pdfDoc: PDFDocument, form: PDFForm, pages: PDFPage[], entry: ImageEntry, imgBytes: ArrayBuffer) {
    const { field, rect, page } = this._resolveField(form, pages, entry.fieldName);

    const embedded = await this._embedImage(pdfDoc, new Uint8Array(imgBytes));
    this._removeField(pdfDoc, form, field, page);

    const dims = entry.maintainAspectRatio
      ? this._fitImageInRect(embedded.width, embedded.height, rect)
      : { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    page.drawImage(embedded, dims);

    const contentsKey = PDFName.of("Contents");
    const contents = page.node.get(contentsKey);
    if (contents instanceof PDFArray && contents.size() > 1) {
      const lastIdx = contents.size() - 1;
      const imageRef = contents.get(lastIdx);
      contents.remove(lastIdx);
      contents.insert(0, imageRef);
    }
  }

  private async _fillProfilesField(pdfDoc: PDFDocument, form: PDFForm, pages: PDFPage[], entry: Weavemasters, fetched: Map<string, ArrayBuffer>) {
    const { field, rect, page } = this._resolveField(form, pages, entry.fieldName);
    this._removeField(pdfDoc, form, field, page);

    const weaverBytes = fetched.get(entry.templateUrl)!;
    const imageDataList = entry.profiles.map((p) => new Uint8Array(fetched.get(p.src)!));

    const count = entry.profiles.length;
    const cardBounds = { x: 50, y: 420, width: 325, height: 80 };
    const sampleDoc = await PDFDocument.load(weaverBytes);
    const wpWidth = sampleDoc.getPage(0).getWidth();
    const wpHeight = sampleDoc.getPage(0).getHeight();
    const cols = count > 4 ? 2 : 1;
    const rows = Math.ceil(count / cols);

    // Divide the allocated rect into a grid of cells with gap between rows/cols
    const rowGap = 6;
    const colGap = cols > 1 ? 18 : 0;
    const cellWidth = (rect.width - (cols - 1) * colGap) / cols;
    const cellHeight = (rect.height - (rows - 1) * rowGap) / rows;

    // Scale uniformly based on width so card fills the allocated width
    const scale = cellWidth / cardBounds.width;
    const compensation = cols > 1 ? rect.width / cellWidth : 1;
    const scaledCardH = cardBounds.height * scale;

    // Generate all weavemaster pages in parallel
    const filledPages = await Promise.all(
      entry.profiles.map((p, i) => this.generateWeavemasterPage(weaverBytes, p, imageDataList[i], fetched.get(FONT_URL)!, fetched.get(FONT_LIGHT_URL)!, compensation)),
    );

    for (let i = 0; i < filledPages.length; i++) {
      const [embeddedPage] = await pdfDoc.embedPdf(filledPages[i], [0]);

      const col = i % cols;
      const row = Math.floor(i / cols);

      const cellX = rect.x + col * (cellWidth + colGap);
      const cellY = rect.y + rect.height - (row + 1) * cellHeight;
      const offsetY = (cellHeight - scaledCardH) / 2;

      page.drawPage(embeddedPage, {
        x: cellX - cardBounds.x * scale,
        y: cellY + offsetY - cardBounds.y * scale,
        width: wpWidth * scale,
        height: wpHeight * scale,
      });
    }
  }

  private async _fillMaterialsField(pdfDoc: PDFDocument, form: PDFForm, pages: PDFPage[], entry: MaterialsEntry, font: PDFFont, colorBank: Map<string, string>) {
    const { field, rect, page } = this._resolveField(form, pages, entry.fieldName);
    this._removeField(pdfDoc, form, field, page);

    const materials = entry.materials ?? this._calculateMaterials(entry.designInfo!, colorBank);

    // Find the largest swatch size that fits, starting from default
    const defaultSwatch = 57;
    const colGap = 8;
    const rowGap = 10;
    const count = materials.length;

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

    const scale = swatchSize / defaultSwatch;
    const kgFontSize = Math.ceil(8 * scale);
    const nameFontSize = Math.ceil(7 * scale);
    const padding = 3 * scale;
    const kgRightPadding = padding * 1.75;
    const kgTopOffset = 1;
    const brightnessThreshold = 0.8;
    const maxNameLines = 2;
    const nameMaxWidth = swatchSize - padding - 1;

    const wrapText = (text: string, fontSize: number) => {
      const lines: string[] = [];
      const words = text.split(" ");
      let cur = words[0];
      for (let w = 1; w < words.length; w++) {
        const test = cur + " " + words[w];
        if (font.widthOfTextAtSize(test, fontSize) <= nameMaxWidth) {
          cur = test;
        } else {
          lines.push(cur);
          cur = words[w];
        }
      }
      lines.push(cur);
      return lines;
    };

    for (let i = 0; i < materials.length; i++) {
      const mat = materials[i];
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = rect.x + col * (swatchSize + colGap);
      const y = rect.y + rect.height - (row + 1) * swatchSize - row * rowGap;
      const { r, g, b } = this._parseColor(mat.color);

      page.drawRectangle({ x, y, width: swatchSize, height: swatchSize, color: rgb(r, g, b) });

      const brightness = r * 0.299 + g * 0.587 + b * 0.114;
      const textColor = brightness > brightnessThreshold ? DARK_COLOR : rgb(1, 1, 1);

      // Kg text — top right
      const kgText = `${mat.kg} Kg`;
      const kgTextWidth = font.widthOfTextAtSize(kgText, kgFontSize);
      page.drawText(kgText, {
        x: x + swatchSize - kgRightPadding - kgTextWidth,
        y: y + swatchSize - kgTopOffset - kgFontSize,
        size: kgFontSize,
        font,
        color: textColor,
      });

      // Material type — bottom left
      page.drawText(mat.material, {
        x: x + padding,
        y: y + padding,
        size: nameFontSize,
        font,
        color: textColor,
      });

      // Color name — above material, shrinks only if >2 lines
      const nameBottomY = y + padding + nameFontSize;
      let nfs = nameFontSize;
      let nameLines = wrapText(mat.name, nfs);
      while (nameLines.length > maxNameLines && nfs >= 4 * scale) {
        nfs -= 0.5;
        nameLines = wrapText(mat.name, nfs);
      }
      const nlh = nfs + 2 * scale;

      for (let l = 0; l < nameLines.length; l++) {
        page.drawText(nameLines[l], {
          x: x + padding,
          y: nameBottomY + (nameLines.length - 1 - l) * nlh,
          size: nfs,
          font,
          color: textColor,
        });
      }
    }
  }

  private _calculateMaterials(designInfo: DesignInfo, colorBank: Map<string, string>): Material[] {
    const totalKnots = designInfo.eachKnotCount.reduce((s, k) => s + k, 0);
    const totalCuts = designInfo.eachCutsCount.reduce((s, c) => s + c, 0);
    const sqm = (designInfo.physicalWidth / 100) * (designInfo.physicalHeight / 100);
    const weightPerSqm = 4; //default
    const totalConsumption = sqm * weightPerSqm;
    const wastagePerSqm = weightPerSqm * (totalCuts / totalKnots);
    const totalCutWastage = sqm * wastagePerSqm;

    return designInfo.colorsUsed.map((csv, i) => {
      const [rv, gv, bv] = csv.split(",").map((v) => parseInt(v.trim()));
      const key = `${rv},${gv},${bv}`;
      const rp = Math.trunc((100 * rv) / 256);
      const gp = Math.trunc((100 * gv) / 256);
      const bp = Math.trunc((100 * bv) / 256);
      const name = colorBank.get(key) ?? `R${rp.toString().padStart(2, "0")} G${gp.toString().padStart(2, "0")} B${bp.toString().padStart(2, "0")}`;
      const materialType = designInfo.materials[i].match(/^[A-Z][a-z]*/)?.[0] ?? "Wool";

      const consumption = totalConsumption * (designInfo.eachKnotCount[i] / totalKnots);
      const wastage = totalCutWastage * (designInfo.eachCutsCount[i] / totalCuts);
      const kg = (consumption + wastage).toFixed(2);

      return {
        color: `#${rv.toString(16).padStart(2, "0")}${gv.toString(16).padStart(2, "0")}${bv.toString(16).padStart(2, "0")}`,
        name,
        material: materialType,
        kg,
      };
    });
  }

  private async generateWeavemasterPage(templateBytes: ArrayBuffer, profile: Profile, imgBytes: Uint8Array, fontBytes: ArrayBuffer, thinFontBytes: ArrayBuffer, compensation = 1) {
    const doc = await PDFDocument.load(templateBytes);
    doc.registerFontkit(fontkit);
    const wmFont = await doc.embedFont(fontBytes);
    const wmThinFont = await doc.embedFont(thinFontBytes);
    const form = doc.getForm();

    const pages = doc.getPages();

    // Name + role: draw manually so role can be smaller
    const nameField = form.getTextField("Name");
    const nameWidgets = nameField.acroField.getWidgets();
    const nameRect = nameWidgets[0].getRectangle();
    const namePage = this._getWidgetPage(nameWidgets[0], pages);
    this._removeField(doc, form, nameField, namePage);

    const photoField = form.getTextField("Photo");
    const photoWidgets = photoField.acroField.getWidgets();
    const photoRect = photoWidgets.length ? photoWidgets[0].getRectangle() : null;
    const photoPage = photoWidgets.length ? this._getWidgetPage(photoWidgets[0], pages) : namePage;

    const isMultiCol = compensation > 1 && photoRect;
    const imgScale = isMultiCol ? 1.4 : 1;

    // Compute scaled photo rect (used for both photo drawing and text positioning)
    const scaledPhotoRect = photoRect
      ? {
          x: photoRect.x - (photoRect.width * (imgScale - 1)) / 2,
          y: photoRect.y - (photoRect.height * (imgScale - 1)) / 2,
          width: photoRect.width * imgScale,
          height: photoRect.height * imgScale,
        }
      : null;

    // ── Photo
    if (scaledPhotoRect) {
      const embedded = await this._embedImage(doc, imgBytes);
      const dims = this._fitImageInRect(embedded.width, embedded.height, scaledPhotoRect);
      this._removeField(doc, form, photoField, photoPage);
      photoPage.drawImage(embedded, dims);
    }

    // ── Name + Role
    const nameFontSize = isMultiCol ? 20 : 11;
    const roleFontSize = isMultiCol ? 13 : 7;

    let textX: number;
    let nameBaselineY: number;

    if (isMultiCol && scaledPhotoRect) {
      textX = scaledPhotoRect.x + scaledPhotoRect.width + 15;
      const textTopY = scaledPhotoRect.y + scaledPhotoRect.height + 6;
      nameBaselineY = textTopY - nameFontSize;
    } else {
      textX = nameRect.x;
      nameBaselineY = nameRect.y + (nameRect.height - nameFontSize) / 2;
    }

    const nameWidth = wmFont.widthOfTextAtSize(profile.name + " ", nameFontSize);
    namePage.drawText(profile.name, {
      x: textX,
      y: nameBaselineY,
      size: nameFontSize,
      font: wmFont,
      color: DARK_COLOR,
    });
    namePage.drawText(`(${profile.role})`, {
      x: textX + nameWidth,
      y: nameBaselineY,
      size: roleFontSize,
      font: wmThinFont,
      color: rgb(0, 0, 0),
    });

    // ── Description
    const descFontSize = isMultiCol ? 10 : 7;
    const descField = form.getTextField("Description");
    if (isMultiCol) {
      const descWidgets = descField.acroField.getWidgets();
      const descRect = descWidgets[0].getRectangle();
      const descTopY = nameBaselineY - 8;
      const descHeight = descRect.height * compensation;
      descWidgets[0].setRectangle({
        x: textX,
        y: descTopY - descHeight,
        width: descRect.width,
        height: descHeight,
      });
    }
    descField.acroField.setDefaultAppearance(`${DARK_DA} /Helv ${descFontSize} Tf`);
    descField.setFontSize(descFontSize);
    descField.enableMultiline();
    descField.setText(profile.description);
    descField.updateAppearances(wmFont);

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

  private async _embedImage(doc: PDFDocument, bytes: Uint8Array) {
    const isJpg = bytes[0] === 0xff && bytes[1] === 0xd8;
    const isPng = bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;

    if (isJpg) return doc.embedJpg(bytes);
    if (isPng) return doc.embedPng(bytes);

    // Convert any other browser-supported format (WebP, GIF, BMP, AVIF, etc.) to PNG via canvas
    const blob = new Blob([bytes.buffer as ArrayBuffer]);
    const url = URL.createObjectURL(blob);
    try {
      const img = await new Promise<HTMLImageElement>((resolve, reject) => {
        const el = new Image();
        el.onload = () => resolve(el);
        el.onerror = () => reject(new Error("Unsupported image format"));
        el.src = url;
      });
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      canvas.getContext("2d")!.drawImage(img, 0, 0);
      const pngBlob = await new Promise<Blob>((resolve, reject) => {
        canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("Canvas conversion failed"))), "image/png");
      });
      const pngBytes = new Uint8Array(await pngBlob.arrayBuffer());
      return doc.embedPng(pngBytes);
    } finally {
      URL.revokeObjectURL(url);
    }
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