"use client";

import React, { useCallback, useEffect, useRef, useState } from "react";

function MaterialIcon({ name, className = "icon" }: { name: string; className?: string }) {
  return <span className={`material-symbols-rounded ${className}`} aria-hidden="true">{name}</span>;
}

const BACKEND_BASE = process.env.NEXT_PUBLIC_BACKEND_URL || "http://localhost:8001";
const API_BASE = `${BACKEND_BASE}/api`;
const STATIC_BASE = BACKEND_BASE;

type Tab = "xlsx" | "images" | "extract" | "format";
type HeaderMode = "supplier" | "rs" | "both";
type Stage = "idle" | "processing" | "done";

interface Img {
  id: string;
  filename: string;
  path?: string;
  mimeType: string;
  sizeBytes: number;
  url: string;
  copyLink?: string;
  downloadUrl?: string;
  createdAt?: string;
  updatedAt?: string;
}

interface Pdf {
  filename: string;
  status: string;
}

interface JobCounts {
  total: number;
  success: number;
  failed: number;
}

interface DuplicateSpecification {
  productName: string;
  sku: string;
  attribute: string;
  value: string;
  occurrences: number;
  sourceReferences: string[];
}

interface DuplicateValidation {
  duplicateProductCount: number;
  duplicateCount: number;
  duplicates: DuplicateSpecification[];
}

type CropRect = { x: number; y: number; width: number; height: number };
type CropInteractionMode = "move" | "n" | "s" | "e" | "w" | "nw" | "ne" | "sw" | "se";
type CropInteraction = { mode: CropInteractionMode; start: { x: number; y: number }; crop: CropRect };
type ExtractedTable = { section: string; headers: string[]; rows: string[][] };
type ExtractMode = "manual" | "bulk";
type BulkExtractResult = {
  id: string;
  sku: string;
  pdfName: string;
  status: "ready" | "failed";
  error?: string;
  table?: ExtractedTable;
  source?: { page?: number; pages?: number[]; method?: string; confidence?: number | null };
  tablesDetected?: number;
  excludedColumns: number[];
};
type FormatAnalysis = {
  sourceSheet: string;
  sourceRows: number;
  sourceColumns: number;
  outputRows: number;
  outputColumns: number;
  groups: string[];
  dynamicFields?: string[];
  missingFields: string[];
  preview: { headers: string[]; rows: string[][] };
};

const formatBytes = (bytes = 0) => {
  if (!bytes) return "1 KB";
  const kb = Math.max(1, Math.round(bytes / 1024));
  return kb >= 1024 ? `${(kb / 1024).toFixed(1)} MB` : `${kb} KB`;
};

const assetUrl = (url?: string) => {
  if (!url) return "";
  if (/^https?:\/\//i.test(url)) {
    try {
      const parsed = new URL(url);
      if ((parsed.hostname === "localhost" || parsed.hostname === "127.0.0.1") && parsed.pathname.startsWith("/static/")) {
        return `${STATIC_BASE}${parsed.pathname}${parsed.search}`;
      }
    } catch {
      return url;
    }
    return url;
  }
  return `${STATIC_BASE}${url.startsWith("/") ? url : `/${url}`}`;
};

const getApiError = async (response: Response, fallback: string) => {
  const payload = await response.json().catch(() => ({}));
  if (typeof payload?.message === "string" && payload.message.trim()) return payload.message;
  if (Array.isArray(payload?.message)) return payload.message.join(", ");
  return fallback;
};

const isRecord = (value: unknown): value is Record<string, unknown> => Boolean(value) && typeof value === "object" && !Array.isArray(value);

const tableCellText = (value: unknown) => {
  if (value === undefined || value === null) return "";
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
};

const normalizeExtractedTable = (value: unknown): ExtractedTable | null => {
  let payload = value;
  if (typeof payload === "string") {
    try {
      payload = JSON.parse(payload);
    } catch {
      return null;
    }
  }
  if (!isRecord(payload)) return null;

  const rawRows = Array.isArray(payload.rows)
    ? payload.rows
    : isRecord(payload.rows)
      ? Object.values(payload.rows)
      : [];
  const suppliedHeaders = Array.isArray(payload.headers) ? payload.headers.map(tableCellText) : [];
  const objectRows = rawRows.filter(isRecord);
  const keyValueRows = objectRows.length > 0 && objectRows.every(row =>
    "value" in row && ("label" in row || "property" in row || "key" in row),
  );
  const inferredHeaders = keyValueRows
    ? ["Property", "Value"]
    : objectRows.length
      ? Object.keys(objectRows[0])
      : [];
  const widestArrayRow = rawRows.reduce((width, row) => Array.isArray(row) ? Math.max(width, row.length) : width, 0);
  const headers = suppliedHeaders.length
    ? suppliedHeaders
    : inferredHeaders.length
      ? inferredHeaders
      : Array.from({ length: widestArrayRow }, (_, index) => `Column ${index + 1}`);
  const rows = rawRows
    .map(row => {
      if (Array.isArray(row)) return headers.map((_, index) => tableCellText(row[index]));
      if (isRecord(row)) {
        if (keyValueRows) return [tableCellText(row.label ?? row.property ?? row.key), tableCellText(row.value)];
        return headers.map(header => tableCellText(row[header]));
      }
      return headers.length === 1 ? [tableCellText(row)] : [];
    })
    .filter(row => row.some(cell => cell.trim()));

  return { section: tableCellText(payload.section).trim(), headers, rows };
};

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>("xlsx");
  const [drag, setDrag] = useState(false);
  const [stage, setStage] = useState<Stage>("idle");
  const [file, setFile] = useState<File | null>(null);
  const [headerMode, setHeaderMode] = useState<HeaderMode>("supplier");
  const [jobNotice, setJobNotice] = useState("");
  const [validatingDuplicates, setValidatingDuplicates] = useState(false);
  const [duplicateValidation, setDuplicateValidation] = useState<DuplicateValidation | null>(null);
  const [duplicateReviewOpen, setDuplicateReviewOpen] = useState(false);
  const [jobId, setJobId] = useState<string | null>(null);
  const [pct, setPct] = useState(0);
  const [pdfs, setPdfs] = useState<Pdf[]>([]);
  const [jobCounts, setJobCounts] = useState<JobCounts>({ total: 0, success: 0, failed: 0 });
  const [images, setImages] = useState<Img[]>([]);
  const [imageTotal, setImageTotal] = useState(0);
  const [imageSearch, setImageSearch] = useState("");
  const [galleryPage, setGalleryPage] = useState(1);
  const [galleryPageSize, setGalleryPageSize] = useState(24);
  const [galleryNotice, setGalleryNotice] = useState("");
  const [renaming, setRenaming] = useState<{ id: string; value: string } | null>(null);
  const [sync, setSync] = useState({ up: 0, total: 0, loading: false });
  const [previewing, setPreviewing] = useState(false);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [extractFile, setExtractFile] = useState<File | null>(null);
  const [extractMode, setExtractMode] = useState<ExtractMode>("manual");
  const [extractDrag, setExtractDrag] = useState(false);
  const [extracting, setExtracting] = useState(false);
  const [extractJson, setExtractJson] = useState("");
  const [extractedTable, setExtractedTable] = useState<ExtractedTable | null>(null);
  const [excludedColumns, setExcludedColumns] = useState<number[]>([]);
  const [extractNotice, setExtractNotice] = useState("");
  const [extractError, setExtractError] = useState("");
  const [extractResultTick, setExtractResultTick] = useState(0);
  const [extractSource, setExtractSource] = useState<{ page?: number; method?: string; confidence?: number | null } | null>(null);
  const [extractPreviewUrl, setExtractPreviewUrl] = useState<string | null>(null);
  const [extractPreviewLoading, setExtractPreviewLoading] = useState(false);
  const [extractPage, setExtractPage] = useState(1);
  const [extractPageCount, setExtractPageCount] = useState(1);
  const [extractCrop, setExtractCrop] = useState<CropRect>({ x: 0, y: 0, width: 0, height: 0 });
  const [bulkFiles, setBulkFiles] = useState<File[]>([]);
  const [bulkDrag, setBulkDrag] = useState(false);
  const [bulkExtracting, setBulkExtracting] = useState(false);
  const [bulkDownloading, setBulkDownloading] = useState(false);
  const [bulkResults, setBulkResults] = useState<BulkExtractResult[]>([]);
  const [bulkActiveId, setBulkActiveId] = useState<string | null>(null);
  const [bulkNotice, setBulkNotice] = useState("");
  const [formatterFile, setFormatterFile] = useState<File | null>(null);
  const [formatterDrag, setFormatterDrag] = useState(false);
  const [formatterLoading, setFormatterLoading] = useState(false);
  const [formatterConverting, setFormatterConverting] = useState(false);
  const [formatterNotice, setFormatterNotice] = useState("");
  const [formatterAnalysis, setFormatterAnalysis] = useState<FormatAnalysis | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const imgRef = useRef<HTMLInputElement>(null);
  const extractRef = useRef<HTMLInputElement>(null);
  const bulkExtractRef = useRef<HTMLInputElement>(null);
  const formatterRef = useRef<HTMLInputElement>(null);
  const extractPreviewImgRef = useRef<HTMLImageElement>(null);
  const cropStageRef = useRef<HTMLDivElement>(null);
  const extractResultRef = useRef<HTMLDivElement>(null);
  const cropStartRef = useRef<{ x: number; y: number } | null>(null);
  const cropInteractionRef = useRef<CropInteraction | null>(null);
  const poll = useRef<ReturnType<typeof setInterval> | null>(null);
  const finishTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const completionHandled = useRef(false);
  const activeBulkResult = bulkResults.find(result => result.id === bulkActiveId) || null;
  const bulkReadyCount = bulkResults.filter(result => result.status === "ready").length;
  const isGenerating = stage === "processing";
  const isGenerationBusy = isGenerating || validatingDuplicates;

  const fetchImages = useCallback(async (requestedPage = galleryPage) => {
    try {
      const params = new URLSearchParams({
        page: String(requestedPage),
        limit: String(galleryPageSize),
      });
      if (imageSearch.trim()) params.set("search", imageSearch.trim());
      const r = await fetch(`${API_BASE}/image/list?${params.toString()}`);
      if (r.ok) {
        const d = await r.json();
        const nextImages = d.items || d.data || [];
        setImages(nextImages);
        setImageTotal(Number(d.total ?? nextImages.length));
        setGalleryPage(Number(d.page) || requestedPage);
      } else {
        setGalleryNotice(await getApiError(r, "Image library could not be refreshed."));
      }
    } catch {
      setGalleryNotice("Image library could not be refreshed.");
    }
  }, [galleryPage, galleryPageSize, imageSearch]);

  useEffect(() => {
    if (activeTab === "images") fetchImages();
  }, [activeTab, fetchImages]);

  useEffect(() => {
    setGalleryPage(1);
  }, [imageSearch, galleryPageSize]);

  useEffect(() => {
    return () => {
      if (poll.current) clearInterval(poll.current);
      if (finishTimer.current) clearTimeout(finishTimer.current);
    };
  }, []);

  useEffect(() => {
    if (stage !== "processing") return;

    const progressTimer = setInterval(() => {
      setPct((current) => (current < 95 ? current + 1 : current));
    }, 60);

    return () => clearInterval(progressTimer);
  }, [stage]);

  useEffect(() => {
    return () => {
      if (previewUrl) URL.revokeObjectURL(previewUrl);
    };
  }, [previewUrl]);

  useEffect(() => {
    return () => {
      if (extractPreviewUrl) URL.revokeObjectURL(extractPreviewUrl);
    };
  }, [extractPreviewUrl]);

  const uploadImages = async (files: File[]) => {
    const imageFiles = files.filter((f) => f.type.startsWith("image/"));
    if (!imageFiles.length) {
      setGalleryNotice("Select image files only.");
      return;
    }

    setGalleryNotice("");
    setSync({ up: 0, total: imageFiles.length, loading: true });
    const chunkSize = 50;
    let uploaded = 0;
    for (let i = 0; i < imageFiles.length; i += chunkSize) {
      const fd = new FormData();
      const chunk = imageFiles.slice(i, i + chunkSize);
      chunk.forEach((f) => fd.append("files", f));
      try {
        const r = await fetch(`${API_BASE}/image/upload`, { method: "POST", body: fd });
        if (r.ok) {
          const data = await r.json();
          uploaded += data.uploadedCount || chunk.length;
          setSync((s) => ({ ...s, up: Math.min(s.up + chunk.length, imageFiles.length) }));
        }
      } catch {
        setGalleryNotice("Some images could not be uploaded.");
      }
    }
    setSync((s) => ({ ...s, loading: false }));
    setGalleryNotice(`${uploaded} image${uploaded !== 1 ? "s" : ""} uploaded.`);
    if (imgRef.current) imgRef.current.value = "";
    await fetchImages(1);
  };

  const imageLink = (img: Img) => {
    const sourceName = img.filename || img.path?.split("/").pop() || "";
    return assetUrl(img.copyLink || img.url) || `${STATIC_BASE}/static/images/${encodeURIComponent(sourceName)}`;
  };

  const copyImageLink = async (img: Img) => {
    const link = imageLink(img);
    try {
      await navigator.clipboard.writeText(link);
    } catch {
      const input = document.createElement("textarea");
      input.value = link;
      document.body.appendChild(input);
      input.select();
      document.execCommand("copy");
      document.body.removeChild(input);
    }
    setGalleryNotice("Image link copied.");
  };

  const deleteImage = async (img: Img) => {
    if (!window.confirm(`Delete ${img.filename}?`)) return;
    const r = await fetch(`${API_BASE}/image/${img.id}`, { method: "DELETE" });
    if (r.ok) {
      setGalleryNotice("Image deleted.");
      await fetchImages();
    } else {
      setGalleryNotice("Delete failed.");
    }
  };

  const saveRename = async () => {
    if (!renaming) return;
    const nextName = renaming.value.trim();
    if (!nextName) {
      setGalleryNotice("Name cannot be empty.");
      return;
    }
    const r = await fetch(`${API_BASE}/image/${renaming.id}/rename`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename: nextName }),
    });
    if (r.ok) {
      const updated = await r.json();
      setImages((prev) => prev.map((item) => (item.id === updated.id ? updated : item)));
      setRenaming(null);
      setGalleryNotice("Image renamed.");
    } else {
      const err = await r.json().catch(() => ({}));
      setGalleryNotice(err.message || "Rename failed.");
    }
  };

  const showPreview = async () => {
    if (!file) return;
    setPreviewing(true);
    setJobNotice("Opening a fast PDF preview...");
    setPreviewUrl(null);
    setTimeout(() => {
      document.getElementById("preview-panel")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 50);
    const fd = new FormData();
    fd.append("file", file);
    fd.append("headerMode", headerMode);
    try {
      const r = await fetch(`${API_BASE}/pdf/preview`, { method: "POST", body: fd });
      if (!r.ok) throw new Error(await getApiError(r, "Preview could not be created."));
      const b = await r.blob();
      const url = URL.createObjectURL(b);
      setPreviewUrl(url);
      setJobNotice("Preview ready. Check it below, then generate your PDFs.");
    } catch (error) {
      setPreviewUrl(null);
      setJobNotice(error instanceof Error ? error.message : "Preview could not be created.");
    } finally {
      setPreviewing(false);
    }
  };

  const triggerDownload = async (href: string, name: string) => {
    try {
      const res = await fetch(href);
      if (!res.ok) throw new Error(`HTTP error ${res.status}`);
      const contentDisposition = res.headers.get("Content-Disposition");
      let finalName = name;
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="?([^";]+)"?/);
        if (match && match[1]) {
          finalName = match[1];
        }
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = finalName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => window.URL.revokeObjectURL(url), 10000);
    } catch (err) {
      console.warn("Direct blob download failed, falling back to window link", err);
      const a = document.createElement("a");
      a.href = href;
      a.download = name;
      a.target = "_blank";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
  };

  const analyzeFormatterFile = async (sourceFile: File) => {
    setFormatterLoading(true);
    setFormatterNotice("Checking columns and preparing the PDF-ready mapping...");
    setFormatterAnalysis(null);
    const fd = new FormData();
    fd.append("file", sourceFile);
    try {
      const response = await fetch(`${API_BASE}/format-xlsx/analyze`, { method: "POST", body: fd });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.message || "The workbook could not be analysed.");
      setFormatterAnalysis(data as FormatAnalysis);
      setFormatterNotice(`${data.outputRows} product rows are ready for the uniform PDF workbook.`);
    } catch (error) {
      setFormatterNotice(error instanceof Error ? error.message : "The workbook could not be analysed.");
    } finally {
      setFormatterLoading(false);
    }
  };

  const selectFormatterFile = (nextFile: File) => {
    if (!/\.(xlsx|xls)$/i.test(nextFile.name)) {
      setFormatterNotice("Select an XLSX or XLS workbook.");
      return;
    }
    setFormatterFile(nextFile);
    void analyzeFormatterFile(nextFile);
  };

  const downloadUniformWorkbook = async () => {
    if (!formatterFile || !formatterAnalysis) return;
    setFormatterConverting(true);
    setFormatterNotice("Building the uniform XLSX workbook...");
    const fd = new FormData();
    fd.append("file", formatterFile);
    try {
      const response = await fetch(`${API_BASE}/format-xlsx/convert`, { method: "POST", body: fd });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || "The uniform workbook could not be created.");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const serverName = disposition.match(/filename="?([^";]+)"?/i)?.[1];
      const fallbackName = `${formatterFile.name.replace(/\.(xlsx|xls)$/i, "")}_PDF_Ready.xlsx`;
      const url = URL.createObjectURL(blob);
      triggerDownload(url, serverName || fallbackName);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setFormatterNotice("Uniform XLSX downloaded successfully.");
    } catch (error) {
      setFormatterNotice(error instanceof Error ? error.message : "The uniform workbook could not be created.");
    } finally {
      setFormatterConverting(false);
    }
  };

  const resetFormatter = () => {
    setFormatterFile(null);
    setFormatterAnalysis(null);
    setFormatterNotice("");
    setFormatterDrag(false);
    if (formatterRef.current) formatterRef.current.value = "";
  };

  const loadExtractPreview = async (sourceFile: File, requestedPage = 1) => {
    setExtractPreviewLoading(true);
    setExtractNotice("Preparing a high-resolution crop preview...");
    const fd = new FormData();
    fd.append("file", sourceFile);
    fd.append("page", String(requestedPage));
    try {
      const response = await fetch(`${API_BASE}/extract/preview`, { method: "POST", body: fd });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || "Preview could not be prepared.");
      }
      const blob = await response.blob();
      const nextUrl = URL.createObjectURL(blob);
      setExtractPreviewUrl(nextUrl);
      setExtractPage(Number(response.headers.get("X-Page-Number")) || 1);
      setExtractPageCount(Number(response.headers.get("X-Page-Count")) || 1);
      setExtractCrop({ x: 0, y: 0, width: 0, height: 0 });
      setExtractNotice("Drag over the table area, then click Extract selected area.");
    } catch (error) {
      setExtractPreviewUrl(null);
      setExtractNotice(error instanceof Error ? error.message : "Preview could not be prepared.");
    } finally {
      setExtractPreviewLoading(false);
    }
  };

  const selectExtractFile = (nextFile: File) => {
    const supported = nextFile.type === "application/pdf" || nextFile.type.startsWith("image/") || /\.(pdf|png|jpe?g|webp|bmp|gif)$/i.test(nextFile.name);
    if (!supported) {
      setExtractNotice("Select a PDF or image file.");
      return;
    }
    setExtractFile(nextFile);
    setExtractJson("");
    setExtractedTable(null);
    setExcludedColumns([]);
    setExtractSource(null);
    setExtractNotice("");
    setExtractError("");
    setExtractPage(1);
    setExtractPageCount(1);
    void loadExtractPreview(nextFile, 1);
  };

  const changeExtractPage = (page: number) => {
    if (!extractFile || extractPreviewLoading || page < 1 || page > extractPageCount) return;
    setExtractJson("");
    setExtractedTable(null);
    setExcludedColumns([]);
    setExtractSource(null);
    setExtractError("");
    void loadExtractPreview(extractFile, page);
  };

  const pointerPosition = (event: React.PointerEvent<HTMLElement>) => {
    const bounds = cropStageRef.current?.getBoundingClientRect();
    if (!bounds) return null;
    return {
      x: Math.min(100, Math.max(0, ((event.clientX - bounds.left) / bounds.width) * 100)),
      y: Math.min(100, Math.max(0, ((event.clientY - bounds.top) / bounds.height) * 100)),
    };
  };

  const beginCrop = (event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    const point = pointerPosition(event);
    if (!point || extractPreviewLoading) return;
    event.preventDefault();
    cropStartRef.current = point;
    event.currentTarget.setPointerCapture(event.pointerId);
    setExtractCrop({ x: point.x, y: point.y, width: 0, height: 0 });
  };

  const beginCropInteraction = (event: React.PointerEvent<HTMLElement>, mode: CropInteractionMode) => {
    if (event.button !== 0 || extractPreviewLoading || extractCrop.width < 1 || extractCrop.height < 1) return;
    const point = pointerPosition(event);
    if (!point) return;
    event.preventDefault();
    event.stopPropagation();
    cropStartRef.current = null;
    cropInteractionRef.current = { mode, start: point, crop: { ...extractCrop } };
    cropStageRef.current?.setPointerCapture(event.pointerId);
  };

  const moveCrop = (event: React.PointerEvent<HTMLDivElement>) => {
    const point = pointerPosition(event);
    if (!point) return;

    const interaction = cropInteractionRef.current;
    if (interaction) {
      event.preventDefault();
      const minimumSize = 1;
      const dx = point.x - interaction.start.x;
      const dy = point.y - interaction.start.y;
      const original = interaction.crop;

      if (interaction.mode === "move") {
        setExtractCrop({
          ...original,
          x: Math.min(100 - original.width, Math.max(0, original.x + dx)),
          y: Math.min(100 - original.height, Math.max(0, original.y + dy)),
        });
        return;
      }

      let left = original.x;
      let top = original.y;
      let right = original.x + original.width;
      let bottom = original.y + original.height;
      if (interaction.mode.includes("w")) left = Math.min(right - minimumSize, Math.max(0, original.x + dx));
      if (interaction.mode.includes("e")) right = Math.max(left + minimumSize, Math.min(100, original.x + original.width + dx));
      if (interaction.mode.includes("n")) top = Math.min(bottom - minimumSize, Math.max(0, original.y + dy));
      if (interaction.mode.includes("s")) bottom = Math.max(top + minimumSize, Math.min(100, original.y + original.height + dy));
      setExtractCrop({ x: left, y: top, width: right - left, height: bottom - top });
      return;
    }

    const start = cropStartRef.current;
    if (!start) return;
    event.preventDefault();
    setExtractCrop({
      x: Math.min(start.x, point.x),
      y: Math.min(start.y, point.y),
      width: Math.abs(point.x - start.x),
      height: Math.abs(point.y - start.y),
    });
  };

  const endCrop = (event: React.PointerEvent<HTMLDivElement>) => {
    if (!cropStartRef.current && !cropInteractionRef.current) return;
    event.preventDefault();
    cropStartRef.current = null;
    cropInteractionRef.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
    setExtractCrop((current) => current.width < 1 || current.height < 1 ? { x: 0, y: 0, width: 0, height: 0 } : current);
  };

  // After a successful extraction, bring the generated JSON into view so the
  // user does not have to scroll away from the crop to find it.
  useEffect(() => {
    if (extractResultTick > 0) extractResultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }, [extractResultTick]);

  const extractTableJson = async () => {
    if (!extractFile || !extractPreviewUrl || !extractPreviewImgRef.current) return;
    if (extractCrop.width < 1 || extractCrop.height < 1) {
      setExtractNotice("Drag over the table first, then click Extract selected area.");
      return;
    }
    setExtracting(true);
    setExtractError("");
    setExtractNotice("Reading only the selected table area...");
    try {
      const image = extractPreviewImgRef.current;
      const sx = Math.round((extractCrop.x / 100) * image.naturalWidth);
      const sy = Math.round((extractCrop.y / 100) * image.naturalHeight);
      const sw = Math.max(1, Math.round((extractCrop.width / 100) * image.naturalWidth));
      const sh = Math.max(1, Math.round((extractCrop.height / 100) * image.naturalHeight));
      const canvas = document.createElement("canvas");
      canvas.width = sw;
      canvas.height = sh;
      const context = canvas.getContext("2d");
      if (!context) throw new Error("The selected area could not be prepared.");
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, sw, sh);
      context.imageSmoothingEnabled = true;
      context.imageSmoothingQuality = "high";
      context.drawImage(image, sx, sy, sw, sh, 0, 0, sw, sh);
      const fd = new FormData();
      const baseName = extractFile.name.replace(/\.[^.]+$/, "") || "table";
      const isPdf = extractFile.type === "application/pdf" || /\.pdf$/i.test(extractFile.name);
      if (isPdf) {
        fd.append("file", extractFile, extractFile.name);
        fd.append("page", String(extractPage));
        fd.append("cropX", String(extractCrop.x));
        fd.append("cropY", String(extractCrop.y));
        fd.append("cropWidth", String(extractCrop.width));
        fd.append("cropHeight", String(extractCrop.height));
      } else {
        const croppedBlob = await new Promise<Blob>((resolve, reject) => {
          canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("Crop creation failed.")), "image/png", 1);
        });
        fd.append("file", croppedBlob, `${baseName}-page-${extractPage}-crop.png`);
      }
      const response = await fetch(`${API_BASE}/extract/table`, { method: "POST", body: fd });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.message || "Table extraction failed.");
      const table = normalizeExtractedTable(data?.json);
      if (!table || !table.headers.length || !table.rows.length) {
        throw new Error("Table extraction returned no usable headers and rows.");
      }
      setExtractedTable(table);
      setExcludedColumns([]);
      refreshFilteredJson(table, []);
      setExtractSource({ ...(isRecord(data?.source) ? data.source : {}), page: extractPage });
      setExtractNotice("Selected table converted to compatible JSON. You can edit, copy, or download it.");
      setExtractResultTick((tick) => tick + 1);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Table extraction failed.";
      setExtractNotice(message);
      setExtractError(message);
    } finally {
      setExtracting(false);
    }
  };

  const refreshFilteredJson = (table: ExtractedTable, excluded: number[]) => {
    const excludedSet = new Set(excluded);
    setExtractJson(JSON.stringify({
      ...(table.section.trim() ? { section: table.section.trim() } : {}),
      headers: table.headers.filter((_, index) => !excludedSet.has(index)),
      rows: table.rows.map(row => row.filter((_, index) => !excludedSet.has(index))),
    }));
  };

  const toggleExtractColumn = (columnIndex: number) => {
    if (!extractedTable) return;
    const next = excludedColumns.includes(columnIndex)
      ? excludedColumns.filter(index => index !== columnIndex)
      : [...excludedColumns, columnIndex];
    setExcludedColumns(next);
    refreshFilteredJson(extractedTable, next);
  };

  const setAllExtractColumns = (includeAll: boolean) => {
    if (!extractedTable) return;
    const next = includeAll ? [] : extractedTable.headers.map((_, index) => index);
    setExcludedColumns(next);
    refreshFilteredJson(extractedTable, next);
  };

  const renameExtractSection = (section: string) => {
    if (!extractedTable) return;
    const next = { ...extractedTable, section };
    setExtractedTable(next);
    refreshFilteredJson(next, excludedColumns);
  };

  const ensureValidExtractJson = () => {
    try {
      JSON.parse(extractJson);
      return true;
    } catch {
      setExtractNotice("Fix the JSON syntax before copying or downloading.");
      return false;
    }
  };

  const copyExtractJson = async () => {
    if (!extractJson || !ensureValidExtractJson()) return;
    await navigator.clipboard.writeText(extractJson);
    setExtractNotice("JSON copied. Paste it directly into a JSON table column in the XLSX.");
  };

  const downloadExtractJson = async () => {
    if (!extractJson || !ensureValidExtractJson()) return;
    setExtractNotice("Building XLSX with the filtered table JSON...");
    try {
      const response = await fetch(`${API_BASE}/extract/table/xlsx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ json: JSON.parse(extractJson), filename: extractFile?.name || "Extracted_Table_JSON" }),
      });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || "XLSX download could not be created.");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const serverName = disposition.match(/filename="?([^";]+)"?/i)?.[1] || "Extracted_Table_JSON.xlsx";
      const url = URL.createObjectURL(blob);
      triggerDownload(url, serverName);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setExtractNotice("Downloaded XLSX with the selected table JSON in the variant_json column.");
    } catch (error) {
      setExtractNotice(error instanceof Error ? error.message : "XLSX download could not be created.");
    }
  };

  const resetExtraction = () => {
    setExtractFile(null);
    setExtractJson("");
    setExtractedTable(null);
    setExcludedColumns([]);
    setExtractSource(null);
    setExtractNotice("");
    setExtractError("");
    setExtractPreviewUrl(null);
    setExtractPage(1);
    setExtractPageCount(1);
    setExtractCrop({ x: 0, y: 0, width: 0, height: 0 });
    if (extractRef.current) extractRef.current.value = "";
  };

  const filteredBulkTable = (result: BulkExtractResult): ExtractedTable | null => {
    if (!result.table) return null;
    const excluded = new Set(result.excludedColumns);
    return {
      section: result.table.section,
      headers: result.table.headers.filter((_, index) => !excluded.has(index)),
      rows: result.table.rows.map(row => row.filter((_, index) => !excluded.has(index))),
    };
  };

  const bulkJsonText = (result: BulkExtractResult) => {
    const table = filteredBulkTable(result);
    if (!table) return "";
    return JSON.stringify({
      ...(table.section.trim() ? { section: table.section.trim() } : {}),
      headers: table.headers,
      rows: table.rows,
    });
  };

  const copyBulkJson = async (result: BulkExtractResult) => {
    const json = bulkJsonText(result);
    if (!json) return;
    await navigator.clipboard.writeText(json);
    setBulkNotice(`Page ${result.source?.page || 1} filtered JSON copied.`);
  };

  const runBulkExtraction = async (selectedFiles: File[]) => {
    const pdfFiles = selectedFiles.filter(file => file.type === "application/pdf" || /\.pdf$/i.test(file.name)).slice(0, 30);
    if (!pdfFiles.length) {
      setBulkNotice("Select PDF files only.");
      return;
    }
    setBulkFiles(pdfFiles);
    setBulkResults([]);
    setBulkActiveId(null);
    setBulkExtracting(true);
    setBulkNotice(`Reading tables from ${pdfFiles.length} PDF${pdfFiles.length === 1 ? "" : "s"}...`);
    const formData = new FormData();
    pdfFiles.forEach(file => formData.append("files", file, file.name));
    try {
      const response = await fetch(`${API_BASE}/extract/bulk`, { method: "POST", body: formData });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.message || "Bulk table extraction failed.");
      const rawResults: unknown[] = Array.isArray(data?.results) ? data.results : [];
      const results = rawResults
        .filter(isRecord)
        .map((result, index): BulkExtractResult => {
          const table = normalizeExtractedTable(result.table);
          const source = isRecord(result.source)
            ? {
                page: Number(result.source.page) || undefined,
                pages: Array.isArray(result.source.pages)
                  ? result.source.pages.map(page => Number(page)).filter(Number.isFinite)
                  : undefined,
                method: tableCellText(result.source.method) || undefined,
                confidence: Number.isFinite(Number(result.source.confidence)) ? Number(result.source.confidence) : null,
              }
            : undefined;
          const status = result.status === "ready" && table ? "ready" : "failed";
          return {
            id: tableCellText(result.id) || `bulk-result-${index + 1}`,
            sku: tableCellText(result.sku),
            pdfName: tableCellText(result.pdfName),
            status,
            error: status === "failed" ? tableCellText(result.error) || "No usable table was returned." : undefined,
            table: table || undefined,
            source,
            tablesDetected: Number(result.tablesDetected) || undefined,
            excludedColumns: [],
          };
        });
      setBulkResults(results);
      setBulkActiveId(results.find((result: BulkExtractResult) => result.status === "ready")?.id || results[0]?.id || null);
      setBulkNotice(`${data.ready || 0} page table${data.ready === 1 ? " is" : "s are"} ready from ${data.files || pdfFiles.length} PDF${(data.files || pdfFiles.length) === 1 ? "" : "s"}. Review columns, then download one XLSX.`);
    } catch (error) {
      setBulkNotice(error instanceof Error ? error.message : "Bulk table extraction failed.");
    } finally {
      setBulkExtracting(false);
      if (bulkExtractRef.current) bulkExtractRef.current.value = "";
    }
  };

  const toggleBulkColumn = (resultId: string, columnIndex: number) => {
    setBulkResults(current => current.map(result => {
      if (result.id !== resultId) return result;
      const excludedColumns = result.excludedColumns.includes(columnIndex)
        ? result.excludedColumns.filter(index => index !== columnIndex)
        : [...result.excludedColumns, columnIndex];
      return { ...result, excludedColumns };
    }));
  };

  const setBulkColumns = (resultId: string, includeAll: boolean) => {
    setBulkResults(current => current.map(result => result.id === resultId
      ? { ...result, excludedColumns: includeAll ? [] : (result.table?.headers || []).map((_, index) => index) }
      : result));
  };

  const updateBulkSku = (resultId: string, sku: string) => {
    setBulkResults(current => current.map(result => result.id === resultId ? { ...result, sku } : result));
  };

  const downloadBulkWorkbook = async () => {
    const records = bulkResults
      .filter(result => result.status === "ready" && result.table && result.sku.trim())
      .map(result => ({ ...result, table: filteredBulkTable(result) }))
      .filter(result => result.table?.headers.length);
    if (!records.length) {
      setBulkNotice("Keep at least one column in a ready PDF before downloading.");
      return;
    }
    setBulkDownloading(true);
    setBulkNotice("Building the SKU-to-JSON XLSX workbook...");
    try {
      const response = await fetch(`${API_BASE}/extract/bulk/xlsx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ records }),
      });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || "XLSX download could not be created.");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const serverName = disposition.match(/filename="?([^";]+)"?/i)?.[1] || "Page_by_Page_Table_JSON.xlsx";
      const url = URL.createObjectURL(blob);
      triggerDownload(url, serverName);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setBulkNotice(`Downloaded ${records.length} SKU-to-JSON record${records.length === 1 ? "" : "s"}.`);
    } catch (error) {
      setBulkNotice(error instanceof Error ? error.message : "XLSX download could not be created.");
    } finally {
      setBulkDownloading(false);
    }
  };

  const resetBulkExtraction = () => {
    setBulkFiles([]);
    setBulkResults([]);
    setBulkActiveId(null);
    setBulkNotice("");
    if (bulkExtractRef.current) bulkExtractRef.current.value = "";
  };

  const uploadExcel = async (f: File, cleanupDuplicates = false) => {
    if (stage === "processing") return;
    if (!cleanupDuplicates) {
      setValidatingDuplicates(true);
      setDuplicateValidation(null);
      setDuplicateReviewOpen(false);
      setJobNotice("Checking specification data for duplicates...");
      const validationForm = new FormData();
      validationForm.append("file", f);
      try {
        const validationResponse = await fetch(`${API_BASE}/job/validate`, { method: "POST", body: validationForm });
        if (!validationResponse.ok) throw new Error(await getApiError(validationResponse, "Workbook validation could not be completed."));
        const validation = await validationResponse.json() as DuplicateValidation;
        if (Number(validation?.duplicateCount) > 0) {
          setDuplicateValidation({
            duplicateProductCount: Number(validation.duplicateProductCount) || 0,
            duplicateCount: Number(validation.duplicateCount) || 0,
            duplicates: Array.isArray(validation.duplicates) ? validation.duplicates : [],
          });
          setJobNotice(`${Number(validation.duplicateProductCount) || 0} product${Number(validation.duplicateProductCount) === 1 ? " contains" : "s contain"} duplicate specification data. Review duplicates before generating PDFs.`);
          return;
        }
      } catch (error) {
        setJobNotice(error instanceof Error ? error.message : "Workbook validation could not be completed.");
        return;
      } finally {
        setValidatingDuplicates(false);
      }
    }
    if (poll.current) clearInterval(poll.current);
    if (finishTimer.current) clearTimeout(finishTimer.current);
    setFile(f);
    setStage("processing");
    setPct(1);
    setJobCounts({ total: 0, success: 0, failed: 0 });
    completionHandled.current = false;
    setPreviewUrl(null);
    setJobNotice("Your spreadsheet is being converted. Keep this page open.");
    const fd = new FormData();
    fd.append("file", f);
    fd.append("headerMode", headerMode);
    if (cleanupDuplicates) fd.append("cleanupDuplicates", "true");
    try {
      const r = await fetch(`${API_BASE}/job/upload`, { method: "POST", body: fd });
      if (!r.ok) throw new Error(await getApiError(r, "PDF generation could not be started."));
      const j = await r.json();
      if (!j?.id) throw new Error("PDF generation could not be started: the backend did not return a job ID.");
      setJobId(j.id);
      setJobCounts({ total: Number(j.totalRows) || 0, success: 0, failed: 0 });
      const stopPolling = () => {
        if (poll.current) clearInterval(poll.current);
        poll.current = null;
      };
      const failGeneration = (message: string) => {
        if (completionHandled.current) return;
        completionHandled.current = true;
        stopPolling();
        setStage("idle");
        setJobNotice(message);
      };
      const pollJob = async () => {
        try {
          const rr = await fetch(`${API_BASE}/job/${j.id}`);
          if (!rr.ok) throw new Error(await getApiError(rr, "PDF generation status could not be retrieved."));
          const jj = await rr.json();
          const reportedProgress = Number(jj.progress) || 0;
          setPct((current) => Math.max(current, Math.min(reportedProgress, 95)));
          const total = Number(jj.totalRows) || 0;
          const success = Number(jj.successCount) || 0;
          const failed = Number(jj.failedCount) || 0;
          setJobCounts({ total, success, failed });

          if (jj.status === "FAILED") {
            const detailError = (jj.details || []).find((detail: { status?: string; errorLog?: string }) => detail.status === "FAILED")?.errorLog;
            failGeneration(jj.errorLog || detailError || "PDF generation failed.");
            return;
          }

          if (jj.status !== "COMPLETED" || completionHandled.current) return;
          completionHandled.current = true;
          stopPolling();
          const completed: Pdf[] = (jj.details || []).filter((d: Pdf) => d.status === "COMPLETED");
          const completedCount = success || completed.length;
          setPdfs(completed);
          setPct(100);
          setJobNotice(`${completedCount} PDF${completedCount === 1 ? " is" : "s are"} ready to download.`);
          finishTimer.current = setTimeout(() => {
            setStage("done");
            if (completed.length === 1) {
              triggerDownload(
                `${API_BASE}/pdf/download/${encodeURIComponent(completed[0].filename)}`,
                completed[0].filename,
              );
            } else if (completed.length > 1) {
              const zipName = file?.name ? `${file.name.replace(/\.[^/.]+$/, "")}.zip` : `datasheets_${j.id}.zip`;
              triggerDownload(`${API_BASE}/job/${j.id}/zip`, zipName);
            }
          }, 700);
        } catch (error) {
          failGeneration(error instanceof Error ? error.message : "PDF generation failed.");
        }
      };
      poll.current = setInterval(() => void pollJob(), 1500);
      void pollJob();
    } catch (error) {
      if (poll.current) clearInterval(poll.current);
      poll.current = null;
      setStage("idle");
      setJobNotice(error instanceof Error ? error.message : "PDF generation failed.");
    }
  };

  const selectExcelFile = (nextFile: File) => {
    if (!/\.xlsx?$/i.test(nextFile.name)) {
      setJobNotice("Please choose an XLSX or XLS spreadsheet file.");
      return;
    }
    setFile(nextFile);
    setPdfs([]);
    setJobCounts({ total: 0, success: 0, failed: 0 });
    setPreviewUrl(null);
    setDuplicateValidation(null);
    setDuplicateReviewOpen(false);
    setJobNotice(`${nextFile.name} selected. Preview it or generate the PDFs.`);
  };

  const resetJob = () => {
    if (poll.current) clearInterval(poll.current);
    if (finishTimer.current) clearTimeout(finishTimer.current);
    completionHandled.current = false;
    setStage("idle");
    setFile(null);
    setPreviewUrl(null);
    setJobId(null);
    setPct(0);
    setPdfs([]);
    setJobCounts({ total: 0, success: 0, failed: 0 });
    setDuplicateValidation(null);
    setDuplicateReviewOpen(false);
    setJobNotice("");
    if (fileRef.current) fileRef.current.value = "";
  };

  const workflowSteps = [
    { label: "File", state: file ? "done" : "active" },
    { label: "Preview", state: previewUrl ? "done" : file ? "active" : "idle" },
    { label: "Generate", state: stage === "done" ? "done" : stage === "processing" ? "active" : "idle" },
  ];

  const generatedCount = jobCounts.success || pdfs.length;
  const processedCount = jobCounts.success + jobCounts.failed;
  const currentPdfNumber = jobCounts.total > 0
    ? Math.min(processedCount + 1, jobCounts.total)
    : 1;

  const totalGalleryPages = Math.max(1, Math.ceil(imageTotal / galleryPageSize));
  const currentGalleryPage = Math.min(galleryPage, totalGalleryPages);
  const galleryStart = imageTotal ? (currentGalleryPage - 1) * galleryPageSize : 0;
  const galleryEnd = Math.min(galleryStart + images.length, imageTotal);
  const visibleImages = images;

  useEffect(() => {
    setGalleryPage((page) => Math.min(Math.max(page, 1), totalGalleryPages));
  }, [totalGalleryPages]);

  const workspaceMeta = activeTab === "xlsx"
    ? {
        eyebrow: "PDF generator",
        title: "Create product datasheets",
        description: "Upload your workbook, preview the result, and download the generated PDFs.",
        icon: "auto_awesome",
      }
    : activeTab === "images"
      ? {
          eyebrow: "Asset workspace",
          title: "Product image library",
          description: "Upload, search, and manage the product images used in your datasheets.",
          icon: "photo_library",
        }
      : activeTab === "extract"
        ? {
          eyebrow: "Smart extraction",
          title: "Extract table JSON",
          description: "Select a table from a PDF or image and convert it into structured JSON.",
          icon: "data_object",
          }
        : {
            eyebrow: "Workbook formatter",
            title: "Create uniform XLSX",
            description: "Upload a supplier workbook and convert it into the consistent structure required by the PDF generator.",
            icon: "table_convert",
            };

  const assuranceSteps = activeTab === "format"
    ? ["Choose workbook", "Review mapping", "Download XLSX"]
    : activeTab === "extract"
      ? ["Upload source", "Select table", "Export JSON"]
      : ["Choose workbook", "Check preview", "Export PDFs"];

  return (
    <div className={`root premium-ui tab-${activeTab}`}>
      <header className="saas-header">
        <div className="top-brand">
          <div className="brand-mark">
            <img src="/logo.jpg" alt="RS PRO" className="logo" />
          </div>
          <div className="brand-product">
            <strong>Datasheet Studio</strong>
            <span>Content operations</span>
          </div>
        </div>

        <nav className="top-tabs" aria-label="Workspace tools">
          <button className={`tab ${activeTab === "xlsx" ? "on" : ""}`} aria-pressed={activeTab === "xlsx"} onClick={() => setActiveTab("xlsx")}>
            <span className="nav-icon"><MaterialIcon name="picture_as_pdf" /></span>
            <span className="nav-copy"><strong>Create PDFs</strong></span>
          </button>
          <button className={`tab ${activeTab === "images" ? "on" : ""}`} aria-pressed={activeTab === "images"} onClick={() => setActiveTab("images")}>
            <span className="nav-icon"><MaterialIcon name="photo_library" /></span>
            <span className="nav-copy"><strong>Image Library</strong></span>
          </button>
          <button className={`tab ${activeTab === "extract" ? "on" : ""}`} aria-pressed={activeTab === "extract"} onClick={() => setActiveTab("extract")}>
            <span className="nav-icon"><MaterialIcon name="data_object" /></span>
            <span className="nav-copy"><strong>Table JSON</strong></span>
          </button>
          <button className={`tab ${activeTab === "format" ? "on" : ""}`} aria-pressed={activeTab === "format"} onClick={() => setActiveTab("format")}>
            <span className="nav-icon"><MaterialIcon name="table_convert" /></span>
            <span className="nav-copy"><strong>XLSX Formatter</strong></span>
          </button>
        </nav>

      </header>

      <div className="product-frame">
        <header className="product-topbar">
          <div className="breadcrumb"><span>Product tools</span><MaterialIcon name="chevron_right" /><strong>{workspaceMeta.eyebrow}</strong></div>
          <div className="topbar-actions">
            <span className="workspace-note"><MaterialIcon name="lock" /> Files stay in your local workspace</span>
            <span className="top-status" aria-label="System status"><span className="status-dot" />Live</span>
          </div>
        </header>

        <div className="app-shell">
        <header className="page-header">
          <div className="page-heading">
            <span className="page-eyebrow"><MaterialIcon name={workspaceMeta.icon} />{workspaceMeta.eyebrow}</span>
            <h1 className="gallery-title" aria-label={workspaceMeta.title}>
              {workspaceMeta.title.split(" ").map((word, index) => (
                <span key={`${word}-${index}`} style={{ "--word-index": index } as React.CSSProperties}>{word}</span>
              ))}
            </h1>
            <p className="workspace-description">{workspaceMeta.description}</p>
            <div className="confidence-row" aria-label="Product benefits">
              <span><MaterialIcon name="lock" /> Local workspace</span>
              <span><MaterialIcon name="visibility" /> Preview before export</span>
              <span><MaterialIcon name="stacks" /> Built for batches</span>
            </div>
          </div>
          <div className="header-assurance" aria-label="How it works">
            <span>Simple three-step workflow</span>
            <ol>{assuranceSteps.map((step, index) => <li key={step}><b>{index + 1}</b>{step}</li>)}</ol>
          </div>
        </header>

        <main className="main">
        {activeTab === "xlsx" && (
          <>
            <section className="tool-panel">
              {stage === "idle" && (
                <div className="single-converter">
                  <div className="converter-toolbar">
                    <div>
                      <span className="workflow-label">Workbook to datasheet</span>
                      <h2>Three simple steps</h2>
                      <p>Choose your workbook, check the preview, and export polished product PDFs.</p>
                    </div>
                    <div className="step-row" aria-label="Generation steps">
                      {workflowSteps.map((step, index) => (
                        <span key={step.label} className={`step ${step.state}`}>
                          {step.state === "done" ? <MaterialIcon name="check" className="mini-icon" /> : <b>{index + 1}</b>}
                          {step.label}
                        </span>
                      ))}
                    </div>
                    <div className="workflow-detail">
                      <span><MaterialIcon name="verified" />Validated before generation</span>
                      <span><MaterialIcon name="lock" />Processed in your workspace</span>
                    </div>
                  </div>

                  {file ? (
                    <div className="inline-config-card">
                      <div className="modal-header">
                        <div className="modal-success-badge">
                          <MaterialIcon name="check_circle" className="success-icon" />
                          <span>Workbook loaded</span>
                        </div>
                        <button className="modal-close-btn-text" onClick={resetJob} title="Remove workbook">
                          <MaterialIcon name="delete" />
                          <span>Change file</span>
                        </button>
                      </div>

                      <div className="modal-file-row">
                        <MaterialIcon name="description" className="file-icon" />
                        <span className="file-name" title={file.name}>{file.name}</span>
                        <span className="file-size">({formatBytes(file.size)})</span>
                      </div>

                      <div className="modal-options-section">
                        <label className="options-title-label">PDF header layout:</label>
                        <div className="segmented-control" aria-label="PDF header layout">
                          <button type="button" className={`segment-btn ${headerMode === "supplier" ? "selected" : ""}`} onClick={() => { setHeaderMode("supplier"); setPreviewUrl(null); }}>
                            <span>Supplier Part Number</span>
                          </button>
                          <button type="button" className={`segment-btn ${headerMode === "rs" ? "selected" : ""}`} onClick={() => { setHeaderMode("rs"); setPreviewUrl(null); }}>
                            <span>RS Stock No:</span>
                          </button>
                          <button type="button" className={`segment-btn ${headerMode === "both" ? "selected" : ""}`} onClick={() => { setHeaderMode("both"); setPreviewUrl(null); }}>
                            <span>Both fields (RS Stock No: blank)</span>
                          </button>
                        </div>
                      </div>

                      <div className="modal-actions-footer">
                        <button className="btn secondary compact" onClick={showPreview} disabled={previewing}>
                          {previewing ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="visibility" />}
                          <span>{previewing ? "Preparing preview..." : "Preview PDF"}</span>
                        </button>
                        <button className="btn primary compact" onClick={() => uploadExcel(file)} disabled={isGenerationBusy}>
                          <MaterialIcon name="auto_awesome" />
                          <span>{validatingDuplicates ? "Validating..." : "Generate PDFs"}</span>
                        </button>
                      </div>

                      {duplicateValidation && (
                        <section className="duplicate-validation" role="alert" aria-live="polite">
                          <div className="duplicate-validation-heading">
                            <div>
                              <strong>Duplicate specification found</strong>
                              <p>{duplicateValidation.duplicateProductCount} product{duplicateValidation.duplicateProductCount === 1 ? " contains" : "s contain"} duplicate specification data. Review before generating PDFs.</p>
                            </div>
                            <span>{duplicateValidation.duplicateCount}</span>
                          </div>
                          <div className="duplicate-validation-actions">
                            <button type="button" className="btn secondary compact" onClick={() => setDuplicateReviewOpen((open) => !open)}>
                              <MaterialIcon name="fact_check" />
                              <span>{duplicateReviewOpen ? "Hide Duplicates" : "Review Duplicates"}</span>
                            </button>
                            <button type="button" className="btn primary compact" onClick={() => uploadExcel(file, true)} disabled={isGenerationBusy}>
                              <MaterialIcon name="cleaning_services" />
                              <span>Continue After Cleanup</span>
                            </button>
                          </div>
                          {duplicateReviewOpen && (
                            <div className="duplicate-validation-list">
                              {duplicateValidation.duplicates.map((duplicate, index) => (
                                <article key={`${duplicate.sku}-${duplicate.attribute}-${duplicate.value}-${index}`}>
                                  <strong>{duplicate.productName}</strong>
                                  <span>SKU: {duplicate.sku}</span>
                                  <span>Attribute: {duplicate.attribute}</span>
                                  <span>Value: {duplicate.value}</span>
                                  <span>Occurrences: {duplicate.occurrences}</span>
                                  <small>Repeated: {duplicate.sourceReferences.slice(1).join("; ") || "Source reference unavailable"}</small>
                                </article>
                              ))}
                            </div>
                          )}
                        </section>
                      )}
                    </div>
                  ) : (
                    <div
                      className={`dropzone ${drag ? "dragging" : ""}`}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" || e.key === " ") fileRef.current?.click();
                      }}
                      onDragOver={(e) => {
                        e.preventDefault();
                        setDrag(true);
                      }}
                      onDragLeave={() => setDrag(false)}
                      onDrop={(e) => {
                        e.preventDefault();
                        setDrag(false);
                        const f = e.dataTransfer.files?.[0];
                        if (f) selectExcelFile(f);
                      }}
                      onClick={() => fileRef.current?.click()}
                    >
                      <input
                        ref={fileRef}
                        type="file"
                        accept=".xlsx,.xls"
                        hidden
                        onChange={(e) => {
                          const f = e.target.files?.[0];
                          if (f) selectExcelFile(f);
                        }}
                      />
                      <div className="drop-icon">
                        <MaterialIcon name="description" className="drop-file-icon" />
                        <span className="drop-status-badge"><MaterialIcon name="arrow_upward" /></span>
                      </div>
                      <span className="drop-label">XLSX to PDF</span>
                      <strong>Drop your workbook here</strong>
                      <small>or click to choose an XLSX or XLS file</small>
                      <div className="upload-hints" aria-label="Supported workbook details">
                        <span>XLSX</span><span>XLS</span><span>Batch ready</span>
                      </div>
                      <span className="drop-trust"><MaterialIcon name="lock" />Your workbook stays in this local workspace</span>
                    </div>
                  )}

                  <div className="converter-footer">
                    <div className="converter-promises">
                      <span><MaterialIcon name="lock" /> Processed locally</span>
                      <span><MaterialIcon name="visibility" /> Preview before export</span>
                      <span><MaterialIcon name="folder_zip" /> Batch ZIP download</span>
                    </div>
                  </div>
                </div>
              )}

              {stage === "processing" && (
                <div className="processing-panel">
                  <div className="ring" aria-label={`Generation progress ${pct}%`}>
                    <svg viewBox="0 0 120 120">
                      <circle cx="60" cy="60" r="50" fill="none" stroke="#E7EDF3" strokeWidth="10" />
                      <circle
                        cx="60"
                        cy="60"
                        r="50"
                        fill="none"
                        stroke="#CC0000"
                        strokeWidth="10"
                        strokeLinecap="round"
                        strokeDasharray="314"
                        strokeDashoffset={314 - pct * 3.14}
                      />
                    </svg>
                    <strong>{pct}%</strong>
                  </div>
                  <div className="processing-copy">
                    <span>
                      {jobCounts.total > 0
                        ? `Generating PDFs… ${currentPdfNumber} of ${jobCounts.total}`
                        : "Generating PDFs…"}
                    </span>
                    <h2>Building datasheets</h2>
                    <div className="progress-track">
                      <i style={{ width: `${pct}%` }} />
                    </div>
                  </div>
                </div>
              )}

              {stage === "done" && (
                <div className="done-panel">
                  <MaterialIcon name="task_alt" className="done-icon" />
                  <div>
                    <span>Generation complete</span>
                    <h2>{generatedCount} PDF{generatedCount !== 1 ? "s" : ""} ready</h2>
                  </div>
                  <div className="actions centered">
                    {pdfs.length === 1 && (
                      <button
                        className="btn primary"
                        onClick={() =>
                          triggerDownload(`${API_BASE}/pdf/download/${encodeURIComponent(pdfs[0].filename)}`, pdfs[0].filename)
                        }
                      >
                        <MaterialIcon name="download" />
                        <span>Download PDF</span>
                      </button>
                    )}
                    {pdfs.length > 1 && jobId && (
                      <button
                        className="btn primary"
                        onClick={() => {
                          const zipName = file?.name ? `${file.name.replace(/\.[^/.]+$/, "")}.zip` : `datasheets_${jobId}.zip`;
                          triggerDownload(`${API_BASE}/job/${jobId}/zip`, zipName);
                        }}
                      >
                        <MaterialIcon name="folder_zip" />
                        <span>Download ZIP</span>
                      </button>
                    )}
                    <button className="btn secondary" onClick={resetJob}>
                      <MaterialIcon name="add" />
                      <span>New job</span>
                    </button>
                  </div>
                </div>
              )}
            </section>

            {stage === "idle" && (previewUrl || previewing) && (
              <section id="preview-panel" className="preview-panel">
                <div className="preview-topbar">
                  <span>
                    <MaterialIcon name="description" />
                    PDF Preview
                  </span>
                  <button className="icon-action" title="Close preview" onClick={() => setPreviewUrl(null)}>
                    <MaterialIcon name="close" />
                  </button>
                </div>
                {previewUrl ? (
                  <iframe src={previewUrl} className="preview-frame" title="PDF Preview" />
                ) : (
                  <div className="preview-loading" role="status" aria-live="polite">
                    <span className="preview-loading-icon"><MaterialIcon name="description" /></span>
                    <div><strong>Preparing preview</strong><small>Rendering the first datasheet page…</small></div>
                    <span className="preview-loading-bar"><i /></span>
                  </div>
                )}
              </section>
            )}

            {jobNotice && (!(file && stage === "idle") || /failed|error/i.test(jobNotice)) && (
              <div className={`app-notice ${/ready|selected/i.test(jobNotice) ? "success" : ""}`} role="status" aria-live="polite">
                <MaterialIcon name={/ready|selected/i.test(jobNotice) ? "check_circle" : "info"} />
                <span>{jobNotice}</span>
              </div>
            )}

            {stage === "idle" && (
              <section className="experience-strip" aria-label="Conversion workflow">
                <div><span>01</span><MaterialIcon name="upload_file" /><strong>Choose your workbook</strong><p>Drop an XLSX or XLS file into the workspace.</p></div>
                <MaterialIcon name="arrow_forward" className="experience-arrow" />
                <div><span>02</span><MaterialIcon name="preview" /><strong>Check the result</strong><p>Preview the rendered PDF before committing.</p></div>
                <MaterialIcon name="arrow_forward" className="experience-arrow" />
                <div><span>03</span><MaterialIcon name="download_done" /><strong>Download with confidence</strong><p>Receive one PDF or a ZIP for larger batches.</p></div>
              </section>
            )}
          </>
        )}

        {activeTab === "images" && (
          <section className="gallery-view">
            <div className="gallery-top">
              <div className="gallery-title">
                <span>{imageTotal}</span>
                <div>
                  <h2>Product images</h2>
                  <p>Stored library assets</p>
                </div>
              </div>

              <div className="gallery-actions">
                <label className="search-wrap">
                  <MaterialIcon name="search" />
                  <input value={imageSearch} onChange={(e) => setImageSearch(e.target.value)} placeholder="Search products" />
                </label>
                <label className="page-size">
                  <span>Per page</span>
                  <select value={galleryPageSize} onChange={(e) => setGalleryPageSize(Number(e.target.value))}>
                    <option value={12}>12</option>
                    <option value={24}>24</option>
                    <option value={48}>48</option>
                    <option value={96}>96</option>
                  </select>
                </label>
                <button className="btn secondary compact" onClick={() => fetchImages()}>
                  <MaterialIcon name="refresh" />
                  <span>Refresh</span>
                </button>
                <button className="btn primary compact" onClick={() => imgRef.current?.click()} disabled={sync.loading}>
                  {sync.loading ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="upload" />}
                  <span>{sync.loading ? `${sync.up}/${sync.total}` : "Upload"}</span>
                </button>
              </div>
              <input
                ref={imgRef}
                type="file"
                accept="image/*"
                multiple
                hidden
                onChange={(e) => {
                  const fs = Array.from(e.target.files || []);
                  if (fs.length) uploadImages(fs);
                }}
              />
            </div>

            {galleryNotice && <div className="gallery-notice" role="status" aria-live="polite">{galleryNotice}</div>}

            <div className="grid">
              {visibleImages.map((img, i) => (
                <article key={img.id || `${img.filename}-${i}`} className="img-card" style={{ animationDelay: `${i * 0.025}s` }}>
                  <div className="img-h">
                    <img src={imageLink(img)} alt={img.filename} />
                  </div>
                  <div className="img-meta">
                    {renaming?.id === img.id ? (
                      <div className="rename-row">
                        <input
                          value={renaming.value}
                          onChange={(e) => setRenaming({ ...renaming, value: e.target.value })}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") saveRename();
                            if (e.key === "Escape") setRenaming(null);
                          }}
                          autoFocus
                        />
                        <button className="icon-action ok" title="Save name" onClick={saveRename}>
                          <MaterialIcon name="check" />
                        </button>
                        <button className="icon-action" title="Cancel rename" onClick={() => setRenaming(null)}>
                          <MaterialIcon name="close" />
                        </button>
                      </div>
                    ) : (
                      <>
                        <strong title={img.filename}>{img.filename}</strong>
                        <span>{formatBytes(img.sizeBytes)}</span>
                      </>
                    )}
                  </div>
                  <div className="img-tools">
                    <button className="tool-btn" title="Copy image link" onClick={() => copyImageLink(img)}>
                      <MaterialIcon name="content_copy" />
                    </button>
                    <button className="tool-btn" title="Download image" onClick={() => triggerDownload(`${API_BASE}/image/download/${encodeURIComponent(img.filename)}`, img.filename)}>
                      <MaterialIcon name="download" />
                    </button>
                    <button className="tool-btn" title="Rename image" onClick={() => setRenaming({ id: img.id, value: img.filename })}>
                      <MaterialIcon name="edit" />
                    </button>
                    <button className="tool-btn danger" title="Delete image" onClick={() => deleteImage(img)}>
                      <MaterialIcon name="delete" />
                    </button>
                  </div>
                </article>
              ))}
              {images.length === 0 && (
                <div className="empty-gallery">
                  <MaterialIcon name="image_not_supported" className="empty-icon" />
                  <p>No product images yet.</p>
                </div>
              )}
            </div>

            {images.length > 0 && (
              <div className="pagination-bar" aria-label="Gallery pagination">
                <div className="page-range">
                  Showing <strong>{galleryStart + 1}-{galleryEnd}</strong> of <strong>{imageTotal}</strong>
                </div>
                <div className="pager-controls">
                  <button
                    className="icon-action pager-btn"
                    title="Previous page"
                    disabled={currentGalleryPage <= 1}
                    onClick={() => setGalleryPage((page) => Math.max(1, page - 1))}
                  >
                    <MaterialIcon name="chevron_left" />
                  </button>
                  <span className="page-count">Page {currentGalleryPage} / {totalGalleryPages}</span>
                  <button
                    className="icon-action pager-btn"
                    title="Next page"
                    disabled={currentGalleryPage >= totalGalleryPages}
                    onClick={() => setGalleryPage((page) => Math.min(totalGalleryPages, page + 1))}
                  >
                    <MaterialIcon name="chevron_right" />
                  </button>
                </div>
              </div>
            )}
          </section>
        )}

        {activeTab === "format" && (
          <section className="formatter-view">
            <div className="formatter-grid">
              <div
                className={`formatter-dropzone ${formatterDrag ? "dragging" : ""} ${formatterFile ? "has-file" : ""}`}
                role="button"
                tabIndex={0}
                onClick={() => formatterRef.current?.click()}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") formatterRef.current?.click();
                }}
                onDragOver={(event) => { event.preventDefault(); setFormatterDrag(true); }}
                onDragLeave={() => setFormatterDrag(false)}
                onDrop={(event) => {
                  event.preventDefault();
                  setFormatterDrag(false);
                  const selected = event.dataTransfer.files?.[0];
                  if (selected) selectFormatterFile(selected);
                }}
              >
                <input
                  ref={formatterRef}
                  type="file"
                  accept=".xlsx,.xls"
                  hidden
                  onChange={(event) => {
                    const selected = event.target.files?.[0];
                    if (selected) selectFormatterFile(selected);
                  }}
                />
                <div className={`drop-icon ${formatterAnalysis ? "ready" : ""}`}>
                  {formatterLoading
                    ? <MaterialIcon name="progress_activity" className="icon spin" />
                    : <MaterialIcon name={formatterAnalysis ? "check_circle" : "table_view"} />}
                </div>
                <span className="drop-label">Supplier XLSX → Uniform XLSX</span>
                <strong title={formatterFile?.name}>{formatterFile?.name || "Drop supplier workbook here"}</strong>
                <small>{formatterFile ? `${formatBytes(formatterFile.size)} loaded` : "or click to choose an XLSX or XLS file"}</small>
                <span className="drop-trust"><MaterialIcon name="lock" />Processed in your local workspace</span>
              </div>

              <aside className="formatter-summary">
                <span className="formatter-kicker">PDF-ready output</span>
                <h2>Uniform workbook builder</h2>
                <p>Reorders core product fields, groups attributes into JSON columns, normalises feature HTML, and keeps the source data as a backup sheet.</p>

                {formatterAnalysis ? (
                  <>
                    <div className="formatter-metrics">
                      <span><strong>{formatterAnalysis.sourceRows}</strong> source rows</span>
                      <span><strong>{formatterAnalysis.sourceColumns}</strong> source columns</span>
                      <span><strong>{formatterAnalysis.outputColumns}</strong> uniform columns</span>
                    </div>
                    <div className="formatter-groups">
                      <small>JSON sections</small>
                      <div>{formatterAnalysis.groups.map(group => <span key={group}>{group}</span>)}</div>
                    </div>
                    {!!formatterAnalysis.dynamicFields?.length && (
                      <div className="formatter-groups">
                        <small>Dynamic fields preserved</small>
                        <div>{formatterAnalysis.dynamicFields.map(field => <span key={field}>{field}</span>)}</div>
                      </div>
                    )}
                    <button className="btn primary formatter-download" onClick={downloadUniformWorkbook} disabled={formatterConverting}>
                      {formatterConverting ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="download" />}
                      <span>{formatterConverting ? "Building workbook" : "Download uniform XLSX"}</span>
                    </button>
                    <button className="btn secondary" onClick={resetFormatter} disabled={formatterConverting}>
                      <MaterialIcon name="close" /><span>Clear</span>
                    </button>
                  </>
                ) : (
                  <div className="formatter-checklist">
                    <span><MaterialIcon name="check" /> Standard two-row PDF headers</span>
                    <span><MaterialIcon name="check" /> Automatic attribute grouping</span>
                    <span><MaterialIcon name="check" /> Source Backup + Mapping Guide</span>
                  </div>
                )}
              </aside>
            </div>

            {formatterNotice && (
              <div className={`app-notice ${formatterAnalysis ? "success" : ""}`} role="status" aria-live="polite">
                <MaterialIcon name={formatterAnalysis ? "check_circle" : "info"} />
                <span>{formatterNotice}</span>
              </div>
            )}

            {formatterAnalysis && (
              <div className="formatter-preview">
                <div className="formatter-preview-head">
                  <div><span>Output preview</span><h3>First five uniform rows</h3></div>
                  <span className="preview-badge">{formatterAnalysis.outputRows} products</span>
                </div>
                <div className="formatter-table-scroll">
                  <table>
                    <thead><tr>{formatterAnalysis.preview.headers.map((header, index) => <th key={`${header}-${index}`}>{header}</th>)}</tr></thead>
                    <tbody>
                      {formatterAnalysis.preview.rows.map((row, rowIndex) => (
                        <tr key={rowIndex}>{row.map((value, columnIndex) => <td key={columnIndex} title={value}>{value || "—"}</td>)}</tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </section>
        )}

        {activeTab === "extract" && (
          <section className="extract-view">
            <div className="extract-mode-bar">
              <div>
                <span className="extract-mode-kicker">Table extraction workflow</span>
                <strong>{extractMode === "manual" ? "Crop one table manually" : "Auto-extract tables from many PDFs"}</strong>
              </div>
              <div className="extract-mode-switch" role="tablist" aria-label="Table extraction mode">
                <button type="button" role="tab" aria-selected={extractMode === "manual"} className={extractMode === "manual" ? "active" : ""} onClick={() => setExtractMode("manual")}>
                  <MaterialIcon name="crop" /><span>Manual crop</span>
                </button>
                <button type="button" role="tab" aria-selected={extractMode === "bulk"} className={extractMode === "bulk" ? "active" : ""} onClick={() => setExtractMode("bulk")}>
                  <MaterialIcon name="library_add" /><span>Bulk PDF</span>
                </button>
              </div>
            </div>

            {extractMode === "manual" ? (
              <>
            <div className="extract-grid">
              <div
                className={`extract-dropzone ${extractDrag ? "dragging" : ""} ${extractFile ? "has-file" : ""}`}
                role="button"
                tabIndex={0}
                onClick={() => extractRef.current?.click()}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") extractRef.current?.click();
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  setExtractDrag(true);
                }}
                onDragLeave={() => setExtractDrag(false)}
                onDrop={(event) => {
                  event.preventDefault();
                  setExtractDrag(false);
                  const nextFile = event.dataTransfer.files?.[0];
                  if (nextFile) selectExtractFile(nextFile);
                }}
              >
                <input
                  ref={extractRef}
                  type="file"
                  accept="application/pdf,image/*,.pdf,.png,.jpg,.jpeg,.webp,.bmp,.gif"
                  hidden
                  onChange={(event) => {
                    const nextFile = event.target.files?.[0];
                    if (nextFile) selectExtractFile(nextFile);
                  }}
                />
                <div className="drop-icon">{extractFile ? <MaterialIcon name="check_circle" /> : <MaterialIcon name="cloud_upload" />}</div>
                <span className="drop-label">Source document</span>
                <strong title={extractFile?.name}>{extractFile?.name || "Upload image or PDF"}</strong>
                <small>{extractFile ? `${formatBytes(extractFile.size)} loaded` : "PDF, PNG, JPG, WEBP, BMP or GIF"}</small>
              </div>

              <aside className="extract-summary">
                <span className="extract-kicker">Compatible output</span>
                <h2>PDF-ready table JSON</h2>
                <p>Crop the table to get compact <code>section</code>, <code>headers</code>, and <code>rows</code> JSON matching your XLSX examples.</p>
                {extractSource && (
                  <div className="extract-facts">
                    <span>Page <strong>{extractSource.page || 1}</strong></span>
                    <span>Method <strong>{extractSource.method === "pdf-text" ? "PDF text" : "OCR"}</strong></span>
                    {extractSource.confidence != null && <span>Confidence <strong>{extractSource.confidence}%</strong></span>}
                  </div>
                )}
                <div className="actions">
                  <button className="btn primary" onClick={extractTableJson} disabled={!extractPreviewUrl || extractPreviewLoading || extracting || extractCrop.width < 1 || extractCrop.height < 1}>
                    {extracting ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="document_scanner" />}
                    <span>{extracting ? "Extracting" : "Extract selected area"}</span>
                  </button>
                  {extractFile && (
                    <button className="btn secondary" onClick={resetExtraction} disabled={extracting}>
                      <MaterialIcon name="close" />
                      <span>Clear</span>
                    </button>
                  )}
                </div>
              </aside>
            </div>

            {extractNotice && <div className="extract-notice" role="status" aria-live="polite">{extractNotice}</div>}

            {(extractPreviewUrl || extractPreviewLoading) && (
              <div className="crop-panel">
                <div className="crop-toolbar">
                  <div className="crop-title">
                    <MaterialIcon name="crop" />
                    <div>
                      <strong>Select the table area</strong>
                      <span>{extractCrop.width >= 1 && extractCrop.height >= 1 ? "Drag inside to move. Use the handles to resize." : "Drag from one corner of the table to the opposite corner."}</span>
                    </div>
                  </div>
                  <div className="crop-actions">
                    {extractPageCount > 1 && (
                      <div className="crop-pages">
                        <button className="icon-action pager-btn" title="Previous PDF page" disabled={extractPage <= 1 || extractPreviewLoading} onClick={() => changeExtractPage(extractPage - 1)}>
                          <MaterialIcon name="chevron_left" />
                        </button>
                        <span>Page {extractPage} / {extractPageCount}</span>
                        <button className="icon-action pager-btn" title="Next PDF page" disabled={extractPage >= extractPageCount || extractPreviewLoading} onClick={() => changeExtractPage(extractPage + 1)}>
                          <MaterialIcon name="chevron_right" />
                        </button>
                      </div>
                    )}
                    <button className="btn secondary compact" disabled={extractPreviewLoading || extractCrop.width < 1 || extractCrop.height < 1} onClick={() => setExtractCrop({ x: 0, y: 0, width: 0, height: 0 })}>
                      <MaterialIcon name="crop_free" />
                      <span>Clear crop</span>
                    </button>
                    <button className="btn primary compact" disabled={!extractPreviewUrl || extractPreviewLoading || extracting || extractCrop.width < 1 || extractCrop.height < 1} onClick={extractTableJson}>
                      {extracting ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="document_scanner" />}
                      <span>{extracting ? "Extracting" : "Extract"}</span>
                    </button>
                  </div>
                </div>

                <div className="crop-canvas-wrap">
                  {extractPreviewLoading ? (
                    <div className="crop-loading"><MaterialIcon name="progress_activity" className="icon spin" /><span>Rendering page in high quality...</span></div>
                  ) : extractPreviewUrl && (
                    <div
                      ref={cropStageRef}
                      className="crop-stage"
                      onPointerDown={beginCrop}
                      onPointerMove={moveCrop}
                      onPointerUp={endCrop}
                      onPointerCancel={endCrop}
                    >
                      <img ref={extractPreviewImgRef} src={extractPreviewUrl} alt={`Crop preview page ${extractPage}`} draggable={false} />
                      {extractCrop.width >= 1 && extractCrop.height >= 1 ? (
                        <div
                          className="crop-selection"
                          onPointerDown={(event) => beginCropInteraction(event, "move")}
                          aria-label="Selected crop area. Drag to move."
                          style={{
                            left: `${extractCrop.x}%`,
                            top: `${extractCrop.y}%`,
                            width: `${extractCrop.width}%`,
                            height: `${extractCrop.height}%`,
                          }}
                        >
                          <span className="crop-size-label">{Math.round(extractCrop.width)}% × {Math.round(extractCrop.height)}%</span>
                          {(["nw", "n", "ne", "e", "se", "s", "sw", "w"] as CropInteractionMode[]).map((handle) => (
                            <button
                              key={handle}
                              type="button"
                              className={`crop-handle crop-handle-${handle}`}
                              aria-label={`Resize selected area ${handle}`}
                              onPointerDown={(event) => beginCropInteraction(event, handle)}
                            />
                          ))}
                        </div>
                      ) : (
                        <div className="crop-empty-hint">
                          <MaterialIcon name="crop" />
                          <strong>Drag to select the table</strong>
                        </div>
                      )}
                    </div>
                  )}
                </div>

                {extractPreviewUrl && !extractPreviewLoading && (
                  <div className="crop-footer">
                    <span className={`crop-footer-status${extractError ? " is-error" : ""}`} role={extractError ? "alert" : "status"}>
                      {extractError
                        || (extracting
                          ? "Reading the selected table..."
                          : extractCrop.width >= 1 && extractCrop.height >= 1
                            ? `Selected ${Math.round(extractCrop.width)}% × ${Math.round(extractCrop.height)}% of the page. Ready to extract.`
                            : "Drag over the table to select it.")}
                    </span>
                    <button className="btn primary" disabled={extracting || extractCrop.width < 1 || extractCrop.height < 1} onClick={extractTableJson}>
                      {extracting ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="document_scanner" />}
                      <span>{extracting ? "Extracting" : "Extract selected area"}</span>
                    </button>
                  </div>
                )}
              </div>
            )}

            {extractJson && (
              <div className="json-panel" ref={extractResultRef}>
                {extractedTable && (
                  <div className="json-column-editor">
                    <div className="json-column-heading">
                      <div>
                        <strong>Choose columns for JSON</strong>
                        <span>Untick any header to remove that column and its values.</span>
                      </div>
                      <div className="json-column-actions">
                        <button type="button" onClick={() => setAllExtractColumns(true)}>Select all</button>
                        <button type="button" onClick={() => setAllExtractColumns(false)}>Remove all</button>
                      </div>
                    </div>
                    <label className="json-section-field">
                      <span>Section name</span>
                      <input placeholder="Optional" value={extractedTable.section} onChange={(event) => renameExtractSection(event.target.value)} />
                    </label>
                    <div className="json-column-list">
                      {extractedTable.headers.map((header, index) => (
                        <label key={`${header}-${index}`} className={excludedColumns.includes(index) ? "is-removed" : ""}>
                          <input
                            type="checkbox"
                            checked={!excludedColumns.includes(index)}
                            onChange={() => toggleExtractColumn(index)}
                          />
                          <span>{header}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                )}
                <div className="json-toolbar">
                  <div>
                    <span>Generated JSON</span>
                    <strong>Single-line JSON for XLSX</strong>
                  </div>
                  <div className="actions">
                    <button className="btn secondary compact" onClick={copyExtractJson}>
                      <MaterialIcon name="content_copy" />
                      <span>Copy JSON</span>
                    </button>
                    <button className="btn primary compact" onClick={downloadExtractJson}>
                      <MaterialIcon name="download" />
                      <span>Download XLSX</span>
                    </button>
                  </div>
                </div>
                <textarea
                  className="json-editor"
                  value={extractJson}
                  onChange={(event) => setExtractJson(event.target.value)}
                  spellCheck={false}
                  aria-label="Extracted table JSON"
                />
              </div>
            )}
              </>
            ) : (
              <div className="bulk-extract-workspace">
                <div className="bulk-extract-top">
                  <div
                    className={`bulk-extract-dropzone ${bulkDrag ? "dragging" : ""} ${bulkFiles.length ? "has-files" : ""}`}
                    role="button"
                    tabIndex={0}
                    onClick={() => bulkExtractRef.current?.click()}
                    onKeyDown={(event) => {
                      if (event.key === "Enter" || event.key === " ") bulkExtractRef.current?.click();
                    }}
                    onDragOver={(event) => { event.preventDefault(); setBulkDrag(true); }}
                    onDragLeave={() => setBulkDrag(false)}
                    onDrop={(event) => {
                      event.preventDefault();
                      setBulkDrag(false);
                      void runBulkExtraction(Array.from(event.dataTransfer.files || []));
                    }}
                  >
                    <input
                      ref={bulkExtractRef}
                      type="file"
                      accept="application/pdf,.pdf"
                      multiple
                      hidden
                      onChange={(event) => void runBulkExtraction(Array.from(event.target.files || []))}
                    />
                    <div className="bulk-drop-icon">
                      {bulkExtracting ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name={bulkFiles.length ? "task_alt" : "upload_file"} />}
                    </div>
                    <span>Bulk PDF → Table JSON → XLSX</span>
                    <strong>{bulkExtracting ? "Automatically finding tables..." : bulkFiles.length ? `${bulkFiles.length} PDF files loaded` : "Drop PDF files here"}</strong>
                    <small>PDF filename becomes the SKU. Up to 30 PDFs per batch.</small>
                  </div>

                  <aside className="bulk-extract-guide">
                    <span className="extract-kicker">Automatic workflow</span>
                    <h2>One row per PDF</h2>
                    <p>The complete PDF is read page by page. Every page containing a table becomes a separate JSON row in one XLSX.</p>
                    <div className="bulk-flow-list">
                      <span><b>1</b><span><strong>Upload PDFs</strong><small>Example: 677-112.pdf → SKU 677-112</small></span></span>
                      <span><b>2</b><span><strong>Review every page</strong><small>Untick unwanted headers and values per page</small></span></span>
                      <span><b>3</b><span><strong>Download one XLSX</strong><small>One page table JSON per spreadsheet row</small></span></span>
                    </div>
                  </aside>
                </div>

                {bulkNotice && <div className="extract-notice" role="status" aria-live="polite">{bulkNotice}</div>}

                {!!bulkResults.length && (
                  <div className="bulk-review-panel">
                    <div className="bulk-review-toolbar">
                      <div>
                        <span>Extraction review</span>
                        <strong>{bulkReadyCount} page tables ready <i>•</i> {bulkResults.length - bulkReadyCount} need attention</strong>
                      </div>
                      <div className="actions">
                        <button className="btn secondary compact" type="button" onClick={resetBulkExtraction} disabled={bulkDownloading}>
                          <MaterialIcon name="close" /><span>Clear batch</span>
                        </button>
                        <button className="btn primary compact" type="button" onClick={downloadBulkWorkbook} disabled={!bulkReadyCount || bulkDownloading}>
                          {bulkDownloading ? <MaterialIcon name="progress_activity" className="icon spin" /> : <MaterialIcon name="download" />}
                          <span>{bulkDownloading ? "Building XLSX" : "Download XLSX"}</span>
                        </button>
                      </div>
                    </div>

                    <div className="bulk-review-grid">
                      <div className="bulk-file-list" role="list" aria-label="Uploaded PDF extraction results">
                        {bulkResults.map(result => (
                          <button
                            type="button"
                            role="listitem"
                            key={result.id}
                            className={`${result.id === bulkActiveId ? "active" : ""} ${result.status === "failed" ? "failed" : ""}`}
                            onClick={() => setBulkActiveId(result.id)}
                          >
                            <MaterialIcon name={result.status === "ready" ? "check_circle" : "error"} />
                            <span><strong>{result.sku || result.pdfName} · Page {result.source?.page || 1}</strong><small>{result.pdfName}</small></span>
                            {result.table && <em>{result.table.rows.length} rows</em>}
                          </button>
                        ))}
                      </div>

                      <div className="bulk-result-detail">
                        {activeBulkResult?.status === "failed" ? (
                          <div className="bulk-result-error">
                            <MaterialIcon name="error" />
                            <div><strong>Table could not be extracted</strong><p>{activeBulkResult.error || "Use Manual crop mode for this PDF."}</p></div>
                          </div>
                        ) : activeBulkResult?.table ? (
                          <>
                            <div className="bulk-result-head">
                              <label>
                                <span>SKU from PDF filename</span>
                                <input value={activeBulkResult.sku} onChange={(event) => updateBulkSku(activeBulkResult.id, event.target.value)} />
                              </label>
                              <div className="bulk-result-facts">
                                <span>Page <strong>{activeBulkResult.source?.pages?.join(", ") || activeBulkResult.source?.page || 1}</strong></span>
                                <span>Rows <strong>{activeBulkResult.table.rows.length}</strong></span>
                                <span>Included <strong>{activeBulkResult.table.headers.length - activeBulkResult.excludedColumns.length}/{activeBulkResult.table.headers.length}</strong></span>
                              </div>
                            </div>

                            <div className="json-column-heading bulk-column-heading">
                              <div><strong>Filter columns</strong><span>Removed headers and their complete column values will not be saved.</span></div>
                              <div className="json-column-actions">
                                <button type="button" onClick={() => setBulkColumns(activeBulkResult.id, true)}>Select all</button>
                                <button type="button" onClick={() => setBulkColumns(activeBulkResult.id, false)}>Remove all</button>
                              </div>
                            </div>
                            <div className="json-column-list bulk-column-list">
                              {activeBulkResult.table.headers.map((header, index) => (
                                <label key={`${activeBulkResult.id}-${header}-${index}`} className={activeBulkResult.excludedColumns.includes(index) ? "is-removed" : ""}>
                                  <input type="checkbox" checked={!activeBulkResult.excludedColumns.includes(index)} onChange={() => toggleBulkColumn(activeBulkResult.id, index)} />
                                  <span>{header}</span>
                                </label>
                              ))}
                            </div>

                            <div className="bulk-json-preview">
                              <div>
                                <span>Final XLSX JSON</span>
                                <div className="bulk-json-preview-actions">
                                  <small>{filteredBulkTable(activeBulkResult)?.headers.length || 0} columns included</small>
                                  <button type="button" onClick={() => copyBulkJson(activeBulkResult)}>
                                    <MaterialIcon name="content_copy" /> Copy JSON
                                  </button>
                                </div>
                              </div>
                              <textarea value={bulkJsonText(activeBulkResult)} readOnly spellCheck={false} aria-label={`Filtered JSON for ${activeBulkResult.sku}`} />
                            </div>
                          </>
                        ) : (
                          <div className="bulk-result-empty"><MaterialIcon name="table_view" /><span>Select a PDF result to review its table.</span></div>
                        )}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}
          </section>
        )}
        </main>
        </div>
      </div>

      <style suppressHydrationWarning dangerouslySetInnerHTML={{ __html: `
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200');
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200');
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Manrope:wght@400;500;600;700;800&display=swap');

        :root {
          --rs: #CC0000;
          --rs-dark: #A90000;
          --ink: #17202A;
          --muted: #667085;
          --line: #D9E0E8;
          --panel: #FFFFFF;
          --soft: #F5F7FA;
          --soft-2: #EAF1F6;
          --teal: #0F766E;
          --green: #0F9F6E;
          --danger: #B42318;
          --rs-soft: #FFF0F0;
          --surface: #FFFFFF;
        }

        *, *::before, *::after { box-sizing: border-box; }
        html { background: var(--soft); }
        body {
          margin: 0;
          color: var(--ink);
          background:
            linear-gradient(180deg, #FFFFFF 0, #F7FAFC 220px, var(--soft) 520px);
          font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }
        button, input { font: inherit; }
        button { -webkit-tap-highlight-color: transparent; }
        .root { min-height: 100vh; }
        .material-symbols-rounded,
        .material-symbols-outlined {
          font-family: 'Material Symbols Rounded', 'Material Symbols Outlined' !important;
          font-weight: normal;
          font-style: normal;
          font-size: 22px;
          line-height: 1;
          letter-spacing: normal;
          text-transform: none;
          display: inline-block;
          white-space: nowrap;
          word-wrap: normal;
          direction: ltr;
          font-feature-settings: 'liga';
          -webkit-font-feature-settings: 'liga';
          -webkit-font-smoothing: antialiased;
          flex: 0 0 auto;
        }
        .icon { width: 22px; height: 22px; font-size: 22px; }
        .mini-icon { width: 18px; height: 18px; font-size: 18px; }

        .root { display: flex; min-height: 100vh; background: #F3F5F8; }
        .sidebar {
          position: fixed;
          inset: 0 auto 0 0;
          z-index: 30;
          width: 286px;
          padding: 22px 18px 18px;
          color: #E8EDF3;
          background: #111820;
          border-right: 1px solid rgba(255,255,255,.07);
          display: flex;
          flex-direction: column;
        }
        .sidebar-brand { display: flex; align-items: center; gap: 12px; padding: 2px 6px 26px; }
        .brand-mark { width: 70px; height: 42px; padding: 5px; border-radius: 9px; background: #FFFFFF; display: grid; place-items: center; }
        .logo { width: 60px; height: 30px; object-fit: contain; }
        .brand-copy { display: flex; flex-direction: column; gap: 3px; min-width: 0; }
        .brand-copy strong { color: #FFFFFF; font-size: 16px; font-weight: 850; white-space: nowrap; }
        .brand-copy span { color: #8D9AAA; font-size: 11px; font-weight: 750; letter-spacing: .03em; }
        .nav-label { padding: 0 10px 9px; color: #687687; font-size: 10px; font-weight: 900; letter-spacing: .14em; text-transform: uppercase; }
        .sidebar-nav { display: flex; flex-direction: column; gap: 7px; }
        .tab {
          border: 0;
          background: transparent;
          color: #AAB4C0;
          width: 100%;
          min-height: 66px;
          padding: 9px 10px;
          border-radius: 12px;
          display: grid;
          grid-template-columns: 40px minmax(0,1fr) 20px;
          align-items: center;
          gap: 10px;
          text-align: left;
          cursor: pointer;
          transition: color .2s, background .2s, transform .2s;
        }
        .tab.on {
          color: #FFFFFF;
          background: linear-gradient(135deg, #D71920, #A90000);
          box-shadow: 0 10px 28px rgba(204,0,0,.26);
        }
        .tab:hover:not(.on) { color: #FFFFFF; background: #1C2732; transform: translateX(2px); }
        .nav-icon { width: 40px; height: 40px; border-radius: 10px; display: grid; place-items: center; background: rgba(255,255,255,.06); }
        .tab.on .nav-icon { background: rgba(255,255,255,.17); }
        .nav-copy { min-width: 0; display: flex; flex-direction: column; gap: 3px; }
        .nav-copy strong { font-size: 13px; font-weight: 850; }
        .nav-copy small { overflow: hidden; color: #738192; font-size: 10px; font-weight: 650; text-overflow: ellipsis; white-space: nowrap; }
        .tab.on .nav-copy small { color: rgba(255,255,255,.72); }
        .nav-arrow { font-size: 18px; opacity: .55; }
        .sidebar-support { margin-top: auto; padding: 14px; border: 1px solid rgba(255,255,255,.08); border-radius: 13px; background: rgba(255,255,255,.04); display: flex; align-items: center; gap: 10px; }
        .support-icon { width: 36px; height: 36px; border-radius: 10px; color: #75DFC0; background: rgba(15,159,110,.14); display: grid; place-items: center; }
        .sidebar-support > div:last-child { min-width: 0; display: flex; flex-direction: column; gap: 3px; }
        .sidebar-support strong { color: #EDF2F7; font-size: 11px; }
        .sidebar-support span { color: #768494; font-size: 9px; }
        .sidebar-foot { padding: 14px 7px 0; color: #6F7D8D; font-size: 10px; font-weight: 700; display: flex; align-items: center; gap: 7px; }
        .live-dot { width: 7px; height: 7px; border-radius: 50%; background: #2DCC9A; box-shadow: 0 0 0 4px rgba(45,204,154,.1); }

        .app-shell { min-width: 0; width: calc(100% - 286px); margin-left: 286px; }
        .page-header {
          min-height: 188px;
          padding: 34px clamp(24px, 4vw, 58px) 30px;
          background: #FFFFFF;
          border-bottom: 1px solid #E1E6EC;
          display: flex;
          align-items: flex-end;
          justify-content: space-between;
          gap: 30px;
        }
        .page-heading { max-width: 780px; }
        .eyebrow {
          margin: 0 0 8px;
          color: var(--rs);
          font-size: 10px;
          font-weight: 900;
          letter-spacing: .14em;
          text-transform: uppercase;
        }
        h1, h2, p { margin-top: 0; }
        .page-header h1 {
          margin: 0;
          max-width: 780px;
          font-size: clamp(28px, 3.2vw, 42px);
          line-height: 1.1;
          letter-spacing: -.04em;
          font-weight: 850;
        }
        .workspace-description {
          max-width: 650px;
          margin: 10px 0 0;
          color: var(--muted);
          font-size: 14px;
          line-height: 1.55;
          font-weight: 650;
        }
        .header-metrics { display: flex; gap: 10px; flex: 0 0 auto; }
        .header-metrics > div { min-width: 132px; padding: 13px 14px; border: 1px solid #E1E6EC; border-radius: 13px; background: #F8FAFC; display: flex; align-items: center; gap: 10px; }
        .header-metrics > div > .material-symbols-rounded { width: 38px; height: 38px; border-radius: 10px; color: var(--rs); background: #FFF0F0; display: grid; place-items: center; }
        .header-metrics span { display: flex; flex-direction: column; gap: 1px; }
        .header-metrics strong { font-size: 18px; line-height: 1; }
        .header-metrics small { color: var(--muted); font-size: 9px; font-weight: 750; }

        .main { width: min(1280px, calc(100% - 48px)); margin: 0 auto; padding: 26px 0 64px; }

        .quick-guide {
          margin-bottom: 20px;
          padding: 12px 16px;
          border: 1px solid #E3E8EE;
          border-radius: 14px;
          background: rgba(255,255,255,.8);
          display: flex;
          align-items: center;
          gap: 8px;
          box-shadow: 0 8px 28px rgba(17,24,39,.04);
        }
        .quick-guide > div { min-width: 0; flex: 1; display: flex; align-items: center; gap: 9px; }
        .quick-guide > div > span {
          width: 30px;
          height: 30px;
          border-radius: 50%;
          display: grid;
          place-items: center;
          color: #FFFFFF;
          background: var(--rs);
          font-size: 12px;
          font-weight: 900;
        }
        .quick-guide strong { font-size: 13px; font-weight: 800; white-space: nowrap; }
        .guide-arrow { margin-left: auto; color: #A8B3C1; font-size: 20px; }

        .tool-panel, .gallery-view, .preview-panel {
          background: rgba(255,255,255,.96);
          border: 1px solid var(--line);
          border-radius: 18px;
          box-shadow: 0 18px 48px rgba(17,24,39,.07);
        }
        .tool-panel { padding: 22px; }
        .converter-grid {
          min-height: 390px;
          display: grid;
          grid-template-columns: minmax(0, 1.2fr) minmax(320px, .8fr);
          gap: 22px;
        }
        .dropzone {
          min-height: 344px;
          border: 2px dashed #B8C3CF;
          border-radius: 16px;
          background: #F8FBFD;
          color: var(--ink);
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          text-align: center;
          padding: 28px;
          cursor: pointer;
          transition: border-color .2s, background .2s, transform .2s;
          outline: none;
          min-width: 0;
        }
        .dropzone:hover, .dropzone:focus-visible, .dropzone.dragging {
          border-color: var(--rs);
          background: #FFF7F7;
          transform: translateY(-2px);
        }
        .dropzone.has-file { border-color: #86D8BA; background: #F4FBF8; }
        .drop-icon {
          width: 76px;
          height: 76px;
          border-radius: 50%;
          background: #FFFFFF;
          border: 1px solid var(--line);
          color: var(--rs);
          display: grid;
          place-items: center;
          margin-bottom: 16px;
          box-shadow: 0 12px 30px rgba(16,24,40,.08);
        }
        .dropzone.has-file .drop-icon { color: var(--green); }
        .drop-icon .material-symbols-rounded { width: 38px; height: 38px; font-size: 38px; }
        .drop-label {
          color: var(--muted);
          font-size: 12px;
          font-weight: 800;
          text-transform: uppercase;
          margin-bottom: 7px;
        }
        .dropzone strong {
          max-width: 100%;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          font-size: 22px;
          line-height: 1.2;
          font-weight: 800;
        }
        .dropzone small { margin-top: 8px; color: var(--muted); font-size: 14px; font-weight: 700; }

        .job-panel {
          border-left: 1px solid var(--line);
          padding-left: 22px;
          display: flex;
          flex-direction: column;
          justify-content: center;
          gap: 18px;
          min-width: 0;
        }
        .step-row { display: flex; flex-wrap: wrap; gap: 8px; }
        .step {
          height: 32px;
          padding: 0 10px;
          border: 1px solid var(--line);
          border-radius: 999px;
          color: var(--muted);
          background: #FFFFFF;
          display: inline-flex;
          align-items: center;
          gap: 7px;
          font-size: 12px;
          font-weight: 800;
        }
        .step.active { color: var(--rs); border-color: #F2B8B8; background: #FFF7F7; }
        .step.done { color: var(--teal); border-color: #AAE6D7; background: #F0FCF8; }
        .step-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; }

        .file-summary {
          display: grid;
          grid-template-columns: 46px minmax(0,1fr);
          align-items: center;
          gap: 12px;
          padding: 14px;
          border: 1px solid var(--line);
          border-radius: 8px;
          background: #FBFCFE;
          min-width: 0;
        }
        .pdf-options-card {
          margin: 16px 0 8px 0;
          padding: 16px 18px;
          background: #FFFFFF;
          border: 1.5px solid #E2E8F0;
          border-radius: 14px;
          box-shadow: 0 4px 16px rgba(17, 24, 39, 0.04);
          display: flex;
          flex-direction: column;
          gap: 12px;
        }
        .options-header {
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .options-icon {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          background: #FFF0F1;
          color: var(--rs);
          display: grid;
          place-items: center;
          font-size: 18px;
        }
        .options-title {
          display: block;
          font-size: 13.5px;
          font-weight: 750;
          color: #1E293B;
        }
        .options-subtitle {
          display: block;
          font-size: 11.5px;
          color: #64748B;
          font-weight: 500;
          margin-top: 1px;
        }
        .label-selector-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 12px;
        }
        .label-select-btn {
          display: flex;
          align-items: center;
          gap: 12px;
          padding: 12px 16px;
          border: 2px solid #E2E8F0;
          border-radius: 12px;
          background: #FAFCFF;
          text-align: left;
          cursor: pointer;
          transition: all 0.2s ease-in-out;
        }
        .label-select-btn:hover {
          border-color: #F87171;
          background: #FFF9F9;
        }
        .label-select-btn.active {
          border-color: var(--rs);
          background: #FFF4F4;
          box-shadow: 0 4px 14px rgba(228, 6, 19, 0.12);
        }
        .radio-indicator {
          color: #94A3B8;
          display: grid;
          place-items: center;
          font-size: 22px;
          transition: color 0.2s;
        }
        .label-select-btn.active .radio-indicator {
          color: var(--rs);
        }
        .btn-label-group {
          display: flex;
          flex-direction: column;
          gap: 2px;
        }
        .btn-title {
          font-size: 13.5px;
          font-weight: 750;
          color: #1E293B;
        }
        .label-select-btn.active .btn-title {
          color: var(--rs);
        }
        .btn-example {
          font-size: 11px;
          font-weight: 600;
          color: #64748B;
        }
        .summary-icon {
          width: 46px;
          height: 46px;
          border-radius: 12px;
          background: #EEF7F6;
          color: var(--teal);
          display: grid;
          place-items: center;
          font-size: 26px;
        }
        .file-summary span {
          display: block;
          color: var(--muted);
          font-size: 12px;
          font-weight: 800;
          text-transform: uppercase;
          margin-bottom: 4px;
        }
        .file-summary strong {
          display: block;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          font-size: 15px;
          font-weight: 800;
        }

        .actions { display: flex; gap: 10px; flex-wrap: wrap; }
        .actions.centered { justify-content: center; margin-top: 8px; }
        .btn {
          min-height: 48px;
          border-radius: 12px;
          padding: 0 16px;
          border: 1px solid transparent;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          font-size: 14px;
          font-weight: 800;
          cursor: pointer;
          transition: transform .2s, box-shadow .2s, background .2s, border-color .2s;
          text-decoration: none;
          white-space: nowrap;
        }
        .btn:hover:not(:disabled) { transform: translateY(-2px); }
        .btn:disabled { opacity: .6; cursor: not-allowed; }
        .btn.primary {
          background: var(--rs);
          color: #FFFFFF;
          box-shadow: 0 8px 22px rgba(204,0,0,.22);
        }
        .btn.primary:hover:not(:disabled) { background: var(--rs-dark); }
        .btn.secondary {
          color: var(--rs);
          background: #FFFFFF;
          border-color: #F0B4B4;
        }
        .btn.secondary:hover:not(:disabled) { background: #FFF7F7; }
        .btn.compact { min-height: 40px; padding: 0 13px; }

        .processing-panel {
          min-height: 360px;
          display: grid;
          grid-template-columns: 180px minmax(0,1fr);
          align-items: center;
          gap: 26px;
          padding: 36px;
        }
        .ring { width: 150px; height: 150px; position: relative; }
        .ring svg { width: 100%; height: 100%; transform: rotate(-90deg); }
        .ring circle { transition: stroke-dashoffset .45s ease; }
        .ring strong {
          position: absolute;
          inset: 0;
          display: grid;
          place-items: center;
          color: var(--rs);
          font-size: 30px;
          font-weight: 800;
        }
        .processing-copy span, .done-panel span {
          color: var(--teal);
          font-size: 12px;
          text-transform: uppercase;
          font-weight: 800;
        }
        .processing-copy h2, .done-panel h2 { margin: 8px 0 16px; font-size: 28px; line-height: 1.1; }
        .progress-track {
          height: 10px;
          background: #E7EDF3;
          border-radius: 999px;
          overflow: hidden;
        }
        .progress-track i {
          display: block;
          height: 100%;
          background: var(--rs);
          border-radius: inherit;
          transition: width .4s ease;
        }
        .done-panel {
          min-height: 340px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          text-align: center;
          padding: 34px;
        }
        .done-icon {
          width: 74px;
          height: 74px;
          font-size: 74px;
          color: var(--green);
          margin-bottom: 14px;
        }

        .preview-panel { margin-top: 18px; overflow: hidden; }
        .app-notice {
          margin-top: 14px;
          padding: 13px 15px;
          border: 1px solid #B9DFD5;
          border-radius: 12px;
          color: #076B55;
          background: #F0FCF8;
          font-size: 13px;
          font-weight: 800;
          line-height: 1.45;
        }
        .preview-topbar {
          height: 52px;
          padding: 0 16px;
          border-bottom: 1px solid var(--line);
          background: #FBFCFE;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
        }
        .preview-topbar span {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          color: var(--muted);
          font-size: 13px;
          font-weight: 800;
        }
        .preview-frame { width: 100%; height: 78vh; border: 0; display: block; background: #FFFFFF; }

        .gallery-view { padding: 20px; }
        .gallery-top {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 18px;
          margin-bottom: 16px;
        }
        .gallery-title {
          min-width: 0;
          display: flex;
          align-items: center;
          gap: 12px;
        }
        .gallery-title > span {
          width: 48px;
          height: 48px;
          border-radius: 8px;
          background: #EEF7F6;
          color: var(--teal);
          display: grid;
          place-items: center;
          font-size: 20px;
          font-weight: 800;
        }
        .gallery-title h2 { margin: 0; font-size: 22px; line-height: 1.1; }
        .gallery-title p { margin: 4px 0 0; color: var(--muted); font-size: 13px; font-weight: 700; }
        .gallery-actions {
          display: flex;
          align-items: center;
          justify-content: flex-end;
          gap: 9px;
          flex-wrap: wrap;
        }
        .search-wrap {
          min-width: 260px;
          height: 40px;
          border: 1px solid var(--line);
          border-radius: 8px;
          background: #FFFFFF;
          color: var(--muted);
          display: flex;
          align-items: center;
          gap: 9px;
          padding: 0 12px;
        }
        .search-wrap input {
          width: 100%;
          min-width: 0;
          border: 0;
          outline: 0;
          color: var(--ink);
          background: transparent;
          font-weight: 700;
        }
        .page-size {
          height: 40px;
          border: 1px solid var(--line);
          border-radius: 8px;
          background: #FFFFFF;
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 0 8px 0 12px;
          color: var(--muted);
          font-size: 12px;
          font-weight: 800;
          white-space: nowrap;
        }
        .page-size select {
          height: 28px;
          min-width: 58px;
          border: 1px solid #E2E8F0;
          border-radius: 7px;
          background: #F8FAFC;
          color: var(--ink);
          font-size: 13px;
          font-weight: 800;
          outline: 0;
          cursor: pointer;
        }
        .page-size select:focus { border-color: var(--teal); }
        .gallery-notice {
          margin-bottom: 14px;
          border: 1px solid #BCEBD9;
          background: #F0FCF8;
          color: #076B55;
          border-radius: 8px;
          padding: 11px 12px;
          font-size: 13px;
          font-weight: 800;
        }
        .grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(212px, 1fr));
          gap: 14px;
        }
        .img-card {
          min-width: 0;
          background: #FFFFFF;
          border: 1px solid var(--line);
          border-radius: 8px;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          animation: pop .32s ease both;
          transition: border-color .2s, transform .2s, box-shadow .2s;
        }
        .img-card:hover {
          border-color: #F0B4B4;
          transform: translateY(-3px);
          box-shadow: 0 14px 30px rgba(17,24,39,.08);
        }
        .img-h {
          height: 156px;
          padding: 14px;
          background: #F8FAFC;
          border-bottom: 1px solid #E9EEF3;
          display: grid;
          place-items: center;
        }
        .img-h img {
          max-width: 100%;
          max-height: 100%;
          object-fit: contain;
        }
        .img-meta {
          min-height: 62px;
          padding: 12px;
          display: flex;
          flex-direction: column;
          justify-content: center;
          min-width: 0;
        }
        .img-meta strong {
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          font-size: 13px;
          line-height: 1.3;
        }
        .img-meta span {
          margin-top: 4px;
          color: var(--muted);
          font-size: 12px;
          font-weight: 800;
        }
        .img-tools {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          border-top: 1px solid #E9EEF3;
          background: #FBFCFE;
        }
        .tool-btn {
          height: 38px;
          border: 0;
          border-right: 1px solid #E9EEF3;
          background: transparent;
          color: #5D6978;
          display: grid;
          place-items: center;
          cursor: pointer;
          transition: color .2s, background .2s;
        }
        .tool-btn:last-child { border-right: 0; }
        .tool-btn:hover { background: #FFFFFF; color: var(--rs); }
        .tool-btn.danger:hover { color: var(--danger); background: #FFF5F5; }
        .icon-action {
          width: 34px;
          height: 34px;
          border: 1px solid var(--line);
          border-radius: 8px;
          background: #FFFFFF;
          color: var(--muted);
          display: grid;
          place-items: center;
          cursor: pointer;
          transition: color .2s, border-color .2s, background .2s;
        }
        .icon-action:hover { color: var(--rs); border-color: #F0B4B4; background: #FFF7F7; }
        .icon-action.ok { color: var(--teal); border-color: #AAE6D7; background: #F0FCF8; }
        .rename-row {
          display: grid;
          grid-template-columns: minmax(0,1fr) 34px 34px;
          gap: 7px;
          align-items: center;
        }
        .rename-row input {
          height: 34px;
          min-width: 0;
          border: 1px solid var(--line);
          border-radius: 8px;
          padding: 0 9px;
          outline: 0;
          color: var(--ink);
          font-weight: 800;
        }
        .rename-row input:focus { border-color: var(--teal); }
        .empty-gallery {
          grid-column: 1 / -1;
          min-height: 260px;
          border: 1px dashed #B8C3CF;
          border-radius: 8px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          color: var(--muted);
          background: #F8FAFC;
        }
        .empty-icon { width: 48px; height: 48px; font-size: 48px; color: #A8B3C1; }
        .empty-gallery p { margin: 12px 0 0; font-weight: 800; }
        .pagination-bar {
          margin-top: 16px;
          padding-top: 16px;
          border-top: 1px solid #E9EEF3;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 14px;
          flex-wrap: wrap;
        }
        .page-range {
          color: var(--muted);
          font-size: 13px;
          font-weight: 800;
        }
        .page-range strong { color: var(--ink); }
        .pager-controls {
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .page-count {
          min-width: 104px;
          text-align: center;
          color: var(--ink);
          font-size: 13px;
          font-weight: 800;
        }
        .pager-btn:disabled {
          opacity: .42;
          cursor: not-allowed;
          background: #F8FAFC;
        }
        .pager-btn:disabled:hover {
          color: var(--muted);
          border-color: var(--line);
          background: #F8FAFC;
        }

        .formatter-view {
          display: grid;
          gap: 22px;
        }
        .formatter-grid {
          display: grid;
          grid-template-columns: minmax(0, 1.45fr) minmax(340px, .75fr);
          min-height: 410px;
          overflow: hidden;
          border: 1px solid var(--line);
          border-radius: 22px;
          background: #fff;
          box-shadow: 0 24px 70px rgba(30, 40, 51, .1);
        }
        .formatter-dropzone {
          min-height: 410px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 12px;
          padding: 44px;
          text-align: center;
          cursor: pointer;
          border-right: 1px dashed #CAD7E5;
          background: radial-gradient(circle at 50% 35%, rgba(227,6,19,.075), transparent 40%), #FBFCFE;
          transition: .25s ease;
        }
        .formatter-dropzone:hover, .formatter-dropzone:focus-visible, .formatter-dropzone.dragging {
          outline: none;
          border-color: #E30613;
          background: radial-gradient(circle at 50% 35%, rgba(227,6,19,.12), transparent 42%), #FFF9FA;
        }
        .formatter-dropzone.has-file { background: radial-gradient(circle at 50% 35%, rgba(8,145,98,.1), transparent 42%), #F7FCFA; }
        .formatter-dropzone strong { max-width: 560px; font-size: 21px; color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .formatter-dropzone small { color: var(--muted); }
        .formatter-summary {
          display: flex;
          flex-direction: column;
          align-items: flex-start;
          justify-content: center;
          padding: 38px;
          background: linear-gradient(150deg, #FFFFFF, #F8FAFD);
        }
        .formatter-kicker { color: #E30613; font-size: 12px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; }
        .formatter-summary h2 { margin: 9px 0 11px; font-size: 25px; color: var(--ink); }
        .formatter-summary p { margin: 0; color: var(--muted); font-size: 14px; line-height: 1.65; }
        .formatter-checklist { display: grid; gap: 11px; margin-top: 25px; }
        .formatter-checklist span { display: flex; align-items: center; gap: 8px; color: #344256; font-size: 13px; font-weight: 700; }
        .formatter-checklist .material-symbols-rounded { width: 22px; height: 22px; display: grid; place-items: center; border-radius: 50%; background: #E9F8F1; color: #087A55; font-size: 15px; }
        .formatter-metrics { width: 100%; display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin: 22px 0 16px; }
        .formatter-metrics span { padding: 11px 8px; border: 1px solid #E2E9F1; border-radius: 12px; background: #fff; text-align: center; color: var(--muted); font-size: 10px; }
        .formatter-metrics strong { display: block; color: var(--ink); font-size: 19px; }
        .formatter-groups { width: 100%; margin-bottom: 18px; }
        .formatter-groups small { display: block; margin-bottom: 7px; color: var(--muted); font-weight: 700; }
        .formatter-groups div { display: flex; flex-wrap: wrap; gap: 6px; }
        .formatter-groups span { padding: 6px 9px; border-radius: 999px; background: #FFF0F1; color: #B9000B; font-size: 10px; font-weight: 800; }
        .formatter-download { width: 100%; justify-content: center; margin-bottom: 9px; }
        .formatter-summary > .btn.secondary { width: 100%; justify-content: center; }
        .formatter-preview { overflow: hidden; border: 1px solid var(--line); border-radius: 20px; background: #fff; box-shadow: 0 18px 45px rgba(30,40,51,.08); }
        .formatter-preview-head { display: flex; align-items: center; justify-content: space-between; padding: 18px 20px; border-bottom: 1px solid var(--line); }
        .formatter-preview-head span { color: #E30613; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .08em; }
        .formatter-preview-head h3 { margin: 4px 0 0; color: var(--ink); font-size: 18px; }
        .formatter-preview-head .preview-badge { padding: 7px 10px; border-radius: 999px; background: #EEF4FA; color: #3C5067; letter-spacing: 0; text-transform: none; }
        .formatter-table-scroll { overflow-x: auto; }
        .formatter-table-scroll table { min-width: 1800px; width: 100%; border-collapse: collapse; font-size: 11px; }
        .formatter-table-scroll th { position: sticky; top: 0; max-width: 190px; padding: 11px; background: #FCE7D8; color: #111827; border-right: 1px solid #E7C5AA; text-align: left; }
        .formatter-table-scroll td { max-width: 190px; padding: 10px 11px; color: #344256; border-top: 1px solid #EDF1F5; border-right: 1px solid #EDF1F5; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
        .formatter-table-scroll tr:nth-child(even) td { background: #FAFBFC; }

        .extract-view {
          background: #FFFFFF;
          border: 1px solid var(--line);
          border-radius: 18px;
          padding: 20px;
          box-shadow: 0 12px 34px rgba(17,24,39,.06);
        }
        .extract-grid {
          display: grid;
          grid-template-columns: minmax(0, 1.18fr) minmax(320px, .82fr);
          min-height: 330px;
          border: 1px solid var(--line);
          border-radius: 14px;
          overflow: hidden;
        }
        .extract-dropzone {
          min-width: 0;
          padding: 34px;
          background:
            radial-gradient(circle at 50% 35%, rgba(204,0,0,.055), transparent 40%),
            #FBFCFE;
          border-right: 1px solid var(--line);
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          text-align: center;
          cursor: pointer;
          transition: background .2s, box-shadow .2s;
        }
        .extract-dropzone.dragging {
          background: #FFF5F5;
          box-shadow: inset 0 0 0 2px var(--rs);
        }
        .extract-dropzone.has-file { background: #F7FCFA; }
        .extract-dropzone strong {
          max-width: 100%;
          margin-top: 10px;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          color: var(--ink);
          font-size: 22px;
        }
        .extract-dropzone small {
          margin-top: 7px;
          color: var(--muted);
          font-size: 13px;
          font-weight: 800;
        }
        .extract-summary {
          padding: 30px;
          display: flex;
          flex-direction: column;
          justify-content: center;
          background: #FFFFFF;
        }
        .extract-kicker {
          color: var(--rs);
          font-size: 11px;
          font-weight: 900;
          letter-spacing: .12em;
          text-transform: uppercase;
        }
        .extract-summary h2 { margin: 8px 0 10px; font-size: 24px; }
        .extract-summary p { margin: 0; color: var(--muted); line-height: 1.65; font-size: 14px; }
        .extract-summary code {
          border-radius: 5px;
          padding: 2px 5px;
          color: #7A1212;
          background: #FFF0F0;
          font-size: 12px;
          font-weight: 800;
        }
        .extract-facts {
          margin: 18px 0;
          display: flex;
          gap: 8px;
          flex-wrap: wrap;
        }
        .extract-facts span {
          border: 1px solid #DDE5EC;
          border-radius: 7px;
          padding: 7px 9px;
          background: #F8FAFC;
          color: var(--muted);
          font-size: 11px;
          font-weight: 800;
        }
        .extract-facts strong { color: var(--ink); }
        .extract-notice {
          margin-top: 14px;
          border: 1px solid #BCEBD9;
          border-radius: 8px;
          padding: 11px 13px;
          background: #F0FCF8;
          color: #076B55;
          font-size: 13px;
          font-weight: 800;
        }
        .crop-panel {
          margin-top: 16px;
          overflow: hidden;
          border: 1px solid var(--line);
          border-radius: 10px;
          background: #FFFFFF;
        }
        .crop-toolbar {
          min-height: 72px;
          padding: 12px 14px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 14px;
          border-bottom: 1px solid var(--line);
        }
        .crop-title { display: flex; align-items: center; gap: 10px; }
        .crop-title > .icon { color: var(--rs); }
        .crop-title > div { display: flex; flex-direction: column; gap: 3px; }
        .crop-title strong { font-size: 15px; }
        .crop-title span { color: var(--muted); font-size: 12px; font-weight: 700; }
        .crop-actions, .crop-pages { display: flex; align-items: center; gap: 8px; }
        .crop-pages {
          border: 1px solid var(--line);
          border-radius: 8px;
          padding: 4px;
          background: #F8FAFC;
        }
        .crop-pages span { min-width: 82px; text-align: center; color: var(--muted); font-size: 12px; font-weight: 900; }
        .crop-pages .icon-action { width: 34px; height: 34px; }
        .crop-canvas-wrap {
          min-height: 360px;
          max-height: 72vh;
          overflow: auto;
          padding: 22px;
          display: flex;
          align-items: flex-start;
          justify-content: center;
          background-color: #26313A;
          background-image: linear-gradient(45deg, rgba(255,255,255,.025) 25%, transparent 25%), linear-gradient(-45deg, rgba(255,255,255,.025) 25%, transparent 25%);
          background-size: 20px 20px;
        }
        .crop-stage {
          position: relative;
          isolation: isolate;
          overflow: hidden;
          width: fit-content;
          max-width: 100%;
          line-height: 0;
          cursor: crosshair;
          touch-action: none;
          user-select: none;
          box-shadow: 0 12px 35px rgba(0,0,0,.35);
        }
        .crop-stage img {
          display: block;
          width: auto;
          height: auto;
          max-width: 100%;
          max-height: calc(72vh - 44px);
          object-fit: contain;
          pointer-events: none;
        }
        .crop-empty-hint {
          position: absolute;
          z-index: 2;
          left: 50%;
          top: 50%;
          transform: translate(-50%, -50%);
          display: inline-flex;
          align-items: center;
          gap: 9px;
          padding: 10px 14px;
          border: 1px solid rgba(255,255,255,.88);
          border-radius: 9px;
          color: #FFFFFF;
          background: rgba(15,23,42,.78);
          box-shadow: 0 10px 28px rgba(15,23,42,.24);
          font-size: 13px;
          line-height: 1;
          pointer-events: none;
          white-space: nowrap;
        }
        .crop-empty-hint .icon { font-size: 20px; }
        .crop-selection {
          position: absolute;
          z-index: 2;
          border: 3px solid #FF3B30;
          border-radius: 2px;
          box-shadow: 0 0 0 9999px rgba(5, 12, 18, .58), inset 0 0 0 1px rgba(255,255,255,.75);
          pointer-events: none;
        }
        .crop-selection span {
          position: absolute;
          top: 7px;
          left: 7px;
          padding: 5px 7px;
          border-radius: 5px;
          color: #FFFFFF;
          background: #D71920;
          font-size: 11px;
          font-weight: 900;
          line-height: 1;
          white-space: nowrap;
        }
        .crop-loading {
          min-height: 320px;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 10px;
          color: #FFFFFF;
          font-size: 14px;
          font-weight: 800;
        }
        .json-panel {
          margin-top: 16px;
          border: 1px solid var(--line);
          border-radius: 10px;
          overflow: hidden;
          background: #101820;
        }
        .json-toolbar {
          min-height: 66px;
          padding: 12px 14px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          background: #FFFFFF;
          border-bottom: 1px solid var(--line);
        }
        .json-toolbar > div:first-child { display: flex; flex-direction: column; gap: 3px; }
        .json-toolbar span { color: var(--muted); font-size: 11px; font-weight: 900; text-transform: uppercase; letter-spacing: .08em; }
        .json-toolbar strong { color: var(--ink); font-size: 15px; }
        .json-editor {
          width: 100%;
          min-height: 420px;
          resize: vertical;
          border: 0;
          outline: 0;
          padding: 20px;
          color: #D9F7E8;
          background: #101820;
          font-family: Consolas, 'Courier New', monospace;
          font-size: 13px;
          line-height: 1.65;
          tab-size: 2;
        }

        /* Premium glass and colour-grading layer */
        body {
          font-family: 'Manrope Variable', Manrope, 'Segoe UI Variable', system-ui, sans-serif;
          background:
            radial-gradient(circle at 78% 6%, rgba(204,0,0,.11), transparent 30%),
            radial-gradient(circle at 24% 82%, rgba(15,118,110,.10), transparent 34%),
            linear-gradient(135deg, #F8FAFC 0%, #EEF2F6 44%, #F7F2F3 100%);
          background-attachment: fixed;
        }
        .root { position: relative; display: block; overflow-x: hidden; background: transparent; }
        .top-nav {
          position: sticky;
          top: 0;
          z-index: 50;
          width: 100%;
          padding: 12px 24px;
          border-bottom: 1px solid rgba(255,255,255,.76);
          background: rgba(250,251,252,.76);
          box-shadow: 0 10px 36px rgba(20,29,40,.075);
          backdrop-filter: blur(28px) saturate(160%);
        }
        .top-nav-inner {
          width: min(1320px, 100%);
          min-height: 58px;
          margin: 0 auto;
          display: grid;
          grid-template-columns: minmax(230px,1fr) auto minmax(180px,1fr);
          align-items: center;
          gap: 24px;
        }
        .top-brand { min-width: 0; display: flex; align-items: center; gap: 12px; }
        .top-brand .brand-mark { width: 68px; height: 42px; box-shadow: 0 9px 24px rgba(20,29,40,.12), inset 0 1px 0 #FFFFFF; }
        .top-brand .brand-copy strong { color: #151D26; font-size: 15px; letter-spacing: -.02em; }
        .top-brand .brand-copy span { color: var(--rs); font-size: 9px; font-weight: 900; letter-spacing: .14em; text-transform: uppercase; }
        .top-tabs {
          padding: 5px;
          border: 1px solid rgba(210,218,226,.72);
          border-radius: 15px;
          background: rgba(239,243,247,.68);
          display: flex;
          align-items: center;
          gap: 3px;
          box-shadow: inset 0 1px 0 rgba(255,255,255,.95);
        }
        .top-tabs .tab {
          width: auto;
          min-height: 44px;
          padding: 0 15px;
          border: 1px solid transparent;
          border-radius: 11px;
          color: #667281;
          background: transparent;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          transform: translateZ(0);
          transition: color .25s ease, background-color .25s ease, box-shadow .25s ease, transform .25s ease;
        }
        .top-tabs .tab strong { font-size: 12px; font-weight: 760; letter-spacing: -.01em; white-space: nowrap; }
        .top-tabs .tab .material-symbols-rounded { font-size: 19px; }
        .top-tabs .tab:hover:not(.on) { color: #1E2833; background: rgba(255,255,255,.68); transform: translateY(-1px); }
        .top-tabs .tab.on {
          color: #B00007;
          background: rgba(255,255,255,.92);
          border-color: rgba(255,255,255,.94);
          box-shadow: 0 7px 20px rgba(20,29,40,.09), inset 0 -2px 0 #CC0000;
          transform: none;
        }
        .top-tabs .tab::before { display: none; }
        .top-status { justify-self: end; color: #596574; font-size: 10px; font-weight: 750; display: flex; align-items: center; gap: 9px; white-space: nowrap; }
        .app-shell { position: relative; isolation: isolate; width: 100%; margin-left: 0; }
        .page-header {
          min-height: 210px;
          padding: 44px max(24px, calc((100% - 1240px) / 2)) 34px;
        }
        .page-heading { animation: premiumReveal .55s cubic-bezier(.22,1,.36,1) both; }
        .page-header h1 { font-weight: 720; letter-spacing: -.048em; }
        .workspace-description { font-weight: 520; letter-spacing: -.01em; }
        .header-metrics { animation: premiumReveal .55s .08s cubic-bezier(.22,1,.36,1) both; }
        .main { width: min(1240px, calc(100% - 48px)); padding-top: 28px; }
        .main > section, .main > .app-notice, .main > .preview-panel {
          animation: contentEnter .48s cubic-bezier(.22,1,.36,1) both;
        }
        .quick-guide { margin-bottom: 24px; padding: 14px 18px; border-radius: 16px; }
        .quick-guide strong { font-weight: 700; letter-spacing: -.01em; }
        .tool-panel, .gallery-view, .preview-panel, .extract-view { border-radius: 22px; }
        .tool-panel, .gallery-view, .extract-view { transition: box-shadow .35s ease, transform .35s ease; }
        .dropzone, .extract-dropzone, .img-card, .crop-panel, .json-panel { transform: translateZ(0); }
        .dropzone, .extract-dropzone { transition: border-color .28s ease, background .28s ease, box-shadow .28s ease, transform .28s cubic-bezier(.22,1,.36,1); }
        .btn, .icon-action { transition: transform .22s cubic-bezier(.22,1,.36,1), box-shadow .22s ease, background .22s ease, border-color .22s ease; }
        .btn:active:not(:disabled), .icon-action:active:not(:disabled) { transform: translateY(0) scale(.975); }
        .img-card { transition: transform .32s cubic-bezier(.22,1,.36,1), box-shadow .32s ease, border-color .32s ease; }
        .json-panel, .crop-panel { animation: contentEnter .5s cubic-bezier(.22,1,.36,1) both; }
        @keyframes premiumReveal {
          from { opacity: 0; transform: translateY(14px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes contentEnter {
          from { opacity: 0; transform: translateY(18px) scale(.992); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .app-shell::before,
        .app-shell::after {
          content: '';
          position: fixed;
          z-index: -1;
          border-radius: 50%;
          pointer-events: none;
          filter: blur(2px);
        }
        .app-shell::before {
          width: 430px;
          height: 430px;
          right: -150px;
          top: 150px;
          background: radial-gradient(circle, rgba(215,25,32,.13), rgba(215,25,32,0) 68%);
        }
        .app-shell::after {
          width: 520px;
          height: 520px;
          left: 20%;
          bottom: -240px;
          background: radial-gradient(circle, rgba(15,118,110,.11), rgba(15,118,110,0) 68%);
        }
        .sidebar {
          background:
            radial-gradient(circle at 12% 8%, rgba(215,25,32,.20), transparent 27%),
            linear-gradient(180deg, rgba(19,27,36,.97), rgba(9,15,22,.985));
          box-shadow: 18px 0 54px rgba(8,15,24,.18);
          backdrop-filter: blur(30px) saturate(130%);
        }
        .brand-mark {
          border: 1px solid rgba(255,255,255,.72);
          box-shadow: 0 12px 30px rgba(0,0,0,.24), inset 0 1px 0 rgba(255,255,255,.9);
        }
        .tab { position: relative; overflow: hidden; }
        .tab::before {
          content: '';
          position: absolute;
          inset: 0;
          opacity: 0;
          background: linear-gradient(110deg, transparent 20%, rgba(255,255,255,.12), transparent 78%);
          transform: translateX(-100%);
          transition: transform .55s ease, opacity .25s ease;
        }
        .tab:hover::before { opacity: 1; transform: translateX(100%); }
        .tab.on {
          background: linear-gradient(135deg, #EF3038 0%, #C90008 48%, #8E0006 100%);
          border: 1px solid rgba(255,255,255,.15);
          box-shadow: 0 14px 32px rgba(166,0,7,.34), inset 0 1px 0 rgba(255,255,255,.24);
        }
        .sidebar-support {
          border-color: rgba(255,255,255,.11);
          background: linear-gradient(135deg, rgba(255,255,255,.075), rgba(255,255,255,.025));
          box-shadow: inset 0 1px 0 rgba(255,255,255,.07);
          backdrop-filter: blur(18px);
        }
        .page-header {
          position: relative;
          overflow: hidden;
          background: linear-gradient(135deg, rgba(255,255,255,.88), rgba(255,255,255,.62));
          border-bottom-color: rgba(255,255,255,.86);
          box-shadow: 0 18px 50px rgba(22,31,43,.07), inset 0 -1px 0 rgba(152,164,177,.18);
          backdrop-filter: blur(28px) saturate(145%);
        }
        .page-header::after {
          content: '';
          position: absolute;
          width: 360px;
          height: 360px;
          right: -80px;
          top: -230px;
          border-radius: 50%;
          background: radial-gradient(circle, rgba(215,25,32,.15), transparent 70%);
          pointer-events: none;
        }
        .page-heading, .header-metrics { position: relative; z-index: 1; }
        .header-metrics > div {
          border-color: rgba(255,255,255,.82);
          background: linear-gradient(145deg, rgba(255,255,255,.76), rgba(245,248,251,.48));
          box-shadow: 0 14px 34px rgba(22,31,43,.08), inset 0 1px 0 rgba(255,255,255,.92);
          backdrop-filter: blur(18px);
        }
        .header-metrics > div > .material-symbols-rounded {
          background: linear-gradient(145deg, rgba(255,239,240,.96), rgba(255,255,255,.65));
          box-shadow: inset 0 0 0 1px rgba(204,0,0,.10);
        }
        .quick-guide {
          border-color: rgba(255,255,255,.84);
          background: linear-gradient(120deg, rgba(255,255,255,.72), rgba(255,255,255,.47));
          box-shadow: 0 18px 44px rgba(22,31,43,.07), inset 0 1px 0 rgba(255,255,255,.95);
          backdrop-filter: blur(22px) saturate(140%);
        }
        .quick-guide > div > span {
          background: linear-gradient(145deg, #ED2D35, #A90000);
          box-shadow: 0 7px 18px rgba(204,0,0,.24), inset 0 1px 0 rgba(255,255,255,.3);
        }
        .tool-panel, .gallery-view, .preview-panel, .extract-view {
          border-color: rgba(255,255,255,.86);
          background: linear-gradient(145deg, rgba(255,255,255,.77), rgba(255,255,255,.53));
          box-shadow: 0 28px 70px -28px rgba(20,29,40,.30), inset 0 1px 0 rgba(255,255,255,.96);
          backdrop-filter: blur(26px) saturate(145%);
        }
        .dropzone, .extract-dropzone {
          border-color: rgba(142,157,171,.55);
          background:
            radial-gradient(circle at 50% 20%, rgba(255,255,255,.95), transparent 38%),
            linear-gradient(145deg, rgba(255,255,255,.64), rgba(238,243,247,.48));
          box-shadow: inset 0 1px 0 rgba(255,255,255,.95);
        }
        .dropzone:hover, .dropzone:focus-visible, .dropzone.dragging,
        .extract-dropzone:hover, .extract-dropzone:focus-visible, .extract-dropzone.dragging {
          background: linear-gradient(145deg, rgba(255,255,255,.91), rgba(255,238,239,.72));
          box-shadow: 0 18px 42px rgba(204,0,0,.10), inset 0 1px 0 rgba(255,255,255,1);
        }
        .drop-icon {
          border-color: rgba(255,255,255,.9);
          background: linear-gradient(145deg, rgba(255,255,255,.98), rgba(246,248,251,.74));
          box-shadow: 0 18px 36px rgba(20,29,40,.12), inset 0 1px 0 #FFFFFF;
        }
        .job-panel, .extract-summary {
          border-color: rgba(160,172,184,.30);
          background: linear-gradient(145deg, rgba(255,255,255,.55), rgba(248,250,252,.34));
        }
        .file-summary, .search-wrap, .page-size, .crop-toolbar, .json-toolbar, .pagination-bar {
          border-color: rgba(255,255,255,.86);
          background: rgba(255,255,255,.62);
          box-shadow: inset 0 1px 0 rgba(255,255,255,.94);
          backdrop-filter: blur(16px);
        }
        .btn.primary {
          background: linear-gradient(135deg, #EA2830, #C30007 52%, #990006);
          box-shadow: 0 14px 28px rgba(190,0,7,.25), inset 0 1px 0 rgba(255,255,255,.24);
        }
        .btn.primary:hover:not(:disabled) {
          background: linear-gradient(135deg, #F2353D, #B70006);
          box-shadow: 0 18px 34px rgba(170,0,6,.32), inset 0 1px 0 rgba(255,255,255,.30);
        }
        .btn.secondary, .icon-action {
          background: rgba(255,255,255,.66);
          border-color: rgba(204,0,0,.25);
          box-shadow: inset 0 1px 0 rgba(255,255,255,.9);
          backdrop-filter: blur(14px);
        }
        .img-card {
          border-color: rgba(255,255,255,.86);
          background: linear-gradient(150deg, rgba(255,255,255,.82), rgba(255,255,255,.56));
          box-shadow: 0 18px 42px -24px rgba(18,28,39,.35), inset 0 1px 0 rgba(255,255,255,.95);
          backdrop-filter: blur(18px);
        }
        .img-card:hover {
          transform: translateY(-5px);
          box-shadow: 0 28px 52px -24px rgba(18,28,39,.42), 0 0 0 1px rgba(204,0,0,.10);
        }
        .crop-panel {
          border-color: rgba(255,255,255,.82);
          background: rgba(255,255,255,.58);
          box-shadow: 0 22px 50px rgba(18,28,39,.10), inset 0 1px 0 rgba(255,255,255,.94);
          backdrop-filter: blur(22px);
        }
        .json-panel {
          border-color: rgba(255,255,255,.16);
          background: linear-gradient(145deg, rgba(10,18,26,.98), rgba(22,31,42,.95));
          box-shadow: 0 28px 60px rgba(8,15,23,.22), inset 0 1px 0 rgba(255,255,255,.08);
        }
        .json-editor { background: linear-gradient(145deg, #0D1620, #121E29); }
        .extract-notice, .gallery-notice, .app-notice {
          border-color: rgba(15,159,110,.20);
          background: linear-gradient(120deg, rgba(240,252,248,.86), rgba(255,255,255,.58));
          box-shadow: 0 12px 30px rgba(15,118,110,.07), inset 0 1px 0 rgba(255,255,255,.95);
          backdrop-filter: blur(16px);
        }
        .top-nav {
          background: #FFFFFF;
          border-bottom-color: #E5E9EE;
          box-shadow: 0 6px 24px rgba(20,29,40,.06);
          backdrop-filter: none;
        }
        .page-header {
          min-height: 138px;
          padding-top: 30px;
          padding-bottom: 28px;
          align-items: center;
          background: #FFFFFF;
          border-bottom-color: #E5E9EE;
          box-shadow: none;
          backdrop-filter: none;
        }
        .page-header::after { display: none; }
        .page-header h1 { max-width: 900px; font-size: clamp(28px, 3.2vw, 40px); }
        .top-brand .brand-mark,
        .brand-mark {
          width: 92px;
          height: 44px;
          padding: 0;
          border: 0;
          border-radius: 0;
          background: transparent;
          box-shadow: none;
        }
        .top-brand .logo,
        .logo {
          width: 92px;
          height: 44px;
          object-fit: contain;
        }
        body {
          background: #F5F6F8;
        }
        .app-shell::before,
        .app-shell::after { display: none; }
        .page-header {
          min-height: 106px;
          padding-top: 22px;
          padding-bottom: 22px;
        }
        .page-header h1 {
          font-size: clamp(27px, 3vw, 36px);
          font-weight: 700;
          letter-spacing: -.045em;
        }
        .main {
          width: min(1200px, calc(100% - 48px));
          padding-top: 24px;
        }
        .tool-panel, .gallery-view, .preview-panel, .extract-view {
          border-color: #E3E7EB;
          background: rgba(255,255,255,.94);
          box-shadow: 0 16px 45px rgba(22,31,43,.07);
          backdrop-filter: blur(16px);
        }
        .dropzone, .extract-dropzone {
          background: #FAFBFC;
          box-shadow: none;
        }
        .job-panel, .extract-summary {
          background: #FFFFFF;
        }
        .header-metrics > div,
        .file-summary, .search-wrap, .page-size, .crop-toolbar, .pagination-bar {
          background: #FFFFFF;
          box-shadow: none;
        }
        @media (prefers-reduced-motion: reduce) {
          *, *::before, *::after { scroll-behavior: auto !important; animation-duration: .01ms !important; transition-duration: .01ms !important; }
        }

        .spin { animation: spin .9s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes pop {
          from { opacity: 0; transform: translateY(8px); }
          to { opacity: 1; transform: translateY(0); }
        }

        @media (max-width: 1100px) and (min-width: 861px) {
          .sidebar { width: 240px; }
          .app-shell { width: calc(100% - 240px); margin-left: 240px; }
          .brand-copy strong { font-size: 14px; }
          .page-header { align-items: flex-start; flex-direction: column; }
        }

        @media (max-width: 860px) {
          .root { display: block; }
          .top-nav { padding: 10px 14px; }
          .top-nav-inner { min-height: 52px; grid-template-columns: auto minmax(0,1fr); gap: 12px; }
          .top-status { display: none; }
          .top-tabs { justify-self: end; }
          .top-tabs .tab { min-height: 40px; padding: 0 11px; }
          .sidebar { position: sticky; top: 0; width: 100%; min-height: 0; padding: 10px 14px; flex-direction: row; align-items: center; gap: 12px; border-right: 0; border-bottom: 1px solid #27323D; }
          .sidebar-brand { flex: 0 0 auto; padding: 0; }
          .brand-mark { width: 62px; height: 38px; }
          .brand-copy, .nav-label, .sidebar-support, .sidebar-foot, .nav-arrow { display: none; }
          .sidebar-nav { min-width: 0; width: 100%; flex-direction: row; justify-content: flex-end; gap: 5px; }
          .tab { width: auto; min-height: 46px; grid-template-columns: 30px auto; gap: 5px; padding: 5px 10px; }
          .nav-icon { width: 30px; height: 30px; background: transparent; }
          .tab.on .nav-icon { background: transparent; }
          .nav-copy small { display: none; }
          .nav-copy strong { font-size: 11px; }
          .app-shell { width: 100%; margin-left: 0; }
          .page-header { min-height: 0; padding: 26px 20px 22px; align-items: flex-start; flex-direction: column; gap: 20px; }
          .header-metrics { width: 100%; }
          .header-metrics > div { flex: 1; min-width: 0; }
          .main { width: min(100% - 28px, 1180px); padding-top: 18px; }
          .quick-guide { align-items: stretch; flex-direction: column; padding: 14px; }
          .quick-guide > div { width: 100%; }
          .guide-arrow { display: none; }
          .converter-grid { grid-template-columns: 1fr; min-height: 0; }
          .dropzone { min-height: 280px; }
          .job-panel { border-left: 0; border-top: 1px solid var(--line); padding-left: 0; padding-top: 18px; }
          .processing-panel { grid-template-columns: 1fr; text-align: center; justify-items: center; padding: 28px 18px; }
          .gallery-top { align-items: stretch; flex-direction: column; }
          .extract-grid, .formatter-grid { grid-template-columns: 1fr; }
          .formatter-dropzone { min-height: 300px; border-right: 0; border-bottom: 1px dashed #CAD7E5; }
          .crop-toolbar { align-items: stretch; flex-direction: column; }
          .crop-actions { justify-content: space-between; }
          .extract-dropzone { min-height: 280px; border-right: 0; border-bottom: 1px solid var(--line); }
          .gallery-actions { justify-content: stretch; }
          .search-wrap { width: 100%; min-width: 0; }
          .page-size { flex: 1; justify-content: space-between; }
          .gallery-actions .btn { flex: 1; }
          .pagination-bar { align-items: stretch; flex-direction: column; }
          .pager-controls { justify-content: space-between; }
        }

        @media (max-width: 520px) {
          .top-nav { padding: 8px 10px; }
          .top-nav-inner { gap: 7px; }
          .top-brand .brand-mark { width: 50px; height: 34px; }
          .top-tabs { width: 100%; padding: 3px; border-radius: 12px; }
          .top-tabs .tab { min-width: 0; flex: 1; min-height: 38px; padding: 0 7px; gap: 4px; border-radius: 9px; }
          .top-tabs .tab strong { font-size: 9px; }
          .top-tabs .tab .material-symbols-rounded { font-size: 17px; }
          .sidebar { padding: 8px 10px; }
          .brand-mark { width: 52px; height: 36px; }
          .logo { width: 46px; }
          .tab { min-height: 42px; padding: 4px 7px; grid-template-columns: 28px auto; }
          .nav-icon { width: 28px; height: 28px; }
          .tab .material-symbols-rounded { font-size: 18px; }
          .nav-copy strong { font-size: 10px; }
          .page-header h1 { font-size: 26px; }
          .header-metrics > div { padding: 10px; }
          .header-metrics > div > .material-symbols-rounded { display: none; }
          .tool-panel, .gallery-view, .extract-view { padding: 14px; }
          .dropzone { padding: 22px 14px; }
          .dropzone strong { font-size: 19px; }
          .actions .btn { width: 100%; }
          .grid { grid-template-columns: 1fr; }
          .preview-frame { height: 70vh; }
          .page-size { width: 100%; }
          .pager-controls { width: 100%; }
          .json-toolbar { align-items: stretch; flex-direction: column; }
          .json-toolbar .actions { width: 100%; }
          .json-toolbar .btn { flex: 1; }
          .json-editor { min-height: 360px; }
          .crop-actions { align-items: stretch; flex-direction: column; }
          .crop-pages { justify-content: space-between; }
          .crop-canvas-wrap { padding: 10px; }
        }
      ` }} />
    </div>
  );
}
