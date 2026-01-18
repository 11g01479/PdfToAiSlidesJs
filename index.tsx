
import { GoogleGenAI, Type } from "@google/genai";
import * as pdfjsLib from 'pdfjs-dist';
import pptxgen from "pptxgenjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://esm.sh/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs`;

// State
let pdfFile: File | null = null;
let analysis: any = null;
let images: string[] = [];

// Elements
const pdfInput = document.getElementById('pdf-input') as HTMLInputElement;
const fileNameLabel = document.getElementById('file-name');
const startBtn = document.getElementById('start-btn');
const loadingArea = document.getElementById('loading-area');
const loadingMsg = document.getElementById('loading-msg');
const progressBar = document.getElementById('progress-bar');
const step1 = document.getElementById('step1');
const step2 = document.getElementById('step2');
const badge2 = document.getElementById('badge2');
const resultsContent = document.getElementById('results-content');
const resultsPlaceholder = document.getElementById('results-placeholder');
const slidesGrid = document.getElementById('slides-grid');
const resTitle = document.getElementById('res-title');
const resSummary = document.getElementById('res-summary');
const downloadBtn = document.getElementById('download-btn');
const errorModal = document.getElementById('error-modal');
const errorMsg = document.getElementById('error-msg');

// Event Listeners
pdfInput.addEventListener('change', (e: any) => {
  const file = e.target.files?.[0];
  if (file && file.type === 'application/pdf') {
    pdfFile = file;
    fileNameLabel!.textContent = file.name;
    startBtn!.classList.remove('hidden');
  }
});

startBtn?.addEventListener('click', startAnalysis);
downloadBtn?.addEventListener('click', handleDownload);

async function startAnalysis() {
  if (!pdfFile) return;

  try {
    // UI Update
    step1!.classList.add('step-inactive');
    loadingArea!.classList.remove('hidden');
    startBtn!.classList.add('hidden');

    // 1. PDF Rendering
    updateProgress(10, "PDFを展開中...");
    const arrayBuffer = await pdfFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    images = [];
    
    for (let i = 1; i <= pdf.numPages; i++) {
      updateProgress(10 + (i / pdf.numPages * 30), `PDFを展開中... (${i}/${pdf.numPages} ページ)`);
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d')!;
      canvas.height = viewport.height;
      canvas.width = viewport.width;
      await page.render({ canvasContext: ctx, viewport }).promise;
      images.push(canvas.toDataURL('image/jpeg', 0.85));
    }

    // 2. AI Analysis
    updateProgress(50, "AIがドキュメントを解析中...");
    const base64 = await fileToBase64(pdfFile);
    analysis = await fetchAiAnalysis(base64, pdf.numPages);

    // 3. Mapping Images
    const finalSlides = [];
    for (let i = 0; i < pdf.numPages; i++) {
      const aiSlide = analysis.slides.find((s: any) => s.pageIndex === i);
      finalSlides.push({
        pageIndex: i,
        title: aiSlide?.title || `ページ ${i + 1}`,
        notes: aiSlide?.notes || "解説を生成できませんでした。",
        imageUrl: images[i]
      });
    }
    analysis.slides = finalSlides;

    // 4. Show Results
    renderResults();
  } catch (err: any) {
    showError(err.message);
  }
}

async function fetchAiAnalysis(base64Pdf: string, pageCount: number) {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  const prompt = `このPDFドキュメントをページごとに詳細に分析してください。全 ${pageCount} ページを1枚ずつのスライドとして構成してください。各ページに対して、その内容に基づいたタイトルと詳細なスピーカーノートを作成してください。ページ番号を一切飛ばさず全て含めてください。`;

  const response = await ai.models.generateContent({
    model: 'gemini-3-flash-preview',
    contents: {
      parts: [
        { inlineData: { mimeType: "application/pdf", data: base64Pdf } },
        { text: prompt }
      ]
    },
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          presentationTitle: { type: Type.STRING },
          summary: { type: Type.STRING },
          slides: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                pageIndex: { type: Type.INTEGER },
                title: { type: Type.STRING },
                notes: { type: Type.STRING }
              },
              required: ["pageIndex", "title", "notes"]
            }
          }
        },
        required: ["presentationTitle", "summary", "slides"]
      }
    }
  });

  return JSON.parse(response.text || '{}');
}

function renderResults() {
  loadingArea!.classList.add('hidden');
  step2!.classList.remove('step-inactive');
  badge2!.classList.replace('bg-slate-700', 'bg-blue-500');
  badge2!.classList.replace('text-slate-400', 'text-white');
  badge2!.textContent = '✓';
  resultsPlaceholder!.classList.add('hidden');
  resultsContent!.classList.remove('hidden');

  resTitle!.textContent = analysis.presentationTitle;
  resSummary!.textContent = analysis.summary;

  slidesGrid!.innerHTML = analysis.slides.map((slide: any, idx: number) => `
    <div class="bg-slate-900/40 p-5 rounded-3xl border border-slate-700/50 flex flex-col gap-4">
      <div class="aspect-video bg-black rounded-2xl overflow-hidden border border-slate-800 shadow-lg relative">
        <img src="${slide.imageUrl}" alt="" class="w-full h-full object-contain" />
        <div class="absolute top-3 left-3 bg-black/70 backdrop-blur-md text-white text-[10px] px-2 py-1 rounded-lg font-black tracking-widest border border-white/10">PAGE ${idx + 1}</div>
      </div>
      <div class="space-y-2">
        <h5 class="font-bold text-slate-100 truncate text-sm">${slide.title}</h5>
        <div class="text-xs text-slate-400 line-clamp-3 bg-slate-950/50 p-4 rounded-xl border border-slate-800 leading-relaxed italic">
          ${slide.notes}
        </div>
      </div>
    </div>
  `).join('');
}

async function handleDownload() {
  try {
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';

    analysis.slides.forEach((slide: any) => {
      const s = pres.addSlide();
      if (slide.imageUrl) {
        s.addImage({
          data: slide.imageUrl,
          x: 0, y: 0, w: "100%", h: "100%",
          sizing: { type: 'contain', w: 10, h: 5.625 }
        });
      }
      s.addNotes(slide.notes);
    });

    const safeName = analysis.presentationTitle.replace(/[/\\?%*:|"<>]/g, '-').substring(0, 50);
    await pres.writeFile({ fileName: `${safeName || 'presentation'}.pptx` });
  } catch (err: any) {
    showError("PPTX生成に失敗しました: " + err.message);
  }
}

function updateProgress(percent: number, message: string) {
  progressBar!.style.width = `${percent}%`;
  loadingMsg!.textContent = message;
}

function showError(msg: string) {
  errorMsg!.textContent = msg;
  errorModal!.classList.remove('hidden');
}

async function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.readAsDataURL(file);
  });
}
