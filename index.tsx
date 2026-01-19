
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
    // UI Initial State
    step1!.classList.add('step-inactive');
    loadingArea!.classList.remove('hidden');
    startBtn!.classList.add('hidden');
    errorModal!.classList.add('hidden');

    // 1. PDF Rendering (Sequential to ensure order)
    updateProgress(10, "PDFの各ページを画像として展開中...");
    const arrayBuffer = await pdfFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    images = [];
    
    for (let i = 1; i <= pdf.numPages; i++) {
      updateProgress(10 + (i / pdf.numPages * 30), `ページを処理中... (${i}/${pdf.numPages})`);
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
    updateProgress(50, "AIがドキュメントを詳細に解析中...");
    const base64 = await fileToBase64(pdfFile);
    const rawAnalysis = await fetchAiAnalysis(base64, pdf.numPages);

    // 3. Robust Mapping and Sorting
    updateProgress(90, "解析結果を整理中...");
    
    // AIが1始まりのインデックスを返した場合の自動補正
    let aiSlides = rawAnalysis.slides || [];
    const minIdx = Math.min(...aiSlides.map((s: any) => s.pageIndex));
    if (minIdx === 1) {
      aiSlides = aiSlides.map((s: any) => ({ ...s, pageIndex: s.pageIndex - 1 }));
    }

    const finalSlides = [];
    for (let i = 0; i < pdf.numPages; i++) {
      // インデックスが完全に一致するスライドを探す
      const aiSlide = aiSlides.find((s: any) => s.pageIndex === i);
      finalSlides.push({
        pageIndex: i,
        title: aiSlide?.title || `ページ ${i + 1}`,
        notes: aiSlide?.notes || "このページの解説を生成できませんでした。内容を確認してください。",
        imageUrl: images[i]
      });
    }

    analysis = {
      ...rawAnalysis,
      slides: finalSlides.sort((a, b) => a.pageIndex - b.pageIndex)
    };

    // 4. Show Results
    renderResults();
    updateProgress(100, "完了しました！");
  } catch (err: any) {
    console.error(err);
    showError("解析中にエラーが発生しました: " + err.message);
  }
}

async function fetchAiAnalysis(base64Pdf: string, pageCount: number) {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  // ズレを防止するための極めて厳格なプロンプト
  const prompt = `
このPDFドキュメントをページごとに詳細に分析してください。
【必須条件】
1. PDFの全 ${pageCount} ページを、1ページも欠かすことなく解析してください。
2. 出力の 'slides' 配列には必ず ${pageCount} 個の要素を含めてください。
3. 各要素の 'pageIndex' は必ず 0（1ページ目）から順に ${pageCount - 1} まで連番で割り当ててください。
4. ページを統合したり、スキップしたりしないでください。
5. 各ページに対して、そのページの内容に基づいたタイトルと、発表者がそのまま読み上げられるような詳細な「スピーカーノート」を作成してください。
`;

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
                pageIndex: { type: Type.INTEGER, description: "0-based index of the page" },
                title: { type: Type.STRING },
                notes: { type: Type.STRING, description: "Detailed speaker notes in Japanese" }
              },
              required: ["pageIndex", "title", "notes"]
            }
          }
        },
        required: ["presentationTitle", "summary", "slides"]
      }
    }
  });

  const text = response.text;
  if (!text) throw new Error("AIから応答がありませんでした。");
  return JSON.parse(text);
}

function renderResults() {
  loadingArea!.classList.add('hidden');
  step2!.classList.remove('step-inactive');
  badge2!.classList.replace('bg-slate-700', 'bg-blue-500');
  badge2!.classList.replace('text-slate-400', 'text-white');
  badge2!.innerHTML = '<svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7"></path></svg>';
  
  resultsPlaceholder!.classList.add('hidden');
  resultsContent!.classList.remove('hidden');

  resTitle!.textContent = analysis.presentationTitle;
  resSummary!.textContent = analysis.summary;

  slidesGrid!.innerHTML = analysis.slides.map((slide: any, idx: number) => `
    <div class="bg-slate-900/40 p-5 rounded-3xl border border-slate-700/50 flex flex-col gap-4 animate-in fade-in slide-in-from-bottom-4 duration-500" style="animation-delay: ${idx * 50}ms">
      <div class="aspect-video bg-black rounded-2xl overflow-hidden border border-slate-800 shadow-lg relative group">
        <img src="${slide.imageUrl}" alt="" class="w-full h-full object-contain" />
        <div class="absolute top-3 left-3 bg-black/80 backdrop-blur-md text-white text-[10px] px-3 py-1 rounded-full font-black tracking-widest border border-white/10 shadow-2xl">
          PAGE ${idx + 1}
        </div>
      </div>
      <div class="space-y-3">
        <h5 class="font-bold text-slate-100 truncate text-sm px-1 border-l-2 border-cyan-500 ml-1">${slide.title}</h5>
        <div class="text-xs text-slate-400 bg-slate-950/50 p-4 rounded-xl border border-slate-800 leading-relaxed min-h-[80px]">
          ${slide.notes.replace(/\n/g, '<br>')}
        </div>
      </div>
    </div>
  `).join('');
}

async function handleDownload() {
  try {
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';

    // 必ずインデックス順にソートされていることを保証して追加
    const sortedSlides = [...analysis.slides].sort((a, b) => a.pageIndex - b.pageIndex);

    sortedSlides.forEach((slide: any) => {
      const s = pres.addSlide();
      if (slide.imageUrl) {
        s.addImage({
          data: slide.imageUrl,
          x: 0, y: 0, w: "100%", h: "100%",
          sizing: { type: 'contain', w: 10, h: 5.625 }
        });
      }
      // 解説をスピーカーノートに挿入
      s.addNotes(slide.notes);
    });

    const safeName = analysis.presentationTitle.replace(/[/\\?%*:|"<>]/g, '-').substring(0, 50);
    await pres.writeFile({ fileName: `${safeName || 'presentation'}.pptx` });
  } catch (err: any) {
    showError("パワーポイントの生成に失敗しました: " + err.message);
  }
}

function updateProgress(percent: number, message: string) {
  if (progressBar) progressBar.style.width = `${percent}%`;
  if (loadingMsg) loadingMsg.textContent = message;
}

function showError(msg: string) {
  errorMsg!.textContent = msg;
  errorModal!.classList.remove('hidden');
  loadingArea!.classList.add('hidden');
}

async function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
