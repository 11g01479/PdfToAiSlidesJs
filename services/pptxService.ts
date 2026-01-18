
import pptxgen from "pptxgenjs";
import { AnalysisResult } from "../types";

export const createPresentation = async (data: AnalysisResult): Promise<void> => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';

  // 全スライドをループで作成
  data.slides.forEach((slide) => {
    const s = pres.addSlide();
    
    // 背景/画像の設定
    // スライド全面にPDFの画像を配置します
    if (slide.imageUrl) {
      s.addImage({
        data: slide.imageUrl,
        x: 0,
        y: 0,
        w: "100%",
        h: "100%",
        sizing: { type: 'contain', w: 10, h: 5.625 }
      });
    }

    /**
     * ユーザーのリクエストに基づき、スライド上のテキストオーバーレイ（タイトル等）を削除しました。
     * これにより、スライド上にはPDFの画像のみが表示され、
     * AIが生成した解説は「発表者ノート」セクションにのみ格納されます。
     */

    // AIが生成した解説を「発表者ノート」に追加
    s.addNotes(slide.notes);
  });

  // ファイル名の安全な処理
  const safeName = data.presentationTitle.replace(/[/\\?%*:|"<>]/g, '-').substring(0, 50);
  await pres.writeFile({ fileName: `${safeName || 'presentation'}.pptx` });
};
