import { mathjax } from 'mathjax-full/js/mathjax.js';
import { TeX } from 'mathjax-full/js/input/tex.js';
import { SVG } from 'mathjax-full/js/output/svg.js';
import { browserAdaptor } from 'mathjax-full/js/adaptors/browserAdaptor.js';
import { RegisterHTMLHandler } from 'mathjax-full/js/handlers/html.js';
import { AllPackages } from 'mathjax-full/js/input/tex/AllPackages.js';

// 使用 browserAdaptor 以便在真瀏覽器環境下建立 DOM 節點
const adaptor = browserAdaptor();
RegisterHTMLHandler(adaptor);
const tex = new TeX({ packages: AllPackages });
const svg = new SVG({ fontCache: 'local' });
const html = mathjax.document(document, { InputJax: tex, OutputJax: svg });

// 當 Office 準備完成後，將按鈕的 Action 名稱與 JavaScript 函數綁定
Office.onReady(() => {
  Office.actions.associate("convertLatexToSvg", convertLatexToSvg);
});

/**
 * 此函數負責讀取 PowerPoint 中選取的文字，利用 MathJax 轉換為公式，
 * 最後轉繪為高品質 PNG 圖片插入回投影片中。
 * @param {Office.AddinCommands.Event} event - 由 Office 傳入的事件物件
 */
async function convertLatexToSvg(event) {
  try {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        Office.context.document.setSelectedDataAsync("[讀取文字失敗: " + asyncResult.error.message + "]", {coercionType: Office.CoercionType.Text});
        event.completed();
        return;
      }
      
      const latexText = asyncResult.value;
      if (!latexText || latexText.trim() === '') {
        Office.context.document.setSelectedDataAsync("[請先用滑鼠『反白選取』一段文字哦！]", {coercionType: Office.CoercionType.Text});
        event.completed();
        return;
      }

      try {
        const node = html.convert(latexText, { display: true });
        const svgNode = adaptor.firstChild(node);
        const svgString = adaptor.outerHTML(svgNode);
        
        svgToPngBase64(svgString).then(pngBase64 => {
          const cleanBase64 = pngBase64.replace(/^data:image\/(png|jpeg);base64,/, "");

          Office.context.document.setSelectedDataAsync(cleanBase64, {
            coercionType: Office.CoercionType.Image
          }, function (setResult) {
            if (setResult.status === Office.AsyncResultStatus.Failed) {
              Office.context.document.setSelectedDataAsync("[圖片插入失敗: " + setResult.error.message + "]", {coercionType: Office.CoercionType.Text});
            }
            event.completed();
          });
        }).catch(err => {
          Office.context.document.setSelectedDataAsync("[SVG轉圖片失敗!]", {coercionType: Office.CoercionType.Text});
          event.completed();
        });
      } catch (mathErr) {
        Office.context.document.setSelectedDataAsync("[MathJax語法錯誤!]", {coercionType: Office.CoercionType.Text});
        event.completed();
      }
    });
  } catch (err) {
    Office.context.document.setSelectedDataAsync("[未知致命錯誤!]", {coercionType: Office.CoercionType.Text});
    event.completed();
  }
}

/**
 * 輔助函數：將 SVG HTML 字串透過 Blob 與 HTML5 Canvas 轉換為 PNG 之 Base64 DataURI
 * 此方法可以避開 PowerPoint 直接插入原生 SVG 可能遇到的相容性問題。
 * @param {string} svgText - 含有 <svg>...</svg> 的字串
 * @returns {Promise<string>} 回傳 PNG 格式的 Base64 圖片
 */
function svgToPngBase64(svgText) {
  return new Promise((resolve, reject) => {
    // 確保 SVG 有設定 xml_ns 命名空間，否則 Image.src 解析可能會報錯
    let processedSvg = svgText;
    if (!processedSvg.includes('xmlns=')) {
      processedSvg = processedSvg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
    }

    const svgBlob = new Blob([processedSvg], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(svgBlob);
    
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      // 使用 3 倍解析度確保放進簡報時維持清晰銳利
      const pixelRatio = 3;
      canvas.width = img.width * pixelRatio;
      canvas.height = img.height * pixelRatio;
      const ctx = canvas.getContext("2d");
      
      // 若有需要可以給透明底，一般公式背景是透明的
      ctx.scale(pixelRatio, pixelRatio);
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(url);
      
      resolve(canvas.toDataURL("image/png"));
    };
    img.onerror = (e) => reject(e);
    img.src = url;
  });
}
