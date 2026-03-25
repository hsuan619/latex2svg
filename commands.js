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
        console.error("無法讀取選取文字:", asyncResult.error.message);
        event.completed();
        return;
      }
      
      const latexText = asyncResult.value;
      if (!latexText || latexText.trim() === '') {
        console.warn("未選取任何文字");
        event.completed(); // 必須呼叫 completed，否則 PowerPoint 會卡住該命令
        return;
      }

      try {
        // 利用 MathJax 進行轉換，取得包裹層 <mjx-container>
        const node = html.convert(latexText, { display: true });
        
        // 取得內部的 <svg> 元素並將它轉為字串
        const svgNode = adaptor.firstChild(node);
        const svgString = adaptor.outerHTML(svgNode);
        
        // 將 SVG 轉譯為 PNG 圖片格式的 Base64
        svgToPngBase64(svgString).then(pngBase64 => {
          // 去除 dataURI 前綴，因為 CoercionType.Image 只接受純 Base64 內容
          const cleanBase64 = pngBase64.replace(/^data:image\/(png|jpeg);base64,/, "");

          // 插入圖片來替換目前選取的內容
          Office.context.document.setSelectedDataAsync(cleanBase64, {
            coercionType: Office.CoercionType.Image
          }, function (setResult) {
            if (setResult.status === Office.AsyncResultStatus.Failed) {
              console.error("圖片插入失敗:", setResult.error.message);
            }
            event.completed();
          });
        }).catch(err => {
          console.error("SVG 轉 PNG 失敗:", err);
          event.completed();
        });
      } catch (mathErr) {
        console.error("MathJax 轉換失敗:", mathErr);
        event.completed();
      }
    });
  } catch (err) {
    console.error("執行過程中發生錯誤:", err);
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
