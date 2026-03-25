# PowerPoint LaTeX to SVG 增益集 (UI-Less)

這是一個無介面（背景執行）、專為 PowerPoint 設計的 Office Web Add-in。使用者只需在投影片中反白選取 LaTeX 語法（例如 `$E=mc^2$`），點擊上方面板的轉換按鈕，系統就會偷偷在背景使用 MathJax 將其轉換為高畫質公式圖片，並替換原本的文字！

## 👨‍💻 本機開發與測試指南 (Local Development)

### 1. 安裝套件
首先請確保電腦已安裝 [Node.js](https://nodejs.org/)。進入專案資料夾後執行：
```bash
npm install
```

### 2. 安裝本地安全憑證 (必做)
Office 嚴格要求任何增益集都必須跑在 `https` 之上。請透過微軟官方工具安裝本機開發憑證：
```bash
npx office-addin-dev-certs install
```
*(出現 UAC 安全性警告視窗時請按「是」)*

### 3. 一鍵啟動與測試
我們不用在那邊設定惱人的共用資料夾，微軟官方提供了開發者神器，直接執行：
```bash
npx office-addin-debugging start manifest.xml desktop
```
這行指令會自動：
1. 在背景跑起 `vite` 伺服器 (`https://localhost:3000`)。
2. 開啟您的桌面版 PowerPoint。
3. 自動將此增益集掛載進去。

進入 PPT 後，尋找上方 **[常用]** 標籤最右側的 **「Convert to LaTeX」** 按鈕即可測試！

---

## 🚀 佈署到雲端 (部署給自己永久使用)

為了不讓電腦每次都要開著終端機跑 `localhost`，您可以把這個專案打包成純靜態網頁，並丟到 GitHub Pages 等免費空間託管。

### 步驟一：打包靜態檔案
```bash
npm run build
```
指令完成後，專案中會多出一個 `dist` 資料夾。這裡面包含了所有被壓縮與優化過的核心代碼 (`commands.html` 等)。

### 步驟二：上傳到靜態網頁伺服器
將 `dist` 資料夾的內容上傳到任何支援 HTTPS 的靜態網頁空間。
推薦使用 GitHub Pages、Vercel 或是 Netlify。您會獲得一個專屬網址，例如：`https://your-name.github.io/latex-app/`。

### 步驟三：修改 Manifest 網址
1. 打開專案中的 `manifest.xml`。
2. 搜尋裡面所有的 `https://localhost:3000`。
3. 把搜尋到的結果**全部取代**為您剛剛得到的雲端網址 (記得結尾不要留斜線 `/`)。
4. 儲存檔案。

### 步驟四：把外掛永久裝進 PowerPoint
1. 在電腦上建立一個資料夾 (例如 `D:\SharedAddins`)，對它按右鍵將其設定為**「共用資料夾」**。
2. 把修改好的 `manifest.xml` 複製放進去。
3. 開啟 PowerPoint -> 點選左下角「選項」->「信任中心」->「信任中心設定」->「受信任的增益集目錄」。
4. 貼上該資料夾的**網路路徑** (如 `\\您的電腦名稱\SharedAddins`) 並點擊「新增目錄」。
5. **務必勾選「顯示於功能表中」**，按確定並完全關閉 PPT。
6. 重新打開 PPT，點擊「插入」->「我的增益集」->「共用資料夾」，您就能把這套永久版的增益集加進去了！
