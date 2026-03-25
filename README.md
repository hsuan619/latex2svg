# PowerPoint LaTeX to SVG 增益集 (UI-Less)

這是一個無介面（背景執行）、專為 PowerPoint 設計的 Office Web Add-in。使用者只需在投影片中反白選取 LaTeX 語法（例如 `$E=mc^2$` 或 `\frac{1}{2}mv^2`），點擊上方面板 **[常用]** 裡的轉換按鈕，系統就會在背景使用 MathJax 將其轉換為高畫質公式圖片，並替換原本的文字！

<img width="365" height="155" alt="image" src="https://github.com/user-attachments/assets/6bf4b83b-122a-4053-a079-1b8bbc243968" />
<img width="563" height="192" alt="image" src="https://github.com/user-attachments/assets/c0aa32d2-d897-4310-9e05-a7093cd87f47" />

---

## 👨‍💻 本機開發與測試指南 (Local Development)

### 1. 安裝套件
首先請確保電腦已安裝 [Node.js](https://nodejs.org/)。進入專案資料夾後執行：
```bash
npm install
```

### 2. 安裝本地安全憑證 (必做)
Office 嚴格要求任何增益集都必須跑在安全的 `https` 之上。請透過微軟官方工具安裝本機開發信任憑證：
```bash
npx office-addin-dev-certs install
```
*(執行期間，若出現 Windows 或防毒的安全憑證警告視窗，請務必點擊「是 / 允許」)*

### 3. 一鍵啟動與測試
我們不用手動設定共用資料夾，只需執行微軟官方的開發者一鍵神器：
```bash
npx office-addin-debugging start manifest.xml desktop
```
這行指令會自動：
1. 啟動 `vite` 開發伺服器 (`https://localhost:3000`)。
2. 開啟您的桌面版 PowerPoint。
3. 自動將此增益集掛載進去。

進入簡報後，尋找上方 **[常用]** 標籤最右側的 **「Convert to LaTeX」** 按鈕，選取一段 LaTeX 文字再點擊按鈕，就能立刻看到轉換結果！

---

## 🚀 佈署到 GitHub Pages (免費雲端執行，免開伺服器)

為了不讓您的電腦每次都要開著終端機跑 `localhost`，您可以把專案打包成純靜態網頁，直接發佈到 GitHub Pages 上。

### 步驟一：確保 GitHub Repository 權限
> **⚠️ 重要注意**：免費版 GitHub 帳號的 GitHub Pages 功能**不支援 Private (私有) 專案**。請先到您的 GitHub 專案 `Settings` -> `General` 中，將 `Visibility` 改為 **Public**，否則網頁只會顯示 404 找不到！

### 步驟二：打包與上傳到 GitHub `gh-pages` 分支
請在終端機依序執行：
```bash
# 1. 產生正式打包檔 (dist)
npm run build

# 2. 自動將打包好的網頁上傳到您遠端 repo 的 gh-pages 分支
# （請確保這時候專案已經有 git origin，例如您已做過 git remote add ...）
npx gh-pages -d dist
```

### 步驟三：啟動 GitHub Pages 伺服器
上傳完成後，必須教 GitHub 去啟動您的網頁：
1. 回到 GitHub 專案頁面，點擊 **[Settings]**。
2. 左側選單選擇 **[Pages]**。
3. **Build and deployment** 區塊中：
   - Source 選擇 **Deploy from a branch**
   - Branch 下拉選單選擇剛剛推上來的 **`gh-pages`**
   - 點擊 **[Save]**。
4. 等待約 2 分鐘，點進您的專屬網址確認服務是否已經啟動：  
   👉 `https://您的帳號.github.io/專案名稱/`（此頁面會顯示綠色打勾的服務運作中畫面）。

### 步驟四：把雲端外掛永久裝進 PowerPoint
1. 打開本機的 `manifest.xml` 原始碼。
2. 用「尋找與取代」功能將所有的 `https://localhost:3000` 換成您剛剛獲得的真實雲端網址（例如：`https://your-name.github.io/project`，<ins>注意結尾不要多出斜線</ins>）。
3. 將這份改好網址的 `manifest.xml` 放到電腦任意一個新建立的資料夾 (例如 `D:\SharedAddins`)，並對它按右鍵設為**「共用資料夾」**。
4. 打開 PowerPoint ->「檔案」->「選項」->「信任中心」->「信任中心設定」->「受信任的增益集目錄」。
5. 新增您剛建立資料夾的**網址路徑** (如 `\\您的電腦名稱\SharedAddins`)，**務必勾選「顯示於功能表中」**，按確定並完全關閉重開 PPT。
6. 從「插入」->「我的增益集」->「共用資料夾」找到並安裝即可永久使用！
