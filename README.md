

# PubMed 疾病相關文章查詢工具


## **程式簡介**
PubMed 疾病相關文章查詢工具是一款基於 Python 開發的桌面應用程式，旨在幫助研究人員快速查詢與特定基因和疾病相關的文章。透過結合 PubMed 的 API，該工具能夠自動化地執行多基因、多疾病條件的批量查詢，並將結果存儲為 Excel 文件，便於後續分析。

本工具適用於需要大量文獻查詢的科研工作者，特別是基因研究與疾病關聯分析的場景。

---

## **程式功能**
1. **基因列表導入**：支持從 Excel 文件中導入基因列表，要求文件中包含名為 `Gene` 或 `gene` 的列。
2. **多疾病條件查詢**：允許用戶自定義疾病條件及其別名，並基於這些條件查詢相關的文章。
3. **PubMed 批量查詢**：自動執行基因與疾病條件的組合查詢，並返回相關文章的 PubMed ID（PMID）。
4. **結果保存**：將查詢結果（包括文章數量和部分 PMID）保存為 Excel 文件。
5. **進度顯示**：實時顯示查詢進度，方便用戶掌握執行情況。
6. **測試模式**：提供測試模式，允許用戶在查詢前限制處理的基因數量（例如僅處理前 10 行）。

---

## **使用說明**

### **1. 啟動程式**
雙擊執行打包後的 `.exe` 文件（例如 `search.exe`），程式界面將自動打開。

---

### **2. 輸入基因列表**
1. 點擊 **「選擇基因列表文件」** 按鈕，選擇包含基因名稱的 Excel 文件。
2. 文件格式要求：
   - 文件必須為 `.xlsx` 格式。
   - 文件中必須包含一列名為 `Gene` 或 `gene` 的欄位，欄位中每一行為一個基因名稱。

---

### **3. 定義疾病條件**
1. 在程式的文本框中輸入疾病條件及其別名，格式如下：
   ```
   疾病名稱: 別名1, 別名2, 別名3
   ```
   例如：
   ```
   DKD: DKD, diabetic kidney disease
   hypertension: hypertension
   hyperlipidemia: Hyperlipidemia, Hypercholesterolemia, Hypertriglyceridemia, Dyslipidemia
   CKD: CKD, Chronic Kidney Disease
   ```
2. 每一行代表一個疾病及其別名，別名之間用逗號 `,` 分隔。

---

### **4. 啟用測試模式（可選）**
如果希望僅測試少量基因的查詢，可以：
1. 勾選 **「啟用測試模式」**。
2. 在旁邊的輸入框中填寫需要處理的基因數量（默認為 10）。

---

### **5. 開始查詢**
1. 點擊 **「開始查詢」** 按鈕，程式將開始執行基因與疾病條件的組合查詢。
2. 查詢過程中：
   - 進度條會顯示當前查詢進度。
   - 下方的進度信息框會實時更新當前查詢的基因與疾病條件。

---

### **6. 保存查詢結果**
1. 查詢完成後，程式會提示用戶選擇保存結果的文件位置。
2. 結果將保存為 Excel 文件，文件中包含以下欄位：
   - **基因名稱**。
   - **每個疾病條件的相關文章數量**。
   - **每個疾病條件的部分 PubMed ID（最多顯示前 5 個）**。

---

## **注意事項**
1. **網路連接**：該工具依賴 PubMed API，使用時需確保設備連接到網路。
2. **API 查詢限制**：PubMed API 對每分鐘的查詢次數有限制，若查詢過多可能會導致暫時性封鎖。
3. **輸入格式**：基因列表文件和疾病條件的格式必須符合要求，否則可能導致程式報錯。
4. **文件保存**：保存結果時請選擇有效的文件路徑，避免因文件名或路徑問題導致保存失敗。

---

## **常見問題**

### **1. 為什麼程式無法讀取基因列表文件？**
- 請確認文件是否為 `.xlsx` 格式。
- 確保文件中包含名為 `Gene` 或 `gene` 的欄位。

### **2. 為什麼查詢過程中出現「無法連接到 PubMed」的錯誤？**
- 請檢查設備的網路連接是否正常。
- 如果網路正常，可能是 PubMed API 請求過多，請稍等幾分鐘後重試。

### **3. 為什麼結果保存失敗？**
- 請確認選擇的保存路徑有效，並且文件名不包含非法字符（例如 `\ / : * ? " < > |`）。

---

## **版本資訊**
- **版本**：1.0
- **開發者**：陳稟靝
- **發布日期**：2025 年 2 月

---

如果需要更多幫助或遇到其他問題，請聯繫開發者！

--------------------------------------------------

# PubMed Disease-Related Article Query Tool

## **Program Overview**
The PubMed Disease-Related Article Query Tool is a desktop application developed in Python, designed to help researchers quickly search for articles related to specific genes and diseases. By integrating PubMed's API, this tool automates batch queries for multiple genes and disease conditions and stores the results in an Excel file for further analysis.

This tool is ideal for researchers who need to perform large-scale literature searches, particularly in the context of gene research and disease association studies.

---

## **Program Features**
1. **Gene List Import**: Supports importing a gene list from an Excel file, requiring the file to contain a column named `Gene` or `gene`.
2. **Multi-Disease Condition Queries**: Allows users to define disease conditions and their aliases, and queries articles based on these conditions.
3. **Batch PubMed Queries**: Automatically performs combination queries of genes and disease conditions, returning relevant articles' PubMed IDs (PMIDs).
4. **Result Saving**: Saves query results (including article counts and partial PMIDs) to an Excel file.
5. **Progress Display**: Displays real-time query progress, helping users monitor the execution status.
6. **Test Mode**: Provides a test mode that allows users to limit the number of genes processed before the query (e.g., only processing the first 10 rows).

---

## **Instructions for Use**

### **1. Start the Program**
Double-click the packaged `.exe` file (e.g., `search.exe`), and the program interface will open automatically.

---

### **2. Input Gene List**
1. Click the **"Select Gene List File"** button to choose an Excel file containing gene names.
2. File format requirements:
   - The file must be in `.xlsx` format.
   - The file must contain a column named `Gene` or `gene`, with each row under this column representing a gene name.

---

### **3. Define Disease Conditions**
1. Enter disease conditions and their aliases in the program's text box in the following format:
   ```
   Disease Name: Alias1, Alias2, Alias3
   ```
   For example:
   ```
   DKD: DKD, diabetic kidney disease
   hypertension: hypertension
   hyperlipidemia: Hyperlipidemia, Hypercholesterolemia, Hypertriglyceridemia, Dyslipidemia
   CKD: CKD, Chronic Kidney Disease
   ```
2. Each line represents one disease and its aliases, with aliases separated by commas `,`.

---

### **4. Enable Test Mode (Optional)**
If you want to test the query with a limited number of genes:
1. Check the **"Enable Test Mode"** option.
2. Enter the number of genes to process in the adjacent input box (default is 10).

---

### **5. Start the Query**
1. Click the **"Start Query"** button to begin the combination queries of genes and disease conditions.
2. During the query process:
   - The progress bar will display the current query progress.
   - The progress information box below will update in real-time with the current gene and disease condition being queried.

---

### **6. Save Query Results**
1. After the query is complete, the program will prompt you to select a location to save the results.
2. The results will be saved in an Excel file containing the following columns:
   - **Gene Name**.
   - **The number of relevant articles for each disease condition**.
   - **Partial PubMed IDs for each disease condition (up to the first 5)**.

---

## **Notes**
1. **Network Connection**: This tool relies on the PubMed API, so ensure that your device is connected to the internet while using it.
2. **API Query Limits**: The PubMed API imposes limits on the number of queries per minute. Excessive queries may result in temporary blocking.
3. **Input Format**: The gene list file and disease condition format must meet the requirements; otherwise, errors may occur.
4. **File Saving**: When saving results, choose a valid file path to avoid failures due to invalid file names or paths.

---

## **Frequently Asked Questions**

### **1. Why can't the program read the gene list file?**
- Ensure the file is in `.xlsx` format.
- Make sure the file contains a column named `Gene` or `gene`.

### **2. Why does the query process show a "Cannot connect to PubMed" error?**
- Check if your device's network connection is stable.
- If the network is stable, it may be due to exceeding PubMed API's query limits. Wait a few minutes and try again.

### **3. Why does saving the results fail?**
- Ensure the selected save path is valid, and the file name does not contain invalid characters (e.g., `\ / : * ? " < > |`).

---

## **Version Information**
- **Version**: 1.0
- **Developer**: Ben
- **Release Date**: February 2025

---

If you need further assistance or encounter other issues, please contact the developer!