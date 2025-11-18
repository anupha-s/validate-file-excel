const fileInput = document.getElementById("fileInput");
    const statusEl = document.getElementById("status");
    const sheetTabs = document.getElementById("sheetTabs");
    const sheetContents = document.getElementById("sheetContents");
    const downloadBtn = document.getElementById("downloadBtn");
    const progressContainer = document.getElementById("progressContainer");
    const progressBar = document.getElementById("progressBar");
    const debugConsole = document.getElementById("debugConsole");

    let workbook = null;
    let processedData = [];
    let pendingCount = 0;
    let totalCount = 0;
    let sheetStats = {};

    function log(message, type = 'info') {
      const line = document.createElement('div');
      line.className = `debug-line debug-${type}`;
      const timestamp = new Date().toLocaleTimeString('th-TH');
      line.textContent = `[${timestamp}] ${message}`;
      debugConsole.appendChild(line);
      debugConsole.scrollTop = debugConsole.scrollHeight;
      debugConsole.classList.add('show');
      console.log(message);
    }

    function updateProgress() {
      if (totalCount === 0) return;
      const processed = totalCount - pendingCount;
      const percent = Math.round((processed / totalCount) * 100);
      progressBar.style.width = percent + '%';
      progressBar.textContent = `${processed}/${totalCount} (${percent}%)`;
    }

    function extractDriveFileId(url) {
      if (!url) return null;
      let urlStr = String(url).trim();
      
      if (urlStr.startsWith('="') && urlStr.endsWith('"')) {
        urlStr = urlStr.slice(2, -1);
      } else if (urlStr.startsWith("='") && urlStr.endsWith("'")) {
        urlStr = urlStr.slice(2, -1);
      }
      
      if (!urlStr.includes("drive.google.com")) return null;
      
      const patterns = [
        /\/d\/([a-zA-Z0-9_-]+)/,
        /[?&]id=([a-zA-Z0-9_-]+)/,
        /\/file\/d\/([a-zA-Z0-9_-]+)/
      ];
      
      for (const pattern of patterns) {
        const match = urlStr.match(pattern);
        if (match) return match[1];
      }
      
      return null;
    }

    function getImageUrl(fileId) {
      return `https://drive.google.com/thumbnail?id=${fileId}&sz=w400`;
    }

    function loadImageDimensions(fileId, callback) {
      const img = new Image();
      img.crossOrigin = "anonymous";
      
      const urls = [
        `https://drive.google.com/thumbnail?id=${fileId}&sz=w2000`,
        `https://drive.google.com/uc?export=view&id=${fileId}`,
        `https://lh3.googleusercontent.com/d/${fileId}`
      ];
      
      let urlIndex = 0;
      
      function tryNextUrl() {
        if (urlIndex >= urls.length) {
          callback(null, null, "โหลดไม่ได้");
          return;
        }
        
        img.src = urls[urlIndex];
        urlIndex++;
      }
      
      img.onload = () => {
        callback(img.naturalWidth, img.naturalHeight, null);
      };
      
      img.onerror = () => {
        tryNextUrl();
      };
      
      tryNextUrl();
    }

    function updateStatus() {
      if (!workbook) {
        statusEl.textContent = "ยังไม่ได้เลือกไฟล์";
        statusEl.style.borderLeftColor = "#3b82f6";
        return;
      }
      if (pendingCount > 0) {
        statusEl.textContent = `⏳ กำลังประมวลผลรูปภาพ... เหลืออีก ${pendingCount}/${totalCount} รูป`;
        statusEl.style.borderLeftColor = "#f59e0b";
        updateProgress();
      } else {
        statusEl.textContent = "✅ ประมวลผลเสร็จสมบูรณ์! คลิกปุ่มด้านบนเพื่อดาวน์โหลดไฟล์ Excel";
        statusEl.style.borderLeftColor = "#10b981";
        downloadBtn.disabled = false;
        progressBar.style.width = '100%';
        progressBar.textContent = '100%';
        log('ประมวลผลเสร็จสมบูรณ์!', 'success');
      }
    }

    function updateSheetInfo(sheetName) {
      const stats = sheetStats[sheetName];
      const infoEl = document.getElementById(`info-${sheetName}`);
      if (infoEl) {
        infoEl.innerHTML = `
          รูปทั้งหมด: <strong>${stats.total}</strong> | 
          ประมวลผลแล้ว: <strong>${stats.processed}</strong> | 
          <span style="color: #16a34a;">จัตุรัส: ${stats.square}</span> | 
          <span style="color: #dc2626;">สี่เหลี่ยมผืนผ้า: ${stats.rectangle}</span>
          ${stats.error > 0 ? `| <span style="color: #f59e0b;">Error: ${stats.error}</span>` : ''}
        `;
      }
    }

    fileInput.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      try {
        debugConsole.innerHTML = '';
        debugConsole.classList.add('show');
        progressContainer.style.display = 'block';
        progressBar.style.width = '0%';
        progressBar.textContent = '0%';

        log(`เริ่มอ่านไฟล์: ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`, 'info');
        statusEl.textContent = "กำลังอ่านไฟล์ Excel...";
        statusEl.style.borderLeftColor = "#3b82f6";

        const arrayBuffer = await file.arrayBuffer();
        log('อ่านไฟล์สำเร็จ กำลัง parse ด้วย ExcelJS...', 'info');
        
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        
        log(`พบ ${workbook.worksheets.length} sheets`, 'success');
      } catch (error) {
        log(`Error: ${error.message}`, 'error');
        statusEl.textContent = `❌ เกิดข้อผิดพลาด: ${error.message}`;
        statusEl.style.borderLeftColor = "#ef4444";
        return;
      }

      sheetTabs.innerHTML = "";
      sheetContents.innerHTML = "";
      processedData = [];
      pendingCount = 0;
      totalCount = 0;
      downloadBtn.disabled = true;
      sheetStats = {};
      
      let totalImages = 0;
      
      log('กำลังสร้าง UI สำหรับแต่ละ sheet...', 'info');

      // สร้าง tabs และ content
      workbook.worksheets.forEach((worksheet, index) => {
        const sheetName = worksheet.name;

        // สร้าง tab
        const tabBtn = document.createElement("button");
        tabBtn.className = "tab" + (index === 0 ? " active" : "");
        tabBtn.textContent = sheetName;
        tabBtn.onclick = () => switchTab(sheetName);
        sheetTabs.appendChild(tabBtn);

        // สร้าง content
        const contentDiv = document.createElement("div");
        contentDiv.className = "sheet-content" + (index === 0 ? " active" : "");
        contentDiv.id = `sheet-${sheetName}`;
        contentDiv.innerHTML = `
          <div class="sheet-info">
            📊 Sheet: <strong>${sheetName}</strong> | 
            <span id="info-${sheetName}">กำลังโหลดข้อมูล...</span>
          </div>
          <div class="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Row</th>
                  <th>SKU</th>
                  <th>Product Name</th>
                  <th>Link</th>
                  <th>Preview</th>
                  <th>ขนาด</th>
                  <th>จัตุรัส?</th>
                </tr>
              </thead>
              <tbody id="tbody-${sheetName}"></tbody>
            </table>
          </div>
        `;
        sheetContents.appendChild(contentDiv);

        sheetStats[sheetName] = { total: 0, processed: 0, square: 0, rectangle: 0, error: 0 };
      });

      window.switchTab = function(sheetName) {
        document.querySelectorAll(".tab").forEach(tab => {
          tab.classList.remove("active");
          if (tab.textContent === sheetName) {
            tab.classList.add("active");
          }
        });

        document.querySelectorAll(".sheet-content").forEach(content => {
          content.classList.remove("active");
        });
        document.getElementById(`sheet-${sheetName}`).classList.add("active");
      };

      // ประมวลผลแต่ละ sheet
      workbook.worksheets.forEach((worksheet) => {
        const sheetName = worksheet.name;
        log(`กำลังประมวลผล sheet: ${sheetName}`, 'info');

        // เพิ่ม header ในคอลัมน์ I (แถวที่ 2)
        const headerCell = worksheet.getCell('I2');
        if (!headerCell.value) {
          headerCell.value = "รูปจัตุรัส 1:1 ไหม (Y/N)";
          headerCell.font = { bold: true };
          headerCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
          };
        }

        let rowsInSheet = 0;
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber <= 2) return;

          // ExcelJS: ต้องใช้ .text หรือ .result สำหรับ formula
          const cell6 = row.getCell(6);
          let colF = null;
          
          // ลองหลายวิธีในการอ่านค่า
          if (cell6.formula) {
            // ถ้าเป็น formula ให้ใช้ result
            colF = cell6.result || cell6.value;
          } else if (cell6.value && typeof cell6.value === 'object' && cell6.value.text) {
            // ถ้าเป็น rich text
            colF = cell6.value.text;
          } else {
            // ค่าปกติ
            colF = cell6.value;
          }
          
          if (!colF) {
            log(`Row ${rowNumber}: คอลัมน์ F ว่างเปล่า`, 'warning');
            return;
          }
          
          const colFStr = String(colF);
          log(`Row ${rowNumber}: Col F = ${colFStr.substring(0, 60)}...`, 'info');
          
          if (!colFStr.includes("drive.google.com")) {
            log(`Row ${rowNumber}: ไม่ใช่ Google Drive URL`, 'warning');
            return;
          }

          const fileId = extractDriveFileId(colF);
          if (!fileId) {
            log(`Row ${rowNumber}: ไม่สามารถดึง file ID จาก URL: ${colFStr.substring(0, 50)}...`, 'warning');
            return;
          }
          
          rowsInSheet++;
          log(`Row ${rowNumber}: พบรูป file ID = ${fileId}`, 'success');

          totalImages++;
          totalCount++;
          sheetStats[sheetName].total++;

          const colA = row.getCell(1).value || "";
          const colB = row.getCell(2).value || "";

          const rowData = {
            sheetName,
            rowNumber,
            sku: colA,
            name: colB,
            fileId,
            result: null,
            worksheet: worksheet
          };
          processedData.push(rowData);

          const thumbnailUrl = getImageUrl(fileId);
          const viewUrl = `https://drive.google.com/file/d/${fileId}/view`;

          const tableBody = document.getElementById(`tbody-${sheetName}`);
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td>${rowNumber}</td>
            <td style="font-size:0.75rem;">${colA}</td>
            <td style="font-size:0.75rem;">${colB}</td>
            <td><a href="${viewUrl}" target="_blank">เปิด</a></td>
            <td><img src="${thumbnailUrl}" alt="preview"></td>
            <td class="size" style="font-size:0.75rem; color:#6b7280;"><span class="loading">กำลังตรวจสอบ...</span></td>
            <td class="result">
              <span class="badge badge-pending">รอ...</span>
            </td>
          `;
          tableBody.appendChild(tr);

          const resultCell = tr.querySelector(".result");
          const sizeCell = tr.querySelector(".size");

          pendingCount++;
          updateStatus();

          loadImageDimensions(fileId, (width, height, error) => {
            if (error || !width || !height) {
              sizeCell.textContent = "โหลดไม่ได้";
              resultCell.innerHTML = `<span class="badge badge-error">ERROR</span>`;
              rowData.result = "ERROR";
              sheetStats[sheetName].error++;
              log(`Row ${rowNumber}: โหลดรูปไม่ได้ (${error})`, 'error');
            } else {
              const isSquare = width === height;
              const flag = isSquare ? "Y" : "N";

              sizeCell.textContent = `${width}×${height}`;
              resultCell.innerHTML = `
                <span class="badge ${isSquare ? "badge-y" : "badge-n"}">${flag}</span>
              `;

              rowData.result = flag;
              
              if (isSquare) {
                sheetStats[sheetName].square++;
              } else {
                sheetStats[sheetName].rectangle++;
              }

              // เขียนค่าลง Excel
              const cell = rowData.worksheet.getCell(`I${rowNumber}`);
              cell.value = flag;
            }

            sheetStats[sheetName].processed++;
            updateSheetInfo(sheetName);
            
            pendingCount--;
            updateStatus();
          });
        });
        
        log(`Sheet ${sheetName}: พบ ${rowsInSheet} รูป`, 'success');
      });

      if (totalImages === 0) {
        statusEl.textContent = "⚠️ ไม่พบรูปภาพที่ต้องประมวลผล (ตรวจสอบว่าคอลัมน์ F มี Google Drive URL)";
        statusEl.style.borderLeftColor = "#f59e0b";
        log('ไม่พบรูปภาพที่ต้องประมวลผล', 'warning');
        progressContainer.style.display = 'none';
      } else {
        statusEl.textContent = `🔍 พบรูปภาพทั้งหมด ${totalImages} รูป กำลังโหลดและตรวจสอบ...`;
        log(`เริ่มโหลดรูปภาพทั้งหมด ${totalImages} รูป`, 'info');
        updateProgress();
      }
    });

    downloadBtn.addEventListener("click", async () => {
      if (!workbook) return;

      try {
        log('กำลังสร้างไฟล์ Excel...', 'info');
        statusEl.textContent = "กำลังสร้างไฟล์ Excel...";
        statusEl.style.borderLeftColor = "#3b82f6";

        const buffer = await workbook.xlsx.writeBuffer();
        log(`สร้างไฟล์สำเร็จ (${(buffer.byteLength / 1024 / 1024).toFixed(2)} MB)`, 'success');
        
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "Makro_Image_with_rectangle_flag.xlsx";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        statusEl.textContent = "✅ ดาวน์โหลดไฟล์สำเร็จ! ตรวจสอบคอลัมน์ I ในไฟล์ที่ดาวน์โหลด";
        statusEl.style.borderLeftColor = "#10b981";
        log('ดาวน์โหลดไฟล์สำเร็จ!', 'success');
      } catch (error) {
        log(`Error: ${error.message}`, 'error');
        statusEl.textContent = `❌ เกิดข้อผิดพลาด: ${error.message}`;
        statusEl.style.borderLeftColor = "#ef4444";
      }
    });
