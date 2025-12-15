/**
 * 輪休表生成器 - 主要應用邏輯
 */

// ===== Global State =====
const appState = {
  templateWorkbook: null,
  templateBuffer: null,
  rulesData: null,
  sheetNames: [],
  selectedSheets: new Set(),
  sheetConfigs: {},
  targetYear: new Date().getFullYear(),
  targetMonth: new Date().getMonth() + 1,
  generatedBlob: null,
};

// ===== Month Names (Chinese) =====
const monthNamesChinese = {
  一: 1,
  二: 2,
  三: 3,
  四: 4,
  五: 5,
  六: 6,
  七: 7,
  八: 8,
  九: 9,
  十: 10,
  十一: 11,
  十二: 12,
};

// ===== Initialize App =====
document.addEventListener("DOMContentLoaded", () => {
  initializeApp();
});

function initializeApp() {
  // Set up year selector
  const yearSelect = document.getElementById("targetYear");
  const currentYear = new Date().getFullYear();
  for (let y = currentYear - 1; y <= currentYear + 5; y++) {
    const option = document.createElement("option");
    option.value = y;
    option.textContent = `${y} 年`;
    if (y === currentYear) option.selected = true;
    yearSelect.appendChild(option);
  }

  // File upload handlers
  document
    .getElementById("templateFile")
    .addEventListener("change", handleTemplateUpload);
  document
    .getElementById("rulesFile")
    .addEventListener("change", handleRulesUpload);

  // Year/Month change handlers
  document
    .getElementById("targetYear")
    .addEventListener("change", updateMonthInfo);
  document
    .getElementById("targetMonth")
    .addEventListener("change", updateMonthInfo);

  // Select all checkbox
  document
    .getElementById("selectAllSheets")
    .addEventListener("change", handleSelectAll);

  // Generate button
  document
    .getElementById("generateBtn")
    .addEventListener("click", generateExcel);
}

// ===== File Upload Handlers =====
async function handleTemplateUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  try {
    document.getElementById("templateFileName").textContent = file.name;
    document.getElementById("templateUpload").classList.add("uploaded");

    // Store the buffer for later use
    appState.templateBuffer = await file.arrayBuffer();

    // Read workbook to get sheet names
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(appState.templateBuffer);
    appState.templateWorkbook = workbook;

    // Get sheet names
    appState.sheetNames = workbook.worksheets.map((ws) => ws.name);

    checkStep1Complete();
  } catch (error) {
    console.error("Error loading template:", error);
    alert("無法讀取 Excel 範本檔案，請確認格式正確。");
  }
}

async function handleRulesUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  try {
    document.getElementById("rulesFileName").textContent = file.name;
    document.getElementById("rulesUpload").classList.add("uploaded");

    const text = await file.text();
    const result = Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
    });

    appState.rulesData = result.data;
    displayRulesPreview(result.data);

    checkStep1Complete();
  } catch (error) {
    console.error("Error loading rules:", error);
    alert("無法讀取規則 CSV 檔案，請確認格式正確。");
  }
}

function displayRulesPreview(data) {
  const preview = document.getElementById("rulesPreview");
  const table = document.getElementById("rulesTable");

  if (!data || data.length === 0) return;

  const headers = Object.keys(data[0]);
  let html = "<thead><tr>";
  headers.forEach((h) => (html += `<th>${h}</th>`));
  html += "</tr></thead><tbody>";

  data.forEach((row) => {
    if (row["月份"] === "合計") return; // Skip total row
    html += "<tr>";
    headers.forEach((h) => (html += `<td>${row[h] || ""}</td>`));
    html += "</tr>";
  });
  html += "</tbody>";

  table.innerHTML = html;
  preview.classList.remove("hidden");
}

function checkStep1Complete() {
  const hasTemplate = appState.templateBuffer !== null;
  const hasRules = appState.rulesData !== null;

  document.getElementById("step1Next").disabled = !(hasTemplate && hasRules);

  if (hasTemplate && hasRules) {
    document.getElementById("step1Next").onclick = () => goToStep(2);
  }
}

// ===== Step Navigation =====
function goToStep(stepNumber) {
  // Update progress nav
  document.querySelectorAll(".progress-nav .step").forEach((step, index) => {
    step.classList.remove("active", "completed");
    if (index + 1 < stepNumber) {
      step.classList.add("completed");
    } else if (index + 1 === stepNumber) {
      step.classList.add("active");
    }
  });

  // Update content
  document.querySelectorAll(".step-content").forEach((content, index) => {
    content.classList.remove("active");
    if (index + 1 === stepNumber) {
      content.classList.add("active");
    }
  });

  // Step-specific initialization
  if (stepNumber === 2) {
    updateMonthInfo();
  } else if (stepNumber === 3) {
    initializeSheetSelection();
  } else if (stepNumber === 4) {
    displaySummary();
  }
}

// ===== Step 2: Month Info =====
function updateMonthInfo() {
  appState.targetYear = parseInt(document.getElementById("targetYear").value);
  appState.targetMonth = parseInt(document.getElementById("targetMonth").value);

  if (!appState.rulesData) return;

  // Find matching month in rules
  const monthData = appState.rulesData.find((row) => {
    const monthNum = monthNamesChinese[row["月份"]];
    return monthNum === appState.targetMonth;
  });

  if (monthData) {
    document.getElementById("totalLeaveDays").textContent =
      monthData["可休總天數"] || "-";
    document.getElementById("monthNotes").textContent =
      monthData["備註說明"] || "無特殊備註";
  } else {
    document.getElementById("totalLeaveDays").textContent = "-";
    document.getElementById("monthNotes").textContent = "-";
  }
}

// ===== Step 3: Sheet Selection =====
function initializeSheetSelection() {
  const grid = document.getElementById("sheetsGrid");
  grid.innerHTML = "";

  appState.sheetNames.forEach((name) => {
    const item = document.createElement("div");
    item.className = "sheet-item";
    item.innerHTML = `
            <label class="checkbox-wrapper">
                <input type="checkbox" data-sheet="${name}" ${
      appState.selectedSheets.has(name) ? "checked" : ""
    }>
                <span class="checkmark"></span>
                <span>${name}</span>
            </label>
        `;

    const checkbox = item.querySelector("input");
    checkbox.addEventListener("change", () =>
      handleSheetSelection(name, checkbox.checked)
    );

    grid.appendChild(item);
  });

  updateSheetDetails();
}

function handleSelectAll(event) {
  const isChecked = event.target.checked;
  const checkboxes = document.querySelectorAll(
    '.sheets-grid input[type="checkbox"]'
  );

  checkboxes.forEach((cb) => {
    cb.checked = isChecked;
    const sheetName = cb.dataset.sheet;
    if (isChecked) {
      appState.selectedSheets.add(sheetName);
      if (!appState.sheetConfigs[sheetName]) {
        appState.sheetConfigs[sheetName] = createDefaultConfig();
      }
    } else {
      appState.selectedSheets.delete(sheetName);
    }
  });

  updateSheetDetails();
}

function handleSheetSelection(name, isSelected) {
  const item = document
    .querySelector(`input[data-sheet="${name}"]`)
    .closest(".sheet-item");

  if (isSelected) {
    appState.selectedSheets.add(name);
    item.classList.add("selected");
    if (!appState.sheetConfigs[name]) {
      appState.sheetConfigs[name] = createDefaultConfig();
    }
  } else {
    appState.selectedSheets.delete(name);
    item.classList.remove("selected");
  }

  updateSheetDetails();
}

function createDefaultConfig() {
  return {
    staffList: [""],
    departmentManager: "",
    siteManager: "",
    creator: "",
  };
}

function updateSheetDetails() {
  const detailsSection = document.getElementById("sheetDetails");
  const tabs = document.getElementById("sheetTabs");
  const forms = document.getElementById("sheetForms");

  const step3NextBtn = document.getElementById("step3Next");
  step3NextBtn.disabled = appState.selectedSheets.size === 0;

  // Add onclick handler for step3Next
  if (appState.selectedSheets.size > 0) {
    step3NextBtn.onclick = () => goToStep(4);
  }

  if (appState.selectedSheets.size === 0) {
    detailsSection.classList.add("hidden");
    return;
  }

  detailsSection.classList.remove("hidden");
  tabs.innerHTML = "";
  forms.innerHTML = "";

  let isFirst = true;
  appState.selectedSheets.forEach((name) => {
    // Create tab
    const tab = document.createElement("div");
    tab.className = `sheet-tab ${isFirst ? "active" : ""}`;
    tab.textContent = name;
    tab.onclick = () => switchSheetTab(name);
    tabs.appendChild(tab);

    // Create form
    const form = createSheetForm(name, isFirst);
    forms.appendChild(form);

    isFirst = false;
  });
}

function switchSheetTab(name) {
  document.querySelectorAll(".sheet-tab").forEach((tab) => {
    tab.classList.toggle("active", tab.textContent === name);
  });
  document.querySelectorAll(".sheet-form").forEach((form) => {
    form.classList.toggle("active", form.dataset.sheet === name);
  });
}

function createSheetForm(sheetName, isActive) {
  const config = appState.sheetConfigs[sheetName];

  const form = document.createElement("div");
  form.className = `sheet-form ${isActive ? "active" : ""}`;
  form.dataset.sheet = sheetName;

  form.innerHTML = `
        <div class="form-group">
            <label>人員名單</label>
            <div class="staff-list" data-sheet="${sheetName}">
                ${config.staffList
                  .map(
                    (staff, idx) => `
                    <div class="staff-item">
                        <input type="text" class="form-control staff-input" value="${staff}" 
                               data-index="${idx}" placeholder="輸入人員姓名">
                        <button class="btn-remove" onclick="removeStaff('${sheetName}', ${idx})">✕</button>
                    </div>
                `
                  )
                  .join("")}
            </div>
            <button class="btn-add" onclick="addStaff('${sheetName}')">+ 新增人員</button>
        </div>
        <div class="form-group">
            <label>處主管</label>
            <input type="text" class="form-control config-input" data-field="departmentManager" 
                   value="${
                     config.departmentManager
                   }" placeholder="輸入處主管姓名">
        </div>
        <div class="form-group">
            <label>工地主管</label>
            <input type="text" class="form-control config-input" data-field="siteManager" 
                   value="${config.siteManager}" placeholder="輸入工地主管姓名">
        </div>
        <div class="form-group">
            <label>製表人</label>
            <input type="text" class="form-control config-input" data-field="creator" 
                   value="${config.creator}" placeholder="輸入製表人姓名">
        </div>
    `;

  // Add event listeners for config inputs
  form.querySelectorAll(".config-input").forEach((input) => {
    input.addEventListener("input", () => {
      appState.sheetConfigs[sheetName][input.dataset.field] = input.value;
    });
  });

  // Add event listeners for staff inputs
  form.querySelectorAll(".staff-input").forEach((input) => {
    input.addEventListener("input", () => {
      const idx = parseInt(input.dataset.index);
      appState.sheetConfigs[sheetName].staffList[idx] = input.value;
    });
  });

  return form;
}

function addStaff(sheetName) {
  appState.sheetConfigs[sheetName].staffList.push("");
  refreshStaffList(sheetName);
}

function removeStaff(sheetName, index) {
  appState.sheetConfigs[sheetName].staffList.splice(index, 1);
  if (appState.sheetConfigs[sheetName].staffList.length === 0) {
    appState.sheetConfigs[sheetName].staffList.push("");
  }
  refreshStaffList(sheetName);
}

function refreshStaffList(sheetName) {
  const staffList = document.querySelector(
    `.staff-list[data-sheet="${sheetName}"]`
  );
  const config = appState.sheetConfigs[sheetName];

  staffList.innerHTML = config.staffList
    .map(
      (staff, idx) => `
        <div class="staff-item">
            <input type="text" class="form-control staff-input" value="${staff}" 
                   data-index="${idx}" placeholder="輸入人員姓名">
            <button class="btn-remove" onclick="removeStaff('${sheetName}', ${idx})">✕</button>
        </div>
    `
    )
    .join("");

  staffList.querySelectorAll(".staff-input").forEach((input) => {
    input.addEventListener("input", () => {
      const idx = parseInt(input.dataset.index);
      appState.sheetConfigs[sheetName].staffList[idx] = input.value;
    });
  });
}

// ===== Step 4: Summary & Generate =====
function displaySummary() {
  const grid = document.getElementById("summaryGrid");
  const selectedSheetsList = Array.from(appState.selectedSheets).join(", ");

  const monthData = appState.rulesData.find((row) => {
    return monthNamesChinese[row["月份"]] === appState.targetMonth;
  });

  grid.innerHTML = `
        <div class="summary-item">
            <span class="summary-label">目標年月</span>
            <span class="summary-value">${appState.targetYear} 年 ${
    appState.targetMonth
  } 月</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">選取場域</span>
            <span class="summary-value">${selectedSheetsList}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">可休總天數</span>
            <span class="summary-value">${
              monthData?.["可休總天數"] || "-"
            }</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">備註說明</span>
            <span class="summary-value">${
              monthData?.["備註說明"] || "無"
            }</span>
        </div>
    `;
}

// ===== Excel Generation =====
async function generateExcel() {
  const generateBtn = document.getElementById("generateBtn");
  const progressContainer = document.getElementById("progressContainer");
  const progressFill = document.getElementById("progressFill");
  const progressText = document.getElementById("progressText");
  const downloadSection = document.getElementById("downloadSection");

  generateBtn.disabled = true;
  progressContainer.classList.remove("hidden");

  try {
    // Create a fresh workbook from the original buffer
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(appState.templateBuffer);

    // Get month info
    const monthData = appState.rulesData.find((row) => {
      return monthNamesChinese[row["月份"]] === appState.targetMonth;
    });

    const totalLeaveDays = monthData?.["可休總天數"] || "";
    const notes = monthData?.["備註說明"] || "";

    // Parse holiday information from notes
    const holidays = parseHolidayNotes(notes, appState.targetMonth);

    // Get days in month
    const daysInMonth = new Date(
      appState.targetYear,
      appState.targetMonth,
      0
    ).getDate();

    // Process each selected sheet
    const selectedArray = Array.from(appState.selectedSheets);
    const firstSheetName = selectedArray[0]; // First sheet name for formula reference

    for (let i = 0; i < selectedArray.length; i++) {
      const sheetName = selectedArray[i];
      const worksheet = workbook.getWorksheet(sheetName);

      if (worksheet) {
        progressText.textContent = `處理中: ${sheetName}`;
        progressFill.style.width = `${((i + 1) / selectedArray.length) * 80}%`;

        await processWorksheet(worksheet, {
          year: appState.targetYear,
          month: appState.targetMonth,
          daysInMonth,
          totalLeaveDays,
          holidays,
          config: appState.sheetConfigs[sheetName],
          isFirstSheet: i === 0,
          firstSheetName: firstSheetName,
        });
      }
    }

    progressText.textContent = "生成檔案中...";
    progressFill.style.width = "90%";

    // Generate blob
    const buffer = await workbook.xlsx.writeBuffer();
    appState.generatedBlob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    progressFill.style.width = "100%";
    progressText.textContent = "完成!";

    // Show download section
    setTimeout(() => {
      progressContainer.classList.add("hidden");
      downloadSection.classList.remove("hidden");

      document.getElementById("downloadBtn").onclick = () => {
        const fileName = `工程人員輪休表 (${appState.targetYear}.${appState.targetMonth}).xlsx`;
        saveAs(appState.generatedBlob, fileName);
      };
    }, 500);
  } catch (error) {
    console.error("Error generating Excel:", error);
    alert("生成 Excel 時發生錯誤：" + error.message);
    generateBtn.disabled = false;
    progressContainer.classList.add("hidden");
  }
}

// ===== Holiday Notes Parser =====
function parseHolidayNotes(notes, month) {
  const holidays = {};
  console.log("Parsing holiday notes for month:", month);
  console.log("Notes input:", notes);

  if (!notes) {
    console.log("No notes provided");
    return holidays;
  }

  // Split by semicolon or newline
  const parts = notes
    .split(/[;\n]/)
    .map((p) => p.trim())
    .filter((p) => p);

  console.log("Parts after split:", parts);

  parts.forEach((part) => {
    // Pattern 1: 春節2/15(小年夜)~2/19(初三)
    const rangeMatch = part.match(
      /(.+?)(\d+)\/(\d+)\((.+?)\)~(\d+)\/(\d+)\((.+?)\)/
    );
    if (rangeMatch) {
      const [
        ,
        holidayName,
        startMonth,
        startDay,
        startNote,
        endMonth,
        endDay,
        endNote,
      ] = rangeMatch;
      const sMonth = parseInt(startMonth);
      const eMonth = parseInt(endMonth);
      const sDay = parseInt(startDay);
      const eDay = parseInt(endDay);

      if (sMonth === month || eMonth === month) {
        for (let d = sDay; d <= eDay; d++) {
          if (d === sDay) {
            holidays[d] = `${holidayName.trim()}(${startNote})`;
          } else if (d === eDay) {
            holidays[d] = `${holidayName.trim()}(${endNote})`;
          } else {
            holidays[d] = holidayName.trim();
          }
        }
      }
      return;
    }

    // Pattern 2: 2/20補休(2/15小年夜週日) or 2/27彈性補假(2/28和平紀念日週六)
    const singleMatch = part.match(/(\d+)\/(\d+)(.+?)(?:\(|（)/);
    if (singleMatch) {
      const [, m, d, name] = singleMatch;
      if (parseInt(m) === month) {
        holidays[parseInt(d)] = name.trim();
      }
      return;
    }

    // Pattern 3: 1/1元旦 or 5/1勞動節
    const simpleMatch = part.match(/(\d+)\/(\d+)(.+)/);
    if (simpleMatch) {
      const [, m, d, name] = simpleMatch;
      if (parseInt(m) === month) {
        holidays[parseInt(d)] = name.trim();
        console.log(`Pattern 3 match: ${part} -> day ${d}: ${name.trim()}`);
      }
    }
  });

  console.log("Final parsed holidays:", holidays);
  return holidays;
}

// ===== Worksheet Processing =====
async function processWorksheet(worksheet, options) {
  const {
    year,
    month,
    daysInMonth,
    totalLeaveDays,
    holidays,
    config,
    isFirstSheet,
    firstSheetName,
  } = options;

  console.log("Processing worksheet with options:", {
    year,
    month,
    daysInMonth,
    totalLeaveDays,
    holidays,
    isFirstSheet,
    firstSheetName,
  });
  console.log("Config:", config);

  // Analyze worksheet structure
  const structure = analyzeWorksheetStructure(worksheet);
  console.log("Worksheet structure:", structure);

  // 1. Update year/month in the title area (更新右上角年月資訊)
  updateYearMonth(worksheet, year, month, isFirstSheet, firstSheetName);

  // 2. Update date row and apply conditional formatting
  if (structure.dateRow && structure.dateStartCol) {
    updateDateRowWithFormatting(worksheet, structure, year, month, daysInMonth);
  }

  // 3. Update weekday row if exists
  if (structure.weekdayRow && structure.dateStartCol) {
    updateWeekdayRow(worksheet, structure, year, month, daysInMonth);
  }

  // 4. Clear and fill staff names (清除舊人員並填入新人員)
  clearAndFillStaffNames(worksheet, structure, config);

  // 5. Clear and update lunar row (農曆欄位)
  if (structure.lunarRow && structure.dateStartCol) {
    updateLunarRowContent(worksheet, structure, daysInMonth, holidays);
  }

  // 6. Update total leave days (合計欄位)
  updateTotalLeaveDays(worksheet, structure, totalLeaveDays);
}

// Analyze the worksheet structure to find key rows and columns
function analyzeWorksheetStructure(worksheet) {
  const structure = {
    dateRow: null,
    weekdayRow: null,
    dateStartCol: null,
    dateEndCol: null,
    staffRows: [],
    lunarRow: null,
    totalCol: null,
    totalValueRow: null,
  };

  const weekdayNames = ["日", "一", "二", "三", "四", "五", "六"];

  worksheet.eachRow((row, rowNumber) => {
    const firstCellValue = getCellStringValue(row.getCell(1));
    const secondCellValue = getCellStringValue(row.getCell(2));

    // Find date row - look for sequence starting with 1, 2, 3...
    if (!structure.dateRow) {
      for (let col = 1; col <= Math.min(40, row.cellCount || 40); col++) {
        const cellValue = row.getCell(col).value;
        if (cellValue === 1 || cellValue === "1") {
          const nextVal = row.getCell(col + 1).value;
          if (nextVal === 2 || nextVal === "2") {
            structure.dateRow = rowNumber;
            structure.dateStartCol = col;
            console.log(
              `Found date row at row ${rowNumber}, starting col ${col}`
            );
            for (let c = col; c <= col + 31; c++) {
              const v = row.getCell(c).value;
              if (v && (typeof v === "number" || /^\d+$/.test(v.toString()))) {
                structure.dateEndCol = c;
              }
            }
            break;
          }
        }
      }
    }

    // Find weekday row - check row after date row
    if (structure.dateRow && rowNumber === structure.dateRow + 1) {
      for (
        let col = structure.dateStartCol;
        col <= (structure.dateEndCol || structure.dateStartCol + 31);
        col++
      ) {
        const cellValue = getCellStringValue(row.getCell(col));
        if (weekdayNames.includes(cellValue)) {
          structure.weekdayRow = rowNumber;
          console.log(`Found weekday row at row ${rowNumber}`);
          break;
        }
      }
    }

    // Find staff rows - look for rows with number in first column
    // The staff name might be in column 2, 3, or even later columns
    if (firstCellValue && /^[1-9]$/.test(firstCellValue)) {
      // This row has a number 1-9 in first column, likely a staff row
      // Find which column has the name (look for Chinese characters)
      let nameCol = 2;
      for (let col = 2; col <= 5; col++) {
        const cellVal = getCellStringValue(row.getCell(col));
        if (cellVal && /[\u4e00-\u9fa5]{2,}/.test(cellVal)) {
          nameCol = col;
          break;
        }
      }
      structure.staffRows.push({ row: rowNumber, nameCol: nameCol });
      const actualName = getCellStringValue(row.getCell(nameCol));
      console.log(
        `Found staff row at row ${rowNumber}, nameCol: ${nameCol}, name: "${actualName}"`
      );
    }

    // Alternative: look for rows after weekday row that have Chinese names
    if (
      structure.weekdayRow &&
      rowNumber > structure.weekdayRow &&
      rowNumber < structure.weekdayRow + 10
    ) {
      // Check columns 2-5 for Chinese names
      for (let col = 2; col <= 5; col++) {
        const cellVal = getCellStringValue(row.getCell(col));
        if (cellVal && /[\u4e00-\u9fa5]{2,}/.test(cellVal)) {
          // Check if we already have this row
          const alreadyHave = structure.staffRows.some(
            (sr) => sr.row === rowNumber
          );
          if (!alreadyHave) {
            structure.staffRows.push({ row: rowNumber, nameCol: col });
            console.log(
              `Found additional staff row at row ${rowNumber}, nameCol: ${col}, name: "${cellVal}"`
            );
          }
          break;
        }
      }
    }

    // Find lunar row - look for 農曆 in any cell of the first few columns
    if (!structure.lunarRow) {
      for (let col = 1; col <= 3; col++) {
        const cellVal = getCellStringValue(row.getCell(col));
        if (
          cellVal.includes("農曆") ||
          cellVal.includes("農 曆") ||
          cellVal === "農曆"
        ) {
          structure.lunarRow = rowNumber;
          console.log(`Found lunar row at row ${rowNumber}`);
          break;
        }
      }
    }

    // Find total column - look for 合計
    row.eachCell((cell, colNumber) => {
      const value = getCellStringValue(cell);
      if (value.includes("合計") || value.includes("合 計")) {
        structure.totalCol = colNumber;
        structure.totalValueRow = rowNumber;
        console.log(`Found total column at col ${colNumber}, row ${rowNumber}`);
      }
    });
  });

  // Sort staff rows by row number
  structure.staffRows.sort((a, b) => a.row - b.row);

  console.log("Final structure analysis:", structure);
  return structure;
}

// Helper to get string value from cell
function getCellStringValue(cell) {
  if (!cell || cell.value === null || cell.value === undefined) return "";
  if (typeof cell.value === "string") return cell.value;
  if (typeof cell.value === "number") return cell.value.toString();
  if (typeof cell.value === "object") {
    if (cell.value.richText) {
      return cell.value.richText.map((r) => r.text || "").join("");
    }
    if (cell.value.text) return cell.value.text;
    if (cell.value.result !== undefined) return cell.value.result.toString();
  }
  return cell.value.toString();
}

// Helper to safely apply fill
function applySolidFill(cell, argbColor) {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: argbColor },
  };
}

function removeFill(cell) {
  cell.fill = {
    type: "pattern",
    pattern: "none",
  };
}

// Update year and month cell(s)
function updateYearMonth(worksheet, year, month, isFirstSheet, firstSheetName) {
  let yearMonthCellFound = false;
  let yearMonthCellAddress = null;

  // Search rows to find year/month cells
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 10) return;

    row.eachCell((cell, colNumber) => {
      let value = cell.value;
      const stringValue = getCellStringValue(cell);

      // Check if this cell contains a year/month pattern
      // Supported: "2026年1月", "2026.1", "2026年 1月" (with space)
      const hasYearMonth =
        /\d{4}\s*年\s*\d{1,2}\s*月/.test(stringValue) ||
        /\d{4}\.\d{1,2}/.test(stringValue);

      // Handle formula cells
      if (typeof value === "object" && value !== null && value.formula) {
        const formulaResult = value.result?.toString() || "";
        if (
          /\d{4}\s*年\s*\d{1,2}\s*月/.test(formulaResult) ||
          /\d{4}\.\d{1,2}/.test(formulaResult)
        ) {
          if (isFirstSheet) {
            // First sheet: Replace formula with plain text
            cell.value = `${year}年${month}月`;
            yearMonthCellAddress = cell.address;
            console.log(
              `First sheet: Replaced formula at ${cell.address} with: ${year}年${month}月`
            );
          } else {
            // Other sheets: Use formula referencing first sheet with the SAME cell address
            const formula = `ASC(${firstSheetName}!${cell.address})`;
            cell.value = { formula: formula };
            console.log(
              `Other sheet: Set formula at ${cell.address}: =${formula}`
            );
          }
          yearMonthCellFound = true;
          return;
        }
      }

      // Handle rich text
      if (typeof value === "object" && value !== null && value.richText) {
        const hasMatch = value.richText.some(
          (part) =>
            typeof part.text === "string" &&
            (/\d{4}\s*年\s*\d{1,2}\s*月/.test(part.text) ||
              /\d{4}\.\d{1,2}/.test(part.text))
        );

        if (hasMatch) {
          if (isFirstSheet) {
            cell.value = `${year}年${month}月`;
            console.log(
              `First sheet: Replaced rich text at ${cell.address} with plain text`
            );
          } else {
            const formula = `ASC(${firstSheetName}!${cell.address})`;
            cell.value = { formula: formula };
            console.log(
              `Other sheet: Set formula at ${cell.address} (was rich text)`
            );
          }
          yearMonthCellFound = true;
          return;
        }
      }

      // Handle plain string
      if (typeof value === "string" && hasYearMonth) {
        if (isFirstSheet) {
          cell.value = `${year}年${month}月`;
          console.log(`First sheet: Updated string at ${cell.address}`);
        } else {
          // For other sheets, convert to formula
          const formula = `ASC(${firstSheetName}!${cell.address})`;
          cell.value = { formula: formula };
          console.log(
            `Other sheet: Set formula at ${cell.address}: ${formula}`
          );
        }
        yearMonthCellFound = true;
      }
    });
  });

  console.log(`Year/month update complete. Found: ${yearMonthCellFound}`);
}

// Update date row with formatting based on structure
// Update date row with formatting based on structure
function updateDateRowWithFormatting(
  worksheet,
  structure,
  year,
  month,
  daysInMonth
) {
  console.log(
    `Updating date row ${structure.dateRow}, starting col ${structure.dateStartCol}, days: ${daysInMonth}`
  );
  const row = worksheet.getRow(structure.dateRow);
  const startCol = structure.dateStartCol;

  for (let day = 1; day <= 31; day++) {
    const col = startCol + day - 1;
    const cell = row.getCell(col);

    if (day <= daysInMonth) {
      const oldValue = cell.value;
      cell.value = day;
      if (day <= 3) {
        console.log(
          `Date cell [row ${structure.dateRow}, col ${col}]: ${oldValue} -> ${day}`
        );
      }

      const date = new Date(year, month - 1, day);
      const dayOfWeek = date.getDay();
      const existingFont = cell.font ? { ...cell.font } : {};

      if (dayOfWeek === 6) {
        // Saturday: Green font only
        cell.font = { ...existingFont, color: { argb: "FF008000" } };
        removeFill(cell);
      } else if (dayOfWeek === 0) {
        // Sunday: Red font + Yellow background
        cell.font = { ...existingFont, color: { argb: "FFFF0000" } };
        applySolidFill(cell, "FFFFFF00");
      } else {
        // Weekday: Black font only
        cell.font = { ...existingFont, color: { argb: "FF000000" } };
        removeFill(cell);
      }
    } else {
      cell.value = "";
      removeFill(cell);
    }
  }
  console.log("Date row update complete");
}

// Update weekday row
function updateWeekdayRow(worksheet, structure, year, month, daysInMonth) {
  const weekdayNames = ["日", "一", "二", "三", "四", "五", "六"];
  const row = worksheet.getRow(structure.weekdayRow);
  const startCol = structure.dateStartCol;

  for (let day = 1; day <= 31; day++) {
    const col = startCol + day - 1;
    const cell = row.getCell(col);

    if (day <= daysInMonth) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = date.getDay();
      cell.value = weekdayNames[dayOfWeek];

      const existingFont = cell.font ? { ...cell.font } : {};

      if (dayOfWeek === 6) {
        // Saturday: Green font only
        cell.font = { ...existingFont, color: { argb: "FF008000" } };
        removeFill(cell);
      } else if (dayOfWeek === 0) {
        // Sunday: Red font + Yellow background
        cell.font = { ...existingFont, color: { argb: "FFFF0000" } };
        applySolidFill(cell, "FFFFFF00");
      } else {
        // Weekday: Black font only
        cell.font = { ...existingFont, color: { argb: "FF000000" } };
        removeFill(cell);
      }
    } else {
      cell.value = "";
      removeFill(cell);
    }
  }
}

// Clear old staff names and fill new staff names from config
function clearAndFillStaffNames(worksheet, structure, config) {
  console.log("Clearing and filling staff names. Config:", config);
  console.log("Staff rows found:", structure.staffRows);

  // Get the valid staff list from config
  const staffList = config?.staffList?.filter((s) => s.trim()) || [];
  console.log("Valid staff list:", staffList);

  // Process each staff row
  structure.staffRows.forEach((staffInfo, index) => {
    const row = worksheet.getRow(staffInfo.row);
    const hasStaff = index < staffList.length;
    const staffName = hasStaff ? staffList[index] : "";

    // Determine row background: even-numbered staff get yellow, others white
    // Note: index is 0-based, so 1st staff = index 0 (odd, white), 2nd staff = index 1 (even, yellow)
    const isEvenStaff = hasStaff && (index + 1) % 2 === 0;

    // Set the name cell
    const nameCell = row.getCell(staffInfo.nameCol);
    nameCell.value = staffName;

    if (hasStaff) {
      if (isEvenStaff) {
        applySolidFill(nameCell, "FFFFFF00");
      } else {
        applySolidFill(nameCell, "FFFFFFFF"); // White
      }
    } else {
      applySolidFill(nameCell, "FFFFFFFF"); // White for empty
    }

    // Set background for all cells in the date range
    if (structure.dateStartCol) {
      for (
        let col = structure.dateStartCol;
        col <= structure.dateStartCol + 34;
        col++
      ) {
        const cell = row.getCell(col);
        cell.value = ""; // Clear any existing value

        if (hasStaff) {
          if (isEvenStaff) {
            applySolidFill(cell, "FFFFFF00");
          } else {
            applySolidFill(cell, "FFFFFFFF");
          }
        } else {
          applySolidFill(cell, "FFFFFFFF");
        }
      }
    }

    if (hasStaff) {
      console.log(
        `Set staff name at row ${staffInfo.row}: "${staffName}" (${
          isEvenStaff ? "yellow" : "white"
        } bg)`
      );
    } else {
      console.log(`Cleared row ${staffInfo.row} (white bg)`);
    }
  });
}

// Update lunar row with holidays
function updateLunarRowContent(worksheet, structure, daysInMonth, holidays) {
  console.log(
    `Updating lunar row ${structure.lunarRow}, startCol ${structure.dateStartCol}`
  );
  console.log("Holidays to apply:", holidays);

  const row = worksheet.getRow(structure.lunarRow);
  const startCol = structure.dateStartCol;

  for (let day = 1; day <= 31; day++) {
    const col = startCol + day - 1;
    const cell = row.getCell(col);
    const existingAlignment = cell.alignment ? { ...cell.alignment } : {};
    const oldValue = cell.value;

    if (day <= daysInMonth) {
      const holidayInfo = holidays[day];

      if (holidayInfo) {
        cell.value = holidayInfo;
        // Yellow background
        applySolidFill(cell, "FFFFFF00");

        // Font: 細明體-ExtB, size 11, blue color
        cell.font = {
          name: "細明體-ExtB",
          size: 11,
          color: { argb: "FF0000FF" },
        };
        // Vertical text alignment
        cell.alignment = {
          ...existingAlignment,
          textRotation: 255,
          vertical: "middle",
          horizontal: "center",
          wrapText: true,
        };
        console.log(
          `Lunar day ${day}: "${oldValue}" -> "${holidayInfo}" (yellow bg, blue text)`
        );
      } else {
        cell.value = "";
        applySolidFill(cell, "FFFFFFFF"); // White
      }
    } else {
      cell.value = "";
      applySolidFill(cell, "FFFFFFFF"); // White
    }
  }
  console.log("Lunar row update complete");
}

// Update total leave days
function updateTotalLeaveDays(worksheet, structure, totalLeaveDays) {
  console.log("Updating total leave days:", totalLeaveDays);
  console.log(
    "Total column:",
    structure.totalCol,
    "Total row:",
    structure.totalValueRow
  );

  // Find the 合計 cell and update the cell directly below it
  let updated = false;

  worksheet.eachRow((row, rowNumber) => {
    if (updated) return; // Only update once

    row.eachCell((cell, colNumber) => {
      if (updated) return;

      const value = getCellStringValue(cell);
      if (value === "合計" || value === "合 計") {
        console.log(`Found 合計 at row ${rowNumber}, col ${colNumber}`);
        // Update the cell directly below the 合計 header
        const valueCell = worksheet.getRow(rowNumber + 1).getCell(colNumber);
        valueCell.value = parseInt(totalLeaveDays) || totalLeaveDays;
        console.log(
          `Updated total at row ${
            rowNumber + 1
          }, col ${colNumber}: ${totalLeaveDays}`
        );
        updated = true;
      }
    });
  });

  if (!updated) {
    console.log("Could not find 合計 header to update");
  }
}

// ===== Reset App =====
function resetApp() {
  // Reset state
  appState.templateWorkbook = null;
  appState.templateBuffer = null;
  appState.rulesData = null;
  appState.sheetNames = [];
  appState.selectedSheets.clear();
  appState.sheetConfigs = {};
  appState.generatedBlob = null;

  // Reset UI
  document.getElementById("templateFileName").textContent = "";
  document.getElementById("rulesFileName").textContent = "";
  document.getElementById("templateUpload").classList.remove("uploaded");
  document.getElementById("rulesUpload").classList.remove("uploaded");
  document.getElementById("rulesPreview").classList.add("hidden");
  document.getElementById("step1Next").disabled = true;
  document.getElementById("downloadSection").classList.add("hidden");
  document.getElementById("progressContainer").classList.add("hidden");
  document.getElementById("generateBtn").disabled = false;

  // Reset file inputs
  document.getElementById("templateFile").value = "";
  document.getElementById("rulesFile").value = "";

  // Go to step 1
  goToStep(1);
}
