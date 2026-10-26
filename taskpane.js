
let currentResult = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    bindEvents();
    tryReadSelection();
  }
});

function bindEvents() {
  document.getElementById("readSelectionBtn").addEventListener("click", tryReadSelection);
  document.getElementById("convertBtn").addEventListener("click", convertFromInput);
  document.getElementById("numberInput").addEventListener("keydown", (e) => {
    if (e.key === "Enter") convertFromInput();
  });
  document.getElementById("addChan").addEventListener("change", () => {
    if (document.getElementById("numberInput").value.trim()) convertFromInput();
  });
  document.getElementById("insertAfterBtn").addEventListener("click", insertAfterSelection);
  document.getElementById("insertLineBtn").addEventListener("click", insertAsNewLine);
  document.getElementById("replaceBtn").addEventListener("click", replaceSelection);
  document.getElementById("copyBtn").addEventListener("click", copyResult);
}

async function tryReadSelection() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text.trim();
      if (!text) {
        setStatus("Hãy bôi đen số tiền trong Word trước.", true);
        return;
      }

      document.getElementById("numberInput").value = text;
      convertFromInput();
      setStatus("Đã lấy số tiền đang bôi đen.");
    });
  } catch (error) {
    setStatus("Không đọc được vùng chọn: " + error.message, true);
  }
}

function normalizeMoneyInput(raw) {
  let value = String(raw || "").trim();

  // Cho phép người dùng chọn cả cụm "... đồng", "đ", "₫".
  value = value.replace(/\bđồng\b/gi, "");
  value = value.replace(/₫/g, "");
  value = value.replace(/\bđ\b/gi, "");

  // Giữ đúng cách xử lý của tool gốc: bỏ dấu chấm, phẩy và khoảng trắng.
  value = value.replace(/[.,\s]/g, "");

  // Bỏ số 0 dư ở đầu.
  value = value.replace(/^0+(?=\d)/, "");

  return value;
}

function convertFromInput() {
  const resultBox = document.getElementById("resultBox");
  const raw = document.getElementById("numberInput").value;
  const cleaned = normalizeMoneyInput(raw);

  if (!cleaned || !/^\d+$/.test(cleaned)) {
    currentResult = "";
    document.getElementById("result").textContent = "Vui lòng nhập hoặc bôi đen một số tiền hợp lệ.";
    resultBox.classList.add("error");
    toggleActions(false);
    return;
  }

  if (cleaned.length > 21) {
    currentResult = "";
    document.getElementById("result").textContent = "Số tiền quá lớn. Tối đa 21 chữ số.";
    resultBox.classList.add("error");
    toggleActions(false);
    return;
  }

  let text = readNumberToText(cleaned);
  text = text.charAt(0).toUpperCase() + text.slice(1) + " đồng";

  if (document.getElementById("addChan").checked) {
    text += " chẵn";
  }

  currentResult = text;
  document.getElementById("result").textContent = currentResult;
  resultBox.classList.remove("error");
  toggleActions(true);
}

function toggleActions(enabled) {
  ["insertAfterBtn", "insertLineBtn", "replaceBtn", "copyBtn"].forEach((id) => {
    document.getElementById(id).disabled = !enabled;
  });
}

async function insertAfterSelection() {
  if (!currentResult) return;
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(` (Bằng chữ: ${currentResult}).`, Word.InsertLocation.after);
      await context.sync();
    });
    setStatus("Đã chèn số tiền bằng chữ sau vùng chọn.");
  } catch (error) {
    setStatus("Không chèn được: " + error.message, true);
  }
}

async function insertAsNewLine() {
  if (!currentResult) return;
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (paragraphs.items.length > 0) {
        paragraphs.items[paragraphs.items.length - 1]
          .insertParagraph(`Bằng chữ: ${currentResult}.`, Word.InsertLocation.after);
      } else {
        selection.insertText(`\nBằng chữ: ${currentResult}.`, Word.InsertLocation.after);
      }
      await context.sync();
    });
    setStatus("Đã chèn dòng “Bằng chữ”.");
  } catch (error) {
    setStatus("Không chèn được: " + error.message, true);
  }
}

async function replaceSelection() {
  if (!currentResult) return;
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(currentResult, Word.InsertLocation.replace);
      await context.sync();
    });
    setStatus("Đã thay vùng chọn bằng số tiền viết bằng chữ.");
  } catch (error) {
    setStatus("Không thay được: " + error.message, true);
  }
}

async function copyResult() {
  if (!currentResult) return;
  try {
    await navigator.clipboard.writeText(currentResult);
    setStatus("Đã sao chép kết quả.");
  } catch {
    const ta = document.createElement("textarea");
    ta.value = currentResult;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
    setStatus("Đã sao chép kết quả.");
  }
}

function setStatus(message, isError = false) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.classList.toggle("error", isError);
  window.clearTimeout(setStatus.timer);
  setStatus.timer = window.setTimeout(() => {
    el.textContent = "";
    el.classList.remove("error");
  }, 3500);
}

/*
 * Phần lõi dưới đây giữ nguyên thuật toán từ tool của tác giả:
 * đọc theo từng nhóm 3 chữ số, có linh / mốt / lăm.
 */
function readNumberToText(numberString) {
  const groupUnits = [
    "",
    " nghìn",
    " triệu",
    " tỷ",
    " nghìn tỷ",
    " triệu tỷ",
    " tỷ tỷ"
  ];

  const digits = [
    "không", "một", "hai", "ba", "bốn",
    "năm", "sáu", "bảy", "tám", "chín"
  ];

  if (/^0+$/.test(numberString)) return "không";

  const groups = [];
  for (let index = numberString.length; index > 0; index -= 3) {
    const start = Math.max(0, index - 3);
    groups.push(numberString.slice(start, index));
  }

  let highestNonZeroGroup = groups.length - 1;
  while (
    highestNonZeroGroup > 0 &&
    parseInt(groups[highestNonZeroGroup], 10) === 0
  ) {
    highestNonZeroGroup--;
  }

  const resultParts = [];

  for (let groupIndex = highestNonZeroGroup; groupIndex >= 0; groupIndex--) {
    const groupValue = parseInt(groups[groupIndex], 10);
    if (groupValue === 0) continue;

    const readFull =
      groupIndex < highestNonZeroGroup &&
      groupValue < 100;

    const groupText = readThreeDigits(groupValue, readFull, digits);
    resultParts.push(groupText + groupUnits[groupIndex]);
  }

  return resultParts.join(" ").replace(/\s+/g, " ").trim();
}

function readThreeDigits(number, readFull, digits) {
  const hundred = Math.floor(number / 100);
  const ten = Math.floor((number % 100) / 10);
  const unit = number % 10;
  const parts = [];

  if (hundred > 0 || readFull) {
    if (hundred > 0) parts.push(digits[hundred] + " trăm");
    else parts.push("không trăm");

    if (ten === 0 && unit > 0) parts.push("linh");
  }

  if (ten > 0) {
    if (ten === 1) parts.push("mười");
    else parts.push(digits[ten] + " mươi");
  }

  if (unit > 0) {
    if (ten > 1 && unit === 1) parts.push("mốt");
    else if (ten > 0 && unit === 5) parts.push("lăm");
    else parts.push(digits[unit]);
  }

  return parts.join(" ").trim();
}
