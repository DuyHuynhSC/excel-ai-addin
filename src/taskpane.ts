/* global Office */
import './functions'; // Đảm bảo custom functions được tải và đăng ký vào Shared Runtime
import { ApiFactory } from './api/apiFactory';
import { TranslationService } from './services/translationService';
import { AnalysisService } from './services/analysisService';
import { SummarizationService } from './services/summarizationService';
import { HistoryService } from './services/historyService';
import { GlossaryService } from './services/glossaryService';
import { getSelectedCellsData, writeResultAdjacent, writeResultOverwrite, writeToSummariesSheet } from './utils/excelHelpers';
import { loadConfig, saveConfig } from './utils/configLoader';
import { encryptText } from './utils/cryptoUtils';
import { AppConfig, GlossaryItem, HistoryItem, TaskType } from './types';

// Các instances dịch vụ dùng chung trong Taskpane
let apiFactory: ApiFactory;
let historyService: HistoryService;
let glossaryService: GlossaryService;
let translationService: TranslationService;
let analysisService: AnalysisService;
let summarizationService: SummarizationService;

// AbortController để hủy các yêu cầu đang chạy
let currentAbortController: AbortController | null = null;

// Khởi chạy Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Khởi tạo các services
    const config = loadConfig();
    apiFactory = new ApiFactory(config);
    historyService = new HistoryService();
    glossaryService = new GlossaryService();
    translationService = new TranslationService(apiFactory);
    analysisService = new AnalysisService(apiFactory);
    summarizationService = new SummarizationService(apiFactory);

    // Thiết lập giao diện và sự kiện
    initUi(config);
    checkHealthPeriodically();
  }
});

/**
 * Thiết lập các sự kiện giao diện
 */
function initUi(config: AppConfig) {
  setupTabs();
  setupSettingsForm(config);
  setupTaskActions();
  setupHistoryActions();
  setupGlossaryActions();
  setupStatusPanel();
}

/**
 * Điều hướng Tab
 */
function setupTabs() {
  const tabs = document.querySelectorAll('.nav-tab');
  const panes = document.querySelectorAll('.tab-pane');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetTab = tab.getAttribute('data-tab');

      tabs.forEach(t => t.classList.remove('active'));
      panes.forEach(p => p.classList.remove('active'));

      tab.classList.add('active');
      const activePane = document.getElementById(targetTab!);
      if (activePane) activePane.classList.add('active');

      // Tải lại lịch sử nếu mở Tab Lịch sử
      if (targetTab === 'tab-history') {
        renderHistoryList();
      }

      // Tải lại thuật ngữ nếu mở Tab Thuật ngữ
      if (targetTab === 'tab-glossary') {
        renderGlossaryList();
      }
    });
  });
}

/**
 * Đồng bộ dữ liệu Tab Cài đặt
 */
function setupSettingsForm(config: AppConfig) {
  const selectOllama = document.getElementById('service-ollama') as HTMLInputElement;
  const selectCustom = document.getElementById('service-custom') as HTMLInputElement;
  const inputOllamaUrl = document.getElementById('setting-ollama-url') as HTMLInputElement;
  const inputCustomUrl = document.getElementById('setting-custom-url') as HTMLInputElement;
  const inputCustomKey = document.getElementById('setting-custom-key') as HTMLInputElement;
  const checkFallback = document.getElementById('setting-fallback') as HTMLInputElement;
  const checkStreaming = document.getElementById('setting-streaming') as HTMLInputElement;
  const checkProxyDisabled = document.getElementById('setting-custom-proxy-disabled') as HTMLInputElement;

  const modelTranslation = document.getElementById('model-translation') as HTMLInputElement;
  const modelAnalysis = document.getElementById('model-analysis') as HTMLInputElement;
  const modelSummarization = document.getElementById('model-summarization') as HTMLInputElement;

  // Điền thông tin cũ vào form
  if (config.activeService === 'custom') {
    selectCustom.checked = true;
  } else {
    selectOllama.checked = true;
  }

  inputOllamaUrl.value = config.ollamaUrl;
  inputCustomUrl.value = config.customApiUrl;
  checkFallback.checked = config.fallbackEnabled;
  checkStreaming.checked = config.streamingEnabled;
  if (checkProxyDisabled) {
    checkProxyDisabled.checked = !!config.customApiProxyDisabled;
  }

  modelTranslation.value = config.models.translation;
  modelAnalysis.value = config.models.analysis;
  modelSummarization.value = config.models.summarization;

  // Sự kiện Lưu cấu hình
  const btnSave = document.getElementById('btn-save-settings');
  btnSave?.addEventListener('click', async () => {
    btnSave.textContent = 'Đang lưu...';
    btnSave.setAttribute('disabled', 'true');

    try {
      const activeService = selectCustom.checked ? 'custom' : 'ollama';
      const currentConfig = loadConfig();
      let customApiKey = currentConfig.customApiKey;

      // Chỉ mã hóa nếu người dùng nhập Key mới (không rỗng hoặc không phải kí tự che mắt)
      const inputKeyVal = inputCustomKey.value.trim();
      if (inputKeyVal && !inputKeyVal.includes('••••')) {
        customApiKey = await encryptText(inputKeyVal);
      }

      const newConfig: AppConfig = {
        ollamaUrl: inputOllamaUrl.value.trim(),
        customApiUrl: inputCustomUrl.value.trim(),
        customApiKey: customApiKey,
        activeService: activeService as 'ollama' | 'custom',
        fallbackEnabled: checkFallback.checked,
        streamingEnabled: checkStreaming.checked,
        customApiProxyDisabled: checkProxyDisabled ? checkProxyDisabled.checked : false,
        models: {
          translation: modelTranslation.value.trim() || 'translator',
          analysis: modelAnalysis.value.trim() || 'translator',
          summarization: modelSummarization.value.trim() || 'translator',
          custom: currentConfig.models.custom
        }
      };

      saveConfig(newConfig);
      await apiFactory.updateConfig(newConfig);

      showStatusOverlay('success', 'Đã lưu cấu hình thành công!');
      // Xóa hiển thị key thô sau khi lưu
      inputCustomKey.value = '';
      
      // Thực hiện kiểm tra kết nối ngay lập tức
      await runHealthCheck();
    } catch (e) {
      showStatusOverlay('error', `Lỗi khi lưu: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      btnSave.textContent = 'Lưu cấu hình';
      btnSave.removeAttribute('disabled');
    }
  });

  // Sự kiện Kiểm tra kết nối trực tiếp
  const btnTest = document.getElementById('btn-test-settings');
  btnTest?.addEventListener('click', async () => {
    btnTest.textContent = 'Đang kiểm tra...';
    btnTest.setAttribute('disabled', 'true');
    showStatusOverlay('pending', '⏳ Đang kiểm tra kết nối tới các dịch vụ...');

    try {
      const activeService = selectCustom.checked ? 'custom' : 'ollama';
      const currentConfig = loadConfig();
      let customApiKey = currentConfig.customApiKey;

      const inputKeyVal = inputCustomKey.value.trim();
      if (inputKeyVal && !inputKeyVal.includes('••••')) {
        customApiKey = await encryptText(inputKeyVal);
      }

      const tempConfig: AppConfig = {
        ollamaUrl: inputOllamaUrl.value.trim(),
        customApiUrl: inputCustomUrl.value.trim(),
        customApiKey: customApiKey,
        activeService: activeService as 'ollama' | 'custom',
        fallbackEnabled: checkFallback.checked,
        streamingEnabled: checkStreaming.checked,
        customApiProxyDisabled: checkProxyDisabled ? checkProxyDisabled.checked : false,
        models: {
          translation: modelTranslation.value.trim() || 'translator',
          analysis: modelAnalysis.value.trim() || 'translator',
          summarization: modelSummarization.value.trim() || 'translator',
          custom: currentConfig.models.custom
        }
      };

      const tempFactory = new ApiFactory(tempConfig);
      const health = await tempFactory.checkHealth();

      let resultText = '';
      let status: 'success' | 'error' = 'success';

      if (tempConfig.activeService === 'custom') {
        if (health.custom) {
          resultText = '✅ Kết nối tới Custom API Gateway thành công!';
        } else {
          resultText = '❌ Kết nối tới Custom API Gateway thất bại. Vui lòng kiểm tra lại URL và API Key.';
          status = 'error';
        }
      } else {
        if (health.ollama) {
          resultText = '✅ Kết nối tới Ollama (Local) thành công!';
        } else {
          resultText = '❌ Kết nối tới Ollama (Local) thất bại. Vui lòng kiểm tra xem Ollama đã được khởi chạy chưa.';
          status = 'error';
        }
      }

      showStatusOverlay(status, resultText);
    } catch (e) {
      showStatusOverlay('error', `❌ Lỗi kiểm tra kết nối: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      btnTest.textContent = 'Kiểm tra kết nối';
      btnTest.removeAttribute('disabled');
    }
  });
}

/**
 * Cấu hình xử lý cho các tác vụ chính
 */
function setupTaskActions() {
  // 1. Tác vụ Dịch thuật
  document.getElementById('btn-run-translate')?.addEventListener('click', () => {
    const destInput = document.querySelector('input[name="translate-dest"]:checked') as HTMLInputElement;
    const destType = destInput?.value || 'adjacent';

    executeTask('translation', async (text, signal, onChunk) => {
      const sourceLang = (document.getElementById('translate-source') as HTMLSelectElement).value;
      const targetLang = (document.getElementById('translate-target') as HTMLSelectElement).value;
      const toneInput = document.querySelector('input[name="translate-tone"]:checked') as HTMLInputElement;
      const tone = (toneInput?.value || 'neutral') as 'formal' | 'informal' | 'neutral';
      const config = loadConfig();

      return await translationService.translate({
        text,
        sourceLang,
        targetLang,
        tone,
        stream: config.streamingEnabled,
        onChunk,
        signal
      });
    }, async (_sourceText: string, result: string, _model: string, cellInfo: { row: number; col: number }) => {
      if (destType === 'overwrite') {
        // Ghi đè trực tiếp vào ô gốc đang chọn
        await writeResultOverwrite(cellInfo.row, cellInfo.col, result);
        return 'Đã dịch đè vào ô đang chọn.';
      } else if (destType === 'taskpane') {
        // Chỉ hiển thị kết quả trên Sidebar Taskpane
        return 'Đã hoàn thành dịch thuật hiển thị trên Taskpane.';
      } else {
        // Ghi vào ô bên cạnh cột kế tiếp
        await writeResultAdjacent(cellInfo.row, cellInfo.col + 1, result);
        return 'Đã điền kết quả dịch vào cột kế bên.';
      }
    });
  });

  // 2. Tác vụ Phân tích
  document.getElementById('btn-run-analyze')?.addEventListener('click', () => {
    executeTask('analysis', async (text, signal, onChunk) => {
      const focusArea = (document.getElementById('analyze-focus') as HTMLSelectElement).value as 'general' | 'sentiment' | 'entities' | 'data_cleanup';
      const customInstructions = (document.getElementById('analyze-custom') as HTMLTextAreaElement).value.trim();
      const config = loadConfig();

      return await analysisService.analyze({
        text,
        focusArea,
        customInstructions: customInstructions || undefined,
        stream: config.streamingEnabled,
        onChunk,
        signal
      });
    });
  });

  // 3. Tác vụ Tóm tắt
  document.getElementById('btn-run-summarize')?.addEventListener('click', () => {
    executeTask('summarization', async (text, signal, onChunk) => {
      const format = (document.getElementById('summarize-format') as HTMLSelectElement).value as 'paragraph' | 'bullets' | 'key_actions';
      const lengthInput = document.querySelector('input[name="summarize-length"]:checked') as HTMLInputElement;
      const length = (lengthInput?.value || 'concise') as 'concise' | 'detailed';
      const config = loadConfig();

      return await summarizationService.summarize({
        text,
        format,
        length,
        stream: config.streamingEnabled,
        onChunk,
        signal
      });
    }, async (sourceText, result, model) => {
      const newSheetCheck = document.getElementById('summarize-new-sheet') as HTMLInputElement;
      if (newSheetCheck?.checked) {
        // Ghi vào Sheet tóm tắt riêng biệt
        const sheetName = await writeToSummariesSheet(sourceText, result, model);
        return `Đã lưu kết quả vào trang tính "${sheetName}".`;
      }
      return null; // Tiếp tục ghi đè cột kế bên mặc định
    });
  });
}

/**
 * Thực thi một Task AI tổng quát
 */
async function executeTask(
  taskType: TaskType,
  apiCall: (text: string, signal: AbortSignal, onChunk: (chunk: string) => void) => Promise<{ result: string; model: string; processingTime: number }>,
  customWriteResult?: (sourceText: string, result: string, model: string, cellInfo: { row: number; col: number }) => Promise<string | null>
) {
  // Hủy tiến trình trước đó nếu có
  if (currentAbortController) {
    currentAbortController.abort();
  }

  currentAbortController = new AbortController();
  const signal = currentAbortController.signal;

  const config = loadConfig();
  if (config.activeService === 'custom' && config.customApiUrl.includes('api.company.com')) {
    showStatusOverlay('error', '⚠️ Bạn chưa cấu hình địa chỉ API thật của doanh nghiệp! Vui lòng vào tab Cài đặt, thay thế địa chỉ mẫu "api.company.com" bằng địa chỉ máy chủ API thực tế của bạn, rồi nhấn Lưu cấu hình.');
    return;
  }

  showStatusOverlay('pending', '⏳ Đang đọc ô dữ liệu được chọn...');

  try {
    const selectedCells = await getSelectedCellsData();
    // Giới hạn xử lý song song tối đa 3 ô cùng lúc để tránh nghẽn luồng
    const maxConcurrency = 3;
    let completedCount = 0;

    showStatusOverlay('pending', `⏳ Đang xử lý ${selectedCells.length} ô dữ liệu...`);

    // Thực thi tuần tự hoặc song song có giới hạn
    for (let i = 0; i < selectedCells.length; i += maxConcurrency) {
      const chunk = selectedCells.slice(i, i + maxConcurrency);
      const promises = chunk.map(async (cell, index) => {
        let accumulatedText = '';
        const onChunkCallback = (textChunk: string) => {
          // Hiển thị tiến trình stream thời gian thực của ô đầu tiên trong chunk lên bảng Panel
          if (index === 0) {
            const streamEl = document.getElementById('stream-output-content');
            if (streamEl) {
              accumulatedText += textChunk;
              streamEl.innerHTML = markdownToHtml(accumulatedText);
              streamEl.scrollTop = streamEl.scrollHeight; // Auto scroll xuống cuối
            }
          }
        };

        // Reset khung hiển thị stream
        if (index === 0) {
          const streamEl = document.getElementById('stream-output-content');
          if (streamEl) streamEl.innerHTML = '';
        }

        const response = await apiCall(cell.value, signal, onChunkCallback);
        
        let writeMessage = '';
        let targetWritten = false;

        // Nếu có xử lý ghi kết quả tùy chỉnh (như lưu sheet tóm tắt, dịch đè, dịch vào taskpane)
        if (customWriteResult) {
          const customMsg = await customWriteResult(cell.value, response.result, response.model, { row: cell.row, col: cell.col });
          if (customMsg) {
            writeMessage = customMsg;
            targetWritten = true;
          }
        }

        // Nếu không có ghi tùy chỉnh, mặc định ghi vào ô bên cạnh (cột kế tiếp)
        if (!targetWritten) {
          await writeResultAdjacent(cell.row, cell.col + 1, response.result);
          writeMessage = 'Đã điền kết quả vào cột kế bên.';
        }

        completedCount++;
        showStatusOverlay(
          'pending', 
          `⏳ Đang xử lý: ${completedCount}/${selectedCells.length} ô. ${writeMessage}`
        );

        // Lưu vào IndexedDB lịch sử
        const activeConfig = loadConfig();
        await historyService.addHistoryItem({
          task: taskType,
          sourceText: cell.value,
          result: response.result,
          model: response.model,
          processingTime: response.processingTime,
          serviceUsed: activeConfig.activeService,
          status: 'success'
        });
      });

      await Promise.all(promises);
    }

    showStatusOverlay('success', `✅ Đã xử lý xong toàn bộ ${selectedCells.length} ô dữ liệu!`);
  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') {
      showStatusOverlay('error', '❌ Đã hủy thao tác xử lý.');
      return;
    }

    const errorMsg = error instanceof Error ? error.message : String(error);
    showStatusOverlay('error', `❌ Lỗi: ${errorMsg}`);

    // Ghi nhận lỗi vào lịch sử nếu có thông tin văn bản gốc
    try {
      const activeConfig = loadConfig();
      await historyService.addHistoryItem({
        task: taskType,
        sourceText: '[Lỗi dữ liệu đầu vào hoặc lỗi hệ thống]',
        result: errorMsg,
        model: 'unknown',
        processingTime: 0,
        serviceUsed: activeConfig.activeService,
        status: 'error',
        errorMessage: errorMsg
      });
    } catch {}
  } finally {
    currentAbortController = null;
  }
}

/**
 * Thiết lập nhật ký hoạt động
 */
function setupHistoryActions() {
  const btnClear = document.getElementById('btn-clear-history') as HTMLButtonElement | null;
  let confirmMode = false;
  let confirmTimeout: number | null = null;

  btnClear?.addEventListener('click', async () => {
    if (!confirmMode) {
      confirmMode = true;
      btnClear.textContent = '⚠️ Xác nhận xóa?';
      btnClear.style.color = 'var(--color-error)';
      
      confirmTimeout = window.setTimeout(() => {
        confirmMode = false;
        btnClear.textContent = 'Xóa lịch sử';
        btnClear.style.color = '';
      }, 4000);
    } else {
      if (confirmTimeout) {
        window.clearTimeout(confirmTimeout);
      }
      confirmMode = false;
      btnClear.textContent = 'Xóa lịch sử';
      btnClear.style.color = '';
      
      try {
        await historyService.clearHistory();
        renderHistoryList();
        showStatusOverlay('success', 'Đã xóa toàn bộ nhật ký lịch sử thành công!');
        setTimeout(() => {
          document.getElementById('status-panel')?.classList.add('hidden');
        }, 1500);
      } catch (err) {
        showStatusOverlay('error', `Lỗi khi xóa lịch sử: ${err}`);
      }
    }
  });
}

/**
 * Hiển thị danh sách lịch sử trong tab Lịch sử
 */
async function renderHistoryList() {
  const container = document.getElementById('history-container');
  if (!container) return;

  try {
    const list = await historyService.getAllHistory();
    if (list.length === 0) {
      container.innerHTML = '<p class="history-empty">Chưa có hoạt động nào được lưu trữ.</p>';
      return;
    }

    container.innerHTML = list.map((item: HistoryItem) => {
      const dateStr = new Date(item.timestamp).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) + 
        ' ' + new Date(item.timestamp).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });
      
      const statusIcon = item.status === 'success' ? '✅' : '❌';
      const taskNames: Record<TaskType, string> = {
        translation: 'Dịch thuật',
        analysis: 'Phân tích',
        summarization: 'Tóm tắt',
        custom: 'Tùy chọn'
      };

      return `
        <div class="history-card">
          <div class="history-card-meta">
            <span class="history-card-task ${item.task}">${statusIcon} ${taskNames[item.task] || item.task}</span>
            <span>${dateStr} | ${item.serviceUsed.toUpperCase()}</span>
          </div>
          <div class="history-card-text"><strong>Gốc:</strong> ${escapeHtml(item.sourceText)}</div>
          <div class="history-card-res"><strong>Kết quả:</strong> ${escapeHtml(item.result)}</div>
          <div class="meta-output" style="margin-top: 4px; padding: 0;">
            <span>Model: ${item.model}</span>
            <span>⏱️ ${item.processingTime.toFixed(1)}s</span>
          </div>
        </div>
      `;
    }).join('');
  } catch (err) {
    container.innerHTML = `<p class="history-empty" style="color: var(--color-error)">Không thể tải lịch sử: ${err}</p>`;
  }
}

/**
 * Thiết lập các sự kiện liên quan đến Thuật ngữ (Glossary)
 */
function setupGlossaryActions() {
  // 1. Thêm thuật ngữ thủ công
  const btnAdd = document.getElementById('btn-add-glossary');
  btnAdd?.addEventListener('click', async () => {
    const jpInput = document.getElementById('glossary-jp-input') as HTMLInputElement;
    const vnInput = document.getElementById('glossary-vn-input') as HTMLInputElement;
    const noteInput = document.getElementById('glossary-note-input') as HTMLInputElement;

    const jp = jpInput.value.trim();
    const vn = vnInput.value.trim();
    const note = noteInput.value.trim();

    if (!jp || !vn) {
      showStatusOverlay('error', '⚠️ Vui lòng nhập đầy đủ từ gốc tiếng Nhật và bản dịch chuẩn tiếng Việt!');
      return;
    }

    try {
      await glossaryService.addGlossaryItem({ japanese: jp, vietnamese: vn, note: note || undefined });
      jpInput.value = '';
      vnInput.value = '';
      noteInput.value = '';
      
      showStatusOverlay('success', `Đã thêm thuật ngữ: "${jp}" -> "${vn}"`);
      setTimeout(() => {
        document.getElementById('status-panel')?.classList.add('hidden');
      }, 2000);

      renderGlossaryList();
    } catch (err) {
      showStatusOverlay('error', `Lỗi khi lưu thuật ngữ: ${err}`);
    }
  });

  // 2. Kích hoạt trigger chọn file import
  const btnTriggerImport = document.getElementById('btn-trigger-file-import');
  const fileInput = document.getElementById('glossary-file-input') as HTMLInputElement;
  const fileNameDisplay = document.getElementById('glossary-file-name');

  btnTriggerImport?.addEventListener('click', () => {
    fileInput.click();
  });

  // 3. Xử lý khi chọn file import
  fileInput?.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    if (!file) return;

    if (fileNameDisplay) fileNameDisplay.textContent = file.name;

    const reader = new FileReader();
    reader.onload = async (e) => {
      const content = e.target?.result as string;
      if (!content) return;

      try {
        let items: GlossaryItem[] = [];
        const ext = file.name.split('.').pop()?.toLowerCase();

        if (ext === 'json') {
          items = glossaryService.parseJSON(content);
        } else if (ext === 'csv') {
          items = glossaryService.parseCSV(content);
        } else if (ext === 'md') {
          items = glossaryService.parseMarkdown(content);
        } else {
          showStatusOverlay('error', '⚠️ Chỉ hỗ trợ định dạng .json, .csv hoặc .md!');
          return;
        }

        if (items.length === 0) {
          showStatusOverlay('error', '⚠️ Không tìm thấy thuật ngữ nào hợp lệ trong tệp!');
          return;
        }

        const count = await glossaryService.addGlossaryItems(items);
        showStatusOverlay('success', `🎉 Import thành công ${count} thuật ngữ từ tệp "${file.name}"!`);
        
        fileInput.value = '';
        if (fileNameDisplay) fileNameDisplay.textContent = 'Chưa chọn tệp';

        renderGlossaryList();
      } catch (err) {
        showStatusOverlay('error', `Lỗi khi import file: ${err instanceof Error ? err.message : String(err)}`);
      }
    };

    reader.onerror = () => {
      showStatusOverlay('error', 'Lỗi khi đọc tệp tin!');
    };

    reader.readAsText(file, 'UTF-8');
  });

  // 4. Xóa toàn bộ thuật ngữ
  const btnClear = document.getElementById('btn-clear-glossary') as HTMLButtonElement | null;
  let glossaryConfirmMode = false;
  let glossaryConfirmTimeout: number | null = null;

  btnClear?.addEventListener('click', async () => {
    if (!glossaryConfirmMode) {
      glossaryConfirmMode = true;
      btnClear.textContent = '⚠️ Xác nhận xóa toàn bộ?';
      btnClear.style.color = 'var(--color-error)';
      
      glossaryConfirmTimeout = window.setTimeout(() => {
        glossaryConfirmMode = false;
        btnClear.textContent = 'Xóa toàn bộ';
        btnClear.style.color = '';
      }, 4000);
    } else {
      if (glossaryConfirmTimeout) {
        window.clearTimeout(glossaryConfirmTimeout);
      }
      glossaryConfirmMode = false;
      btnClear.textContent = 'Xóa toàn bộ';
      btnClear.style.color = '';
      
      try {
        await glossaryService.clearGlossary();
        renderGlossaryList();
        showStatusOverlay('success', 'Đã xóa toàn bộ thuật ngữ chuyên ngành thành công!');
        setTimeout(() => {
          document.getElementById('status-panel')?.classList.add('hidden');
        }, 1500);
      } catch (err) {
        showStatusOverlay('error', `Lỗi khi xóa bảng thuật ngữ: ${err}`);
      }
    }
  });
}

/**
 * Hiển thị danh sách thuật ngữ ra giao diện
 */
async function renderGlossaryList() {
  const tableBody = document.getElementById('glossary-table-body');
  const countSpan = document.getElementById('glossary-count');
  if (!tableBody) return;

  try {
    const list = await glossaryService.getGlossaryItems();
    if (countSpan) countSpan.textContent = String(list.length);

    if (list.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="3" style="padding: 16px; text-align: center; color: var(--color-text-muted);">
            Chưa có thuật ngữ nào. Bạn hãy thêm thủ công hoặc import từ tệp tin (.json, .csv, .md).
          </td>
        </tr>
      `;
      return;
    }

    tableBody.innerHTML = list.map((item: GlossaryItem) => {
      const escapedJp = escapeHtml(item.japanese);
      const escapedVn = escapeHtml(item.vietnamese);
      const noteHtml = item.note ? `<div style="font-size: 9px; color: var(--color-text-muted); margin-top: 2px;">* ${escapeHtml(item.note)}</div>` : '';

      return `
        <tr style="border-bottom: 1px solid var(--border-color); vertical-align: top;">
          <td style="padding: 8px; font-weight: 500; color: #637ff9; word-break: break-all;">${escapedJp}</td>
          <td style="padding: 8px; word-break: break-all;">
            <div>${escapedVn}</div>
            ${noteHtml}
          </td>
          <td style="padding: 8px; text-align: center; vertical-align: middle;">
            <button class="btn-delete-term" data-jp="${escapedJp}" style="background: none; border: none; color: var(--color-error); cursor: pointer; font-size: 14px; padding: 2px 6px;">✕</button>
          </td>
        </tr>
      `;
    }).join('');

    // Đăng ký sự kiện xóa cho từng dòng
    const deleteButtons = tableBody.querySelectorAll('.btn-delete-term');
    deleteButtons.forEach(btn => {
      btn.addEventListener('click', async (e) => {
        const jpVal = (e.currentTarget as HTMLElement).getAttribute('data-jp');
        if (jpVal) {
          try {
            await glossaryService.deleteGlossaryItem(jpVal);
            renderGlossaryList();
          } catch (err) {
            showStatusOverlay('error', `Lỗi khi xóa thuật ngữ: ${err}`);
          }
        }
      });
    });

  } catch (err) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="3" style="padding: 16px; text-align: center; color: var(--color-error);">
          Không thể tải dữ liệu thuật ngữ: ${err}
        </td>
      </tr>
    `;
  }
}

/**
 * Điều khiển bảng trạng thái Realtime / Console Overlay
 */
function setupStatusPanel() {
  document.getElementById('btn-close-status')?.addEventListener('click', () => {
    document.getElementById('status-panel')?.classList.add('hidden');
  });
}

/**
 * Cập nhật giao diện Trạng thái
 */
function cleanErrorMessage(msg: string): string {
  if (msg.includes('PROXY_CONNECTION_ERROR')) {
    return msg.replace('PROXY_CONNECTION_ERROR', 'Lỗi kết nối từ Add-in Proxy tới Server đích');
  }
  return msg;
}

function markdownToHtml(text: string): string {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .split(/\n\n+/).map(para => {
      const trimmed = para.trim();
      if (!trimmed) return '';
      if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
        const items = trimmed.split(/\n[*-] /).map(item => `<li>${item.replace(/^[*-] /, '')}</li>`).join('');
        return `<ul>${items}</ul>`;
      }
      if (/^\d+\.\s/.test(trimmed)) {
        const items = trimmed.split(/\n\d+\.\s/).map(item => `<li>${item.replace(/^\d+\.\s/, '')}</li>`).join('');
        return `<ol>${items}</ol>`;
      }
      return `<p>${trimmed.replace(/\n/g, '<br>')}</p>`;
    }).join('')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>');
}

/**
 * Cập nhật giao diện Trạng thái
 */
function showStatusOverlay(status: 'pending' | 'success' | 'error', text: string) {
  const panel = document.getElementById('status-panel');
  const spinner = document.getElementById('status-spinner');
  const statusTxt = document.getElementById('status-text');

  if (!panel) return;

  panel.classList.remove('hidden');
  const cleanedText = cleanErrorMessage(text);

  if (status === 'pending') {
    spinner?.classList.remove('hidden');
    if (statusTxt) statusTxt.textContent = cleanedText;
    panel.style.borderTopColor = 'var(--color-primary)';
  } else if (status === 'success') {
    spinner?.classList.add('hidden');
    if (statusTxt) statusTxt.textContent = cleanedText;
    panel.style.borderTopColor = 'var(--color-success)';
  } else {
    spinner?.classList.add('hidden');
    if (statusTxt) statusTxt.textContent = cleanedText;
    panel.style.borderTopColor = 'var(--color-error)';
  }
}

/**
 * Kiểm tra kết nối định kỳ mỗi 30s
 */
function checkHealthPeriodically() {
  runHealthCheck();
  // Lên lịch chạy mỗi 30s
  window.setInterval(runHealthCheck, 30000);
}

async function runHealthCheck() {
  const dot = document.getElementById('connection-dot');
  if (!dot) return;

  try {
    const health = await apiFactory.checkHealth();
    const config = loadConfig();

    if (config.activeService === 'custom') {
      if (health.custom.ok) {
        dot.className = 'connection-status-dot online';
        dot.title = 'Custom API Gateway: Sẵn sàng';
      } else if (config.fallbackEnabled && health.ollama.ok) {
        dot.className = 'connection-status-dot warning';
        dot.title = `Custom API Lỗi: ${cleanErrorMessage(health.custom.error || '')} -> Đã chuyển dự phòng sang Ollama`;
      } else {
        dot.className = 'connection-status-dot offline';
        dot.title = `Custom API Mất kết nối! Chi tiết: ${cleanErrorMessage(health.custom.error || '')}`;
      }
    } else {
      if (health.ollama.ok) {
        dot.className = 'connection-status-dot online';
        dot.title = 'Ollama Local: Sẵn sàng';
      } else if (config.fallbackEnabled && health.custom.ok) {
        dot.className = 'connection-status-dot warning';
        dot.title = `Ollama Lỗi: ${cleanErrorMessage(health.ollama.error || '')} -> Đã chuyển dự phòng sang Custom API`;
      } else {
        dot.className = 'connection-status-dot offline';
        dot.title = `Ollama Local Mất kết nối! Chi tiết: ${cleanErrorMessage(health.ollama.error || '')}`;
      }
    }
  } catch (err) {
    dot.className = 'connection-status-dot offline';
    const errMsg = err instanceof Error ? err.message : String(err);
    dot.title = `Không thể kiểm tra trạng thái kết nối! Lỗi: ${cleanErrorMessage(errMsg)}`;
  }
}

/**
 * Hàm escape HTML cơ bản để ngăn chặn XSS trong giao diện Lịch sử
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
