/* global Excel */

export interface CellData {
  value: string;
  row: number;
  col: number;
}

/**
 * Đọc nội dung văn bản từ các ô đang được chọn trong bảng tính Excel.
 * Hỗ trợ chọn một ô hoặc một vùng nhiều ô.
 */
export async function getSelectedCellsData(): Promise<CellData[]> {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    // Load các thuộc tính cần dùng
    range.load(['values', 'rowCount', 'columnCount']);
    await context.sync();

    const data: CellData[] = [];
    const values = range.values;

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const val = values[r][c];
        // Chỉ xử lý các ô có chứa ký tự chữ/số
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          data.push({
            value: String(val),
            row: r,
            col: c
          });
        }
      }
    }

    if (data.length === 0) {
      throw new Error('Vui lòng chọn ít nhất một ô có dữ liệu.');
    }

    return data;
  });
}

/**
 * Ghi kết quả xử lý vào ô nằm bên phải ngay cạnh ô gốc được chọn (cột tiếp theo)
 * @param relativeRow Khoảng cách dòng tương đối so với ô đầu tiên được chọn
 * @param relativeCol Khoảng cách cột tương đối (mặc định = 1 là cột bên phải cạnh bên)
 * @param text Văn bản kết quả cần ghi vào
 */
export async function writeResultAdjacent(
  relativeRow: number,
  relativeCol: number,
  text: string
): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    // Lấy ô mục tiêu dựa trên offset tương đối
    const targetCell = range.getCell(relativeRow, relativeCol);
    targetCell.values = [[text]];
    
    // Tự động định dạng wrap text cho dễ đọc
    targetCell.format.wrapText = true;
    targetCell.format.autofitRows();
    
    await context.sync();
  });
}

/**
 * Ghi kết quả tóm tắt vào một trang tính riêng mang tên "Summaries".
 * Nếu trang tính chưa tồn tại thì tự động tạo mới kèm dòng tiêu đề đẹp đẽ.
 * @param sourceText Văn bản gốc tóm tắt
 * @param summary Kết quả tóm tắt từ LLM
 * @param model Tên model đã sử dụng
 */
export async function writeToSummariesSheet(
  sourceText: string,
  summary: string,
  model: string
): Promise<string> {
  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    let summariesSheet = worksheets.getItemOrNullObject('Summaries');
    await context.sync();

    if (summariesSheet.isNullObject) {
      // Thêm mới sheet Summaries
      summariesSheet = worksheets.add('Summaries');
      // Thiết lập dòng đầu tiên làm tiêu đề
      const headerRange = summariesSheet.getRange('A1:D1');
      headerRange.values = [['Thời gian', 'Văn bản gốc', 'Bản tóm tắt', 'Mô hình sử dụng']];
      
      // Tạo style đẹp cho dòng tiêu đề
      headerRange.format.fill.color = '#1f4e78'; // Xanh đậm chuyên nghiệp
      headerRange.format.font.color = '#ffffff'; // Chữ trắng
      headerRange.format.font.bold = true;
      headerRange.format.horizontalAlignment = 'Center';
      
      await context.sync();
    }

    // Tìm dòng trống cuối cùng để append dữ liệu mới
    const range = summariesSheet.getUsedRangeOrNullObject();
    await context.sync();

    let nextRow = 1; // 0-indexed, tức là hàng 2 (dưới tiêu đề)
    if (!range.isNullObject) {
      range.load('rowCount');
      await context.sync();
      nextRow = range.rowCount; // Số hàng đã dùng, cũng chính là index hàng trống tiếp theo
    }

    const targetRange = summariesSheet.getRangeByIndexes(nextRow, 0, 1, 4);
    const timestampStr = new Date().toLocaleString('vi-VN');
    
    // Giới hạn hiển thị preview của văn bản gốc nếu quá dài để tránh nổ bảng tính
    const truncatedSource = sourceText.length > 500 
      ? sourceText.substring(0, 500) + '...' 
      : sourceText;

    targetRange.values = [[timestampStr, truncatedSource, summary, model]];
    
    // Căn lề và wrap text
    targetRange.format.wrapText = true;
    targetRange.format.verticalAlignment = 'Top';
    
    // Auto-fit độ rộng cột vừa vặn
    summariesSheet.getUsedRange().format.autofitColumns();
    
    await context.sync();
    return 'Summaries';
  });
}
