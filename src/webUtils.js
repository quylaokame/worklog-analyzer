/**
 * Hàm nhập file JSON từ máy tính
 * @returns {Promise<any|null>}
 */
export const importJsonFile = () => {
    return new Promise((resolve, reject) => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'application/json';

        input.onchange = (event) => {
            const file = event.target.files?.[0];

            if (!file) {
                resolve(null);
                return;
            }

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const content = e.target.result;
                    const json = JSON.parse(content);
                    resolve(json);
                } catch (error) {
                    reject(new Error("File không đúng định dạng JSON chuẩn."));
                }
            };
            reader.onerror = () => reject(new Error("Lỗi khi đọc file."));
            reader.readAsText(file);
        };

        input.click();
    });
};

/**
 * Hàm xuất dữ liệu ra file JSON và tải về
 * @param {any} data - Dữ liệu cần xuất
 * @param {string} fileName - Tên file mặc định là data.json
 */
export const exportJsonFile = (data, fileName = 'data.json') => {
    try {
        const jsonString = JSON.stringify(data, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json' });
        const url = URL.createObjectURL(blob);

        const link = document.createElement('a');
        link.href = url;
        link.download = fileName.endsWith('.json') ? fileName : `${fileName}.json`;

        document.body.appendChild(link);
        link.click();

        // Dọn dẹp sau khi click
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    } catch (error) {
        console.error("Không thể xuất file:", error);
    }
};

export const exportData = (data, type, fileName) => {
    try {
        const jsonString = JSON.stringify(data, null, 2);
        const blob = new Blob([jsonString], { type });
        const url = URL.createObjectURL(blob);

        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;

        document.body.appendChild(link);
        link.click();

        // Dọn dẹp sau khi click
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        console.log("Export file success", error);
    } catch (error) {
        console.error("Không thể xuất file:", error);
    }
};