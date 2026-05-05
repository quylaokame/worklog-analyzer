
export function loadCSV(file, options = {}) {
    return new Promise((resolve, reject) => {
        if (!file) {
            reject(new Error("No file provided"));
            return;
        }
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            encoding: "utf-8",
            ...options,
            complete: (results) => {
                resolve(results.data);
            },
            error: (error) => {
                reject(error);
            },
        });
    });
}

export function loadJSON(file) {
    if (!file) {
        console.error("No file provided");
        return;
    }
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            try {
                const json = JSON.parse(reader.result);
                resolve(json);
            } catch (err) {
                reject(new Error("Invalid JSON file"));
            }
        };
        reader.onerror = () => reject(new Error("Failed to read file"));
        reader.readAsText(file, "utf-8");
    });
}

