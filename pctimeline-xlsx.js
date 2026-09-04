"use strict";

importScripts("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js");

onmessage = event => {
    switch (event.data.op) {
        case "load":
            event.data.result = loadExcel(event.data.file);
            postMessage(event.data);
            break;
        case "download":
            event.data.result = exportExcel(event.data.json);
            postMessage(event.data);
            break;
    }
};

function loadExcel(fileData) {
    const workbook = XLSX.read(fileData);

    const wsComponents = workbook.Sheets.Components;
    const wsCategories = workbook.Sheets.Categories;
    const wsComputers = workbook.Sheets.Computers;

    if (!wsComponents || !wsCategories) {
        return {
            error: "File must contain both 'Components' and 'Categories' worksheets."
        };
    }

    let components = XLSX.utils.sheet_to_json(wsComponents, {});
    let categories = XLSX.utils.sheet_to_json(wsCategories, {});
    categories = categories.map(category => category.Name);

    let computers = [];
    if (wsComputers) {
        computers = XLSX.utils.sheet_to_json(wsComputers, {});
    }

    return {
        error: null,
        components,
        categories,
        computers,
    };
}

function exportExcel(json) {
    const wsComponents = XLSX.utils.json_to_sheet(json.components, {
        headers: ["category", "name", "price", "condition", "currency", "start", "end", "new_purchase", "notes"]
    });

    const categories = json.categories.map(category => {
        return {
            "Name": category
        };
    });

    const wsCategories = XLSX.utils.json_to_sheet(categories, {
        headers: ["Name"]
    });

    const wsComputers = XLSX.utils.json_to_sheet(json.computers, {
        headers: ["name", "start", "end"]
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, wsComponents, "Components");
    XLSX.utils.book_append_sheet(workbook, wsCategories, "Categories");
    XLSX.utils.book_append_sheet(workbook, wsComputers, "Computers");
    
    const file = XLSX.write(workbook, {
        type: "buffer"
    });

    return file;
}