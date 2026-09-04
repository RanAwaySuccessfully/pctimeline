"use strict";

/* DATA */

let _worker = new Worker("pctimeline-xlsx.js");
_worker.onmessage = event => {
    switch (event.data.op) {
        case "load":
            loadedExcel(event.data.result);
            break;
        case "download":
            exportedExcel(event.data.result);
            break;
    }
};

let _vega;
let _vegaConfig;

let _components = [];
let _categories = [];
let _computers = [];
let _totals = {
    "oldest": Date.now(),
    "newest": Date.now()
};

/* UTILS */

function stringToDate(string) {
    if (!string) {
        throw "This function should never receive a non-string value. Do proper validation before a value reaches here!!";
    }

    const match = string.match(/(\d+)\/(\d+)/);
    return new Date(`${ match[2] }-${ match[ 1 ]}-01`);
}

function separatingTime(t) {
    var ms = t % 1000;
    ms = parseInt(ms);
    t = Math.floor(t / 1000);
    var s = t % 60;
    t = Math.floor(t / 60);
    var min = t % 60;
    t = Math.floor(t / 60);
    var h = t % 24;
    t = Math.floor(t / 24);
    var d = t % 30;
    t = Math.floor(t / 30);
    var m = t % 12;
    var y = Math.floor(t / 12);
    var r = y + " years, " + m + " months";
    return r;
}

function formatData() {
    const rows = _components.map(row => {

        const notes = row.notes || "(none)";

        let status = row.status;
        if (!status) {
            status = row.end ? "(unknown)" : "In use";
        }

        let formatted_price = "(unknown)";
        if (row.price) {
            formatted_price = `${ row.currency || "" } ${ row.price.toFixed(2) }`;
        }

        let formatted_new_purchase = row.new_purchase ? "Yes" : "No";

        // DATES

        let start = row.start ? stringToDate(row.start) : new Date(_totals.oldest);
        let end = row.end ? stringToDate(row.end) : new Date();

        let formatted_start = "(unknown)";
        if (row.start) {
            formatted_start = row.start;
        }

        let formatted_end = "Now";
        if (row.end) {
            formatted_end = row.end;
        }

        const diff = end.getTime() - start.getTime();
        formatted_end += ` (${ separatingTime(diff) })`;

        // Y-AXIS INDEX

        let category_index = _categories.indexOf(row.category);
        if (category_index === -1) {
            category_index = 999;
        }

        category_index *= 10000;
        category_index += start.getFullYear();
        category_index *= 100;
        category_index += start.getMonth();
        
        return {
            ...row,
            notes,
            status,
            formatted_price,
            formatted_new_purchase,
            start,
            end,
            formatted_start,
            formatted_end,
            category_index,
        };
    });

    return rows;
}

function setOldest() {
    _components.forEach(row => {
        if (!row.start) {
            return;
        }

        let start = stringToDate(row.start).getTime();

        if (start < _totals.oldest) {
            _totals.oldest = start;
        }
    });
}

async function renderVega() {
    setOldest();
    const data = formatData();
    
    if (!_vega) {
        _vegaConfig.data.values = data;

        _vega = await window.vegaEmbed("#vis", _vegaConfig, {
            renderer: "svg",
            theme: "dark",
            /*
            patch: (spec) => {
                debugger;
                return spec;
            },
            */
            tooltip: {
                theme: "dark",
                /*
                formatTooltip: (obj, sanitize) => {
                    const string = [];
                    for (let key in obj) {
                        let value = obj[key];
                        key = sanitize(key);
                        value = sanitize(value);

                        if (value && (key === "link")) {
                            value = `<a href="${ value }">${ value }</a>`;
                        }

                        string.push(`<tr><td class="key">${ key }</td><td class="value">${ value }</td></tr>`);
                    }

                    return `<table><tbody>${ string.join("\n") }</tbody></table>`;
                }
                */
            }
        });

        const labels = document.getElementsByClassName("mark-text");
        for (const element of labels) {
            element.removeAttribute("pointer-events");
        }

        const axis = labels[1].children;
        for (const element of axis) {
            const name = element.innerHTML;
            const component = data.find(component => component.name === name);

            if (component && component.link) {
                const safeLink = escapeHtml(component.link);
                element.innerHTML = `<a href="${ safeLink }">${ element.innerHTML }</a>`;
                element.setAttribute("text-decoration", "underline");
            }
        }

    } else {
        /*
        let changeSet = window.vega.changeset()
            .remove(() => true)
            .insert(formatData);
        */

       _vega.view.data("main", data).run();
    }
}

function escapeHtml(string) {
    if (!string) {
        return string;
    }

    return string
        .replaceAll('&', '&amp')
        .replaceAll('<', '&lt')
        .replaceAll('>', '&gt;')
        .replaceAll("'", '&#39;')
        .replaceAll('"', '&quot;');
}

/* EVENTS */

function hoverInfo(event, element) {
    let px = (event.clientX - element.getBoundingClientRect().left);
    const container = document.getElementById("vis");
    const chart = container.getElementsByClassName("background")[0];

    const container_box = container.getBoundingClientRect();
    const chart_box = chart.getBoundingClientRect();

    const min = chart_box.left - container_box.left;
    const max = min + chart_box.width;

    const column = document.getElementById("infocolumn");

    if ((px < min) || (px > max)) {
        column.classList.add("invisible");
        return;
    }

    column.style.left = px + "px";
    column.classList.remove("invisible");

    const timePerPX = (_totals.newest - _totals.oldest) / (max - min);
    const currentTime = ((px - min) * timePerPX) + _totals.oldest;

    let currency;

    const selection = _vega.view.signal("price_selection");
    const selectionIndex = selection?._vgsid_;
    const components = _vega.view.data("main");
    
    const totalPrice = components.reduce((totalPrice, row, index) => {
        if (row.currency) {
            currency = row.currency;
        }

        if (!row.price) {
            return totalPrice;
        }

        if (selectionIndex && !selectionIndex.has(index + 1)) {
            return totalPrice;
        }

        const start = row.start.getTime();
        const end = row.end.getTime();

        if ((currentTime >= start) && (currentTime <= end)) {
            totalPrice += row.price;
        }

        return totalPrice;
    }, 0);

    const priceElement = document.getElementById("totalprice");
    priceElement.innerHTML = `${ currency } ${ totalPrice.toFixed(2) }`;

    let currentPC = _computers.find(computer => {
        let start = computer.start ? stringToDate(computer.start).getTime() : _totals.oldest;
        let end = computer.end ? stringToDate(computer.end).getTime() : Date.now();

        if ((currentTime >= start) && (currentTime <= end)) {
            return true;
        }

        return false;
    });

    const computerElement = document.getElementById("computername");
    computerElement.innerHTML = currentPC ? currentPC.name : "-";
}

/*
function changeTitle() {
    let title = window.prompt("Type a new title below:");
    renderVega();
}
*/

function customPC(element) {
    setActiveButton(element);

    const customButtons = document.getElementById("custombuttons");
    customButtons.style = "";
}

async function loadPC(element, fileName) {
    const response = await fetch(fileName);
    const fileData = await response.arrayBuffer();
    loadExcel(fileData);

    setActiveButton(element);
    const customButtons = document.getElementById("custombuttons");
    customButtons.style = "display: none;";
}

function setActiveButton(element) {
    const div = document.getElementById("mainbuttons");
    const children = div.children;
    Array.from(children).forEach(child => {
        if (child === element) {
            child.classList.add("active");
        } else {
            child.classList.remove("active");
        }
    });
}

async function importExcel(element) {
    const file = element.files[0];
    const fileData = await file.arrayBuffer();
    //const title = file.name.replace(".xlsx", "");

    loadExcel(fileData);
}

function exportExcel() {
    _worker.postMessage({
        "op": "download",
        "json": {
            "components": _components,
            "categories": _categories,
            "computers": _computers,
        }
    });
}

function exportedExcel(fileData) {
    const blob = new Blob([ fileData ], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    
    link.href = url;
    link.download = "PCTimeline.xlsx";
    link.click();

    window.URL.revokeObjectURL(url);
}

function loadExcel(fileData) {
    _worker.postMessage({
        "op": "load",
        "file": fileData
    });
}

function loadedExcel(result) {
    if (result.error) {
        alert(result.error);
    } else {
        _categories = result.categories;
        _components = result.components;
        _computers = result.computers;
        
        renderVega();
    }
}

async function init() {
    const vegaResponse = await fetch("vega-lite.json");
    _vegaConfig = await vegaResponse.json();
    const response = await fetch("PCTimeline.xlsx");
    const fileData = await response.arrayBuffer();
    loadExcel(fileData);
};

init();