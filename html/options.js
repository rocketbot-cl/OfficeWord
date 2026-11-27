let lineType = document.getElementById("lineType");
let lineSize = document.getElementById("lineSize");

let sizes = [
    "1/4 point",
    "1/2 point",
    "3/4 point",
    "1 point",
    "1 1/2 points",
    "2 1/4 points",
    "3 points",
    "4 1/2 points",
    "6 points"
];

let valuesSizes = [
    "25_point",
    "50_point",
    "75_point",
    "100_point",
    "150_points",
    "225_points",
    "300_points",
    "450_points",
    "600_points"
];

function showSizes(array, arrayValues) {
    if (!lineSize) {
        return;
    }

    while (lineSize.length > 0) {
        lineSize.remove(0);
    }

    for (let i = 0; i < array.length; i++) {
        let option = document.createElement("option");
        option.value = arrayValues[i];
        option.text = array[i];
        lineSize.appendChild(option);
    }
}

window.ensureLineSizeOption = function (sizeValue) {
    if (!lineSize || !sizeValue) {
        return;
    }
    let idx = valuesSizes.indexOf(sizeValue);
    let label = idx !== -1 ? sizes[idx] : sizeValue;
    let option = document.createElement("option");
    option.value = sizeValue;
    option.text = label;
    lineSize.appendChild(option);
};

window.applyLineTypeFilter = function (valor) {
    if (!lineType || !lineSize) {
        return;
    }

    let previousSize = lineSize.value;
    let cut = [];
    let cutSizes = [];

    switch (valor) {
        case "thin_thick_thin_med_gap":
            cut = sizes.slice(0, 8);
            cutSizes = valuesSizes.slice(0, 8);
            showSizes(cut, cutSizes);
            break;
        case "single_wavy":
            cut = ["3/4 point", "1 1/2 points"];
            cutSizes = ["75_point", "150_points"];
            showSizes(cut, cutSizes);
            break;
        case "double_wavy":
            cut = ["3/4 point"];
            cutSizes = ["75_point"];
            showSizes(cut, cutSizes);
            break;
        case "dash_dot_stroked":
            cut = ["3 points"];
            cutSizes = ["300_points"];
            showSizes(cut, cutSizes);
            break;
        case "double":
        case "triple":
            cut = sizes.slice(0, 7);
            cutSizes = valuesSizes.slice(0, 7);
            showSizes(cut, cutSizes);
            break;
        default:
            showSizes(sizes, valuesSizes);
            break;
    }

    // Conserva el tamaño anterior si sigue siendo válido para el nuevo tipo de línea
    if (previousSize && Array.from(lineSize.options).some(opt => opt.value === previousSize)) {
        lineSize.value = previousSize;
    } else if (lineSize.options.length > 0) {
        lineSize.selectedIndex = 0;
    }
};

function syncAndSend() {
    if (!lineType || !lineSize || typeof message === "undefined") {
        return;
    }
    message.commands["lineType"] = lineType.value;
    message.commands["lineSize"] = lineSize.value;
    if (typeof SendMessage === "function") {
        SendMessage();
    }
}

if (lineType) {
    lineType.addEventListener("change", function () {
        window.applyLineTypeFilter(lineType.value);
        syncAndSend();
    });
}

if (lineSize) {
    lineSize.addEventListener("change", function () {
        syncAndSend();
    });
}

// Sincroniza estado inicial al cargar
if (lineType && lineSize) {
    window.applyLineTypeFilter(lineType.value);
    lineSize.value = "100_point";
    syncAndSend();
}