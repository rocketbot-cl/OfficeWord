var data_;

var message = {
    type: "iframe",
    commands: {}
};

var SendMessage = function () {
    if (window.parent && window.parent !== window) {
        window.parent.postMessage(message, "*");
    }
};

function dataHandler(e) {
    let lineTypeEl = document.getElementById("lineType");
    let lineSizeEl = document.getElementById("lineSize");

    if (!e || !e.data || !lineTypeEl || !lineSizeEl) {
        return;
    }

    // Soporta tanto objeto directo { lineType, lineSize } como { commands: { lineType, lineSize } }
    let payload = e.data.commands && typeof e.data.commands === "object"
        ? e.data.commands
        : e.data;

    let incomingType = payload.lineType;
    let incomingSize = payload.lineSize;

    if (incomingType) {
        lineTypeEl.value = incomingType;
        if (typeof window.applyLineTypeFilter === "function") {
            window.applyLineTypeFilter(incomingType);
        }
    }

    if (incomingSize) {
        // Si el tamaño no existe en el <select> filtrado actual, lo agrega antes de asignarlo
        let exists = Array.from(lineSizeEl.options).some(opt => opt.value === incomingSize);
        if (!exists && typeof window.ensureLineSizeOption === "function") {
            window.ensureLineSizeOption(incomingSize);
        }
        lineSizeEl.value = incomingSize;
    }

    message.commands["lineType"] = lineTypeEl.value;
    message.commands["lineSize"] = lineSizeEl.value;
}

var eventMethod = window.addEventListener ? "addEventListener" : "attachEvent";
var eventer = window[eventMethod];
var messageEvent = eventMethod == "attachEvent" ? "onmessage" : "message";

eventer(messageEvent, dataHandler);