window.InitUserScripts = function()
{
var player = GetPlayer();
var object = player.object;
var addToTimeline = player.addToTimeline;
var setVar = player.SetVar;
var getVar = player.GetVar;
window.Script1 = function()
{
  (function() {
    console.log("🚀 Excel-export script gestart...");

    function loadSheetJS(callback) {
        if (typeof window.XLSX === "undefined") {
            var script = document.createElement("script");
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
            script.onload = callback;
            document.head.appendChild(script);
        } else {
            callback();
        }
    }

    loadSheetJS(function() {
        generateExcel_Smartkit();
    });

    function generateExcel_Smartkit() {
        try {
            var player = GetPlayer();

            // Haal de Storyline-variabelen op
            var data = [
                ["Variabele", "Waarde"],
                ["Export Datum", new Date().toLocaleDateString("nl-NL")],
                ["Export Tijd", new Date().toLocaleTimeString("nl-NL")],
            ];

            var variabelen = ["doelgroep", "missie", "visie", "aanbod", "boodschap", "buyerpersona", "website", "sem", "socials", "rapportering", "sales", "structuur", "it", "kpi", "hr", "strategie"];

            variabelen.forEach(function(v) {
                data.push([v, player.GetVar(v) || "Niet ingevuld"]);
            });

            var wb = XLSX.utils.book_new();
            var ws = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(wb, ws, "MetaFrame Data");

            XLSX.writeFile(wb, "MetaFrame_" + new Date().toLocaleDateString("nl-NL") + ".xlsx");
            alert("✅ Excel-bestand succesvol geëxporteerd!");
        } catch (error) {
            console.error("❌ Fout bij Excel-export:", error);
            alert("⚠ Er is een fout opgetreden bij het exporteren naar Excel.");
        }
    }
})();

}

window.Script2 = function()
{
  (function() {
    console.log("🚀 Excel-import script gestart...");

    function loadSheetJS(callback) {
        if (typeof window.XLSX === "undefined") {
            var script = document.createElement("script");
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
            script.onload = callback;
            document.head.appendChild(script);
        } else {
            callback();
        }
    }

    loadSheetJS(function() {
        importExcel_Smartkit();
    });

    function importExcel_Smartkit() {
        try {
            if (!document.getElementById("excelUpload")) {
                var input = document.createElement("input");
                input.type = "file";
                input.id = "excelUpload";
                input.accept = ".xlsx,.xls";
                input.style.display = "none";
                document.body.appendChild(input);

                input.addEventListener("change", function(event) {
                    var file = event.target.files[0];
                    if (file) {
                        readExcelFile(file);
                    }
                });
            }

            document.getElementById("excelUpload").click();

            function readExcelFile(file) {
                var reader = new FileReader();
                reader.onload = function(event) {
                    var workbook = XLSX.read(event.target.result, { type: "array" });
                    var worksheet = workbook.Sheets["MetaFrame Data"];
                    var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    var player = GetPlayer();
                    var updated = 0;

                    jsonData.forEach(function(row) {
                        if (row[0] && row[1] !== undefined) {
                            player.SetVar(row[0], row[1]);
                            updated++;
                        }
                    });

                    alert(updated > 0 ? "✅ Excel-data succesvol geïmporteerd!" : "⚠ Geen geldige data gevonden in het Excel-bestand.");
                };

                reader.onerror = function() {
                    alert("❌ Fout bij het lezen van het Excel-bestand. Probeer opnieuw.");
                };

                reader.readAsArrayBuffer(file);
            }
        } catch (error) {
            console.error("❌ Fout bij Excel-import:", error);
            alert("⚠ Er is een fout opgetreden bij het importeren van het Excel-bestand.");
        }
    }
})();

}

window.Script3 = function()
{
  var player = GetPlayer(); // Haal de Storyline speler op

var totaal = 0; // Start met 0

// Lijst van specifieke variabelen die je wilt optellen
var variabelen = ["sales"];

// Loop door de lijst en tel alleen deze variabelen op
for (var i = 0; i < variabelen.length; i++) {
    totaal += player.GetVar(variabelen[i]);
}

// Bereken het gemiddelde en rond af
var percentage = Math.round(totaal / variabelen.length);

// Zet de waarden terug in Storyline variabelen
player.SetVar("Sales_total", totaal);
player.SetVar("PercentageSales_total", percentage);
}

window.Script4 = function()
{
  var player = GetPlayer(); // Haal de Storyline speler op

var totaal = 0; // Start met 0

// Lijst van specifieke variabelen die je wilt optellen
var variabelen = ["hr", "kpi", "it", "structuur"];

// Loop door de lijst en tel alleen deze variabelen op
for (var i = 0; i < variabelen.length; i++) {
    totaal += player.GetVar(variabelen[i]);
}

// Bereken het gemiddelde en rond af
var percentage = Math.round(totaal / variabelen.length);

// Zet de waarden terug in Storyline variabelen
player.SetVar("organisatie", totaal);
player.SetVar("PercentageOrganisatie", percentage);
}

window.Script5 = function()
{
  
var player = GetPlayer(); // Haal de Storyline speler op

var totaal = 0; // Start met 0

// Lijst van specifieke variabelen die je wilt optellen
var variabelen = ["boodschap", "buyerpersona", "rapportering", "sem", "socials", "website"];

// Loop door de lijst en tel alleen deze variabelen op
for (var i = 0; i < variabelen.length; i++) {
    totaal += player.GetVar(variabelen[i]);
}

// Bereken het gemiddelde en rond af
var percentage = Math.round(totaal / variabelen.length);

// Zet de waarden terug in Storyline variabelen
player.SetVar("Lead_generation", totaal);
player.SetVar("PercentageLead_generation", percentage);

}

window.Script6 = function()
{
  var player = GetPlayer(); // Haal de Storyline speler op

var totaal = 0; // Start met 0

// Lijst van specifieke variabelen die je wilt optellen
var variabelen = ["aanbod", "doelgroep", "missie", "strategie", "visie"];

// Loop door de lijst en tel alleen deze variabelen op
for (var i = 0; i < variabelen.length; i++) {
    totaal += player.GetVar(variabelen[i]);
}

// Bereken het gemiddelde en rond af
var percentage = Math.round(totaal / variabelen.length);

// Zet de waarden terug in Storyline variabelen
player.SetVar("misie_visie", totaal);
player.SetVar("PercentageMisie_visie", percentage);

}

window.Script7 = function()
{
  var player = GetPlayer(); // Haal de Storyline speler op

var totaal = 0; // Start met 0

// Lijst van specifieke variabelen die je wilt optellen
var variabelen = ["aanbod", "boodschap", "buyerpersona", "doelgroep", "hr", "kpi", "it", "missie", "rapportering", "sales", "sem", "socials", "strategie", "structuur", "visie", "website"];

// Loop door de lijst en tel alleen deze variabelen op
for (var i = 0; i < variabelen.length; i++) {
    totaal += player.GetVar(variabelen[i]);
}

// Bereken het gemiddelde en rond af
var percentage = Math.round(totaal / variabelen.length);

// Zet de waarden terug in Storyline variabelen
player.SetVar("Metaframe", totaal);
player.SetVar("PercentageMetaframe", percentage);

}

window.Script8 = function()
{
  var player = GetPlayer(); // Haal de Storyline speler op

var totaal = 0; // Start met 0

// Lijst van specifieke variabelen die je wilt optellen
var variabelen = ["Aanbod", "Boodschap", "Buyerpersona", "Doelgroep", "HR", "KPI", "IT", "Missie", 
                  "Rapportering", "Sales", "SEM", "Socials", "Strategie", "Structuur", "Visie", "Website"];

// Loop door de lijst en tel alleen deze variabelen op
for (var i = 0; i < variabelen.length; i++) {
    totaal += player.GetVar(variabelen[i]);
}

// Bereken het gemiddelde en rond af
var percentage = Math.round(totaal / variabelen.length);

// Zet de waarden terug in Storyline variabelen
player.SetVar("Metaframe", totaal);
player.SetVar("PercentageMetaframe", percentage);

}

};
