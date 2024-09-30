const reader = require("xlsx");

// Reading our test file
const file = reader.readFile("./excel1.xlsx");

let data = [];
const interval = 432000000;
const today = Date.now();

const sheets = file.SheetNames;

for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
        const day = res.lastReading.split(" ").shift();
        if (
            today - Date.parse(day.split(".").reverse().join("-")) > interval &&
            !res.object.includes("отключен") &&
            !res.object.includes("нет доступа")
        ) {
            data.push(res);
        }
    });
}

// Printing data
console.log(data);
const ws = reader.utils.json_to_sheet(data);
const newTable = reader.utils.book_new();
reader.utils.book_append_sheet(newTable, ws, "Sheet1");

// Writing to our file
reader.writeFile(newTable, "./test.xlsx");
// console.log(
//     today - Date.parse(data[1].split(".").reverse().join("-")) > interval
// );
// console.log(today);

const APIKEY = "api key";
const SECRETKEY = "secret key";

var url = "https://cleaner.dadata.ru/api/v1/clean/address";
var query = "краснодар 40 летия победы 12";

var options = {
    method: "POST",
    mode: "cors",
    headers: {
        "Content-Type": "application/json",
        Authorization: "Token " + APIKEY,
        "X-Secret": SECRETKEY,
    },
    body: JSON.stringify([query]),
};

fetch(url, options)
    .then((response) => response.text())
    .then((result) => console.log(JSON.parse(result)))
    .catch((error) => console.log("error", error));
