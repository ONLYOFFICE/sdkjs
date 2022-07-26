import "./node.js";
import "../NamesOfLiterals.js";
import "../UnicodeParser.js";
import { createRequire } from "module";
const parser = window.AscMath.CUnicodeConverter;

const require = createRequire(import.meta.url);
const fs = require("fs");

const storeData = (data, path) => {
	try {
		fs.writeFileSync(path, JSON.stringify(data, ",", 1));
	} catch (err) {
		console.error(err);
	}
};
//todo _n C_k
const ast = parser(`\\sqrt(6&1+s)^2`, undefined, true);
console.log(JSON.stringify(ast, ",", 1));
storeData(ast, "./output.json");
