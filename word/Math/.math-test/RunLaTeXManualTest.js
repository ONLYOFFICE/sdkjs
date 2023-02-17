import "./node.js";
import "../NamesOfLiterals.js";
import  "../LaTeXParser.js";
import { createRequire } from "module";

const parser = window.AscMath.ConvertLaTeXToTokensList;
const require = createRequire(import.meta.url);
const fs = require("fs");
const storeData = (data, path) => {
	try {
		fs.writeFileSync(path, JSON.stringify(data, ",", 1));
	} catch (err) {
		console.error(err);
	}
};
const ast = parser(`\\frac{1}{2}`, undefined, true);
console.log(JSON.stringify(ast, ",", 1));
storeData(ast, "./output.json");
//todo "\\vec \\frac{k}{2}