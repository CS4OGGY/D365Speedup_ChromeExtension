const { minify } = require("terser");
const fs = require("fs");
const path = require("path");

const files = fs.readdirSync("dist").filter(f => f.endsWith(".js"));

(async () => {
    for (const file of files) {
        const filePath = path.join("dist", file);
        const code = fs.readFileSync(filePath, "utf-8");
        const result = await minify(code, { compress: true, mangle: true });
        fs.writeFileSync(filePath, result.code);
        console.log(`Minified: ${filePath}`);
    }
})();
