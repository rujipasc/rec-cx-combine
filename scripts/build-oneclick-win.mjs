import { copyFile, mkdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const root = process.cwd();
const distDir = path.join(root, "dist");
const exeName = "rec-cx-combine.exe";
const exePath = path.join(distDir, exeName);
const outDir = path.join(distDir, "oneclick-win-x64");

const runBat = `@echo off
setlocal
cd /d "%~dp0"

if not exist ".\\input" mkdir ".\\input"
if not exist ".\\output" mkdir ".\\output"

echo ================================
echo  CardX Recruitment Combine Tool
echo ================================
echo Input folder : .\\input
echo Output file  : .\\output\\recruitment-tracking.xlsx
echo.

".\\${exeName}" --in ".\\input" --out ".\\output\\recruitment-tracking.xlsx"

echo.
echo Finished. Press any key to close.
pause >nul
endlocal
`;

const readme = `CardX Recruitment Combine - One Click (Windows x64)

How to use:
1) Put source Excel files in "input" folder.
   - Candidate file must include column: รหัสบัตรประชาชน
   - JR file must include column: รหัสใบร้องขอ/ID or JR No.
2) Double-click "run.bat"
3) Output will be generated at "output\\recruitment-tracking.xlsx"
`;

async function main() {
  await rm(outDir, { recursive: true, force: true });
  await mkdir(path.join(outDir, "input"), { recursive: true });
  await mkdir(path.join(outDir, "output"), { recursive: true });

  await copyFile(exePath, path.join(outDir, exeName));
  await writeFile(path.join(outDir, "run.bat"), runBat, "utf8");
  await writeFile(path.join(outDir, "README.txt"), readme, "utf8");

  console.log("One-click package created:");
  console.log(outDir);
}

main().catch((err) => {
  console.error("Failed to prepare one-click package:", err?.message || err);
  process.exit(1);
});
