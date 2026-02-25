import { execFile } from "node:child_process";
import { writeFile, readFile, unlink, mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { randomUUID } from "node:crypto";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const SCRIPTS_DIR = join(__dirname, "..", "scripts");
const DEFAULT_TIMEOUT = 30_000;

export interface ExecResult {
  stdout: string;
  stderr: string;
  exitCode: number;
}

/**
 * Run a Python script with the given arguments.
 */
export function runPython(
  scriptName: string,
  args: string[],
  timeout = DEFAULT_TIMEOUT
): Promise<ExecResult> {
  const scriptPath = join(SCRIPTS_DIR, scriptName);
  return new Promise((resolve, reject) => {
    execFile(
      "python",
      [scriptPath, ...args],
      { timeout, maxBuffer: 10 * 1024 * 1024 },
      (error, stdout, stderr) => {
        if (error && !("code" in error)) {
          reject(error);
          return;
        }
        resolve({
          stdout: stdout ?? "",
          stderr: stderr ?? "",
          exitCode: (error as NodeJS.ErrnoException & { code?: number })?.code
            ? 1
            : error
              ? 1
              : 0,
        });
      }
    );
  });
}

/**
 * Run a PowerShell script with the given arguments.
 */
export function runPowerShell(
  scriptName: string,
  args: string[],
  timeout = DEFAULT_TIMEOUT
): Promise<ExecResult> {
  const scriptPath = join(SCRIPTS_DIR, scriptName);
  return new Promise((resolve, reject) => {
    execFile(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        scriptPath,
        ...args,
      ],
      { timeout, maxBuffer: 10 * 1024 * 1024 },
      (error, stdout, stderr) => {
        if (error && !("code" in error)) {
          reject(error);
          return;
        }
        resolve({
          stdout: stdout ?? "",
          stderr: stderr ?? "",
          exitCode: error ? 1 : 0,
        });
      }
    );
  });
}

/**
 * Create a temporary JSON file with the given content. Returns the file path.
 */
export async function writeTempJson(content: string): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "mcp-uipath-"));
  const filePath = join(dir, `${randomUUID()}.json`);
  await writeFile(filePath, content, "utf-8");
  return filePath;
}

/**
 * Read and return the contents of a file, then delete it.
 */
export async function readAndCleanup(filePath: string): Promise<string> {
  const content = await readFile(filePath, "utf-8");
  await unlink(filePath).catch(() => {});
  return content;
}

/**
 * Delete a temp file (best-effort).
 */
export async function cleanupTemp(filePath: string): Promise<void> {
  await unlink(filePath).catch(() => {});
}

/**
 * Read a file's contents.
 */
export async function readTempFile(filePath: string): Promise<string> {
  return readFile(filePath, "utf-8");
}
