import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";
import { mkdir, readFile, readdir, unlink } from "node:fs/promises";
import { OfficeMcpError } from "./errors.js";
import { prepareOutputPath, resolveAllowedPath, resolveReadablePath } from "./paths.js";

type Application = "excel" | "word" | "powerpoint";
type BridgeCommand = "status" | "list-open" | "excel-batch" | "word-batch" | "powerpoint-batch";

interface NativeTarget {
  path?: string;
  active?: boolean;
  create?: boolean;
  visible?: boolean;
}

export interface NativeBatchRequest {
  target: NativeTarget;
  operations: Array<Record<string, unknown> & { op: string }>;
  overwrite?: boolean;
}

const extensionByApplication: Record<Application, string[]> = {
  excel: [".xlsx", ".xlsm", ".xlsb", ".xls"],
  word: [".docx", ".docm", ".doc", ".rtf"],
  powerpoint: [".pptx", ".pptm", ".ppt", ".ppsx"]
};

function powershellPath(): string {
  if (process.platform !== "win32") {
    throw new OfficeMcpError("Native Office automation is available only on Windows", "NATIVE_UNAVAILABLE");
  }
  return path.join(process.env.SystemRoot ?? "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
}

function bridgePath(): string {
  return fileURLToPath(new URL("../native/office-bridge.ps1", import.meta.url));
}

async function runBridge(command: BridgeCommand, payload: unknown = {}, timeoutMs = 180000) {
  const child = spawn(
    powershellPath(),
    ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", bridgePath(), "-Command", command],
    { windowsHide: true, stdio: ["pipe", "pipe", "pipe"] }
  );

  const stdout: Buffer[] = [];
  const stderr: Buffer[] = [];
  child.stdout.on("data", (chunk: Buffer) => stdout.push(chunk));
  child.stderr.on("data", (chunk: Buffer) => stderr.push(chunk));
  child.stdin.end(JSON.stringify(payload));

  const result = await new Promise<{ code: number | null; signal: NodeJS.Signals | null }>((resolve, reject) => {
    const timer = setTimeout(() => {
      child.kill();
      reject(new OfficeMcpError(`Native Office command timed out: ${command}`, "NATIVE_TIMEOUT"));
    }, timeoutMs);
    child.once("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.once("exit", (code, signal) => {
      clearTimeout(timer);
      resolve({ code, signal });
    });
  });

  const output = Buffer.concat(stdout).toString("utf8").trim();
  const errorOutput = Buffer.concat(stderr).toString("utf8").trim();
  let envelope: { ok?: boolean; result?: unknown; error?: { message?: string; [key: string]: unknown } };
  try {
    envelope = JSON.parse(output) as typeof envelope;
  } catch {
    throw new OfficeMcpError("Native Office bridge returned invalid output", "NATIVE_PROTOCOL_ERROR", {
      exitCode: result.code,
      signal: result.signal,
      stdout: output.slice(0, 4000),
      stderr: errorOutput.slice(0, 4000)
    });
  }

  if (!envelope.ok || result.code !== 0) {
    throw new OfficeMcpError(
      envelope.error?.message ?? (errorOutput || `Native Office command failed: ${command}`),
      "NATIVE_OPERATION_FAILED",
      { exitCode: result.code, stderr: errorOutput.slice(0, 4000), bridge: envelope.error }
    );
  }
  return envelope.result;
}

async function secureBatchPaths(application: Application, request: NativeBatchRequest): Promise<NativeBatchRequest> {
  const target = { ...request.target };
  if (target.path) target.path = await resolveReadablePath(target.path, extensionByApplication[application]);

  const operations: NativeBatchRequest["operations"] = [];
  for (const original of request.operations) {
    const operation = { ...original };
    if (typeof operation.path === "string") {
      operation.path = await resolveReadablePath(operation.path, operation.op === "add_media"
        ? [".mp4", ".mov", ".wmv", ".avi", ".mp3", ".wav", ".m4a"]
        : [".png", ".jpg", ".jpeg", ".gif", ".svg", ".emf"]);
    }
    if (typeof operation.outputPath === "string") {
      const extensions = operation.op === "export_pdf" ? [".pdf"] : extensionByApplication[application];
      operation.outputPath = await prepareOutputPath(operation.outputPath, extensions, request.overwrite ?? false);
    }
    if (typeof operation.outputDirectory === "string") {
      const directory = resolveAllowedPath(operation.outputDirectory);
      await mkdir(directory, { recursive: true });
      operation.outputDirectory = directory;
    }
    operations.push(operation);
  }
  return { ...request, target, operations };
}

export async function nativeStatus() {
  return runBridge("status");
}

export async function listOpenOfficeFiles() {
  return runBridge("list-open");
}

export async function runNativeBatch(application: Application, request: NativeBatchRequest) {
  const secured = await secureBatchPaths(application, request);
  return runBridge(`${application}-batch` as BridgeCommand, secured);
}

export async function renderPowerPointNative(
  filePath: string,
  outputDirectory: string,
  width = 1600,
  height = 900,
  overwrite = false
) {
  const directory = resolveAllowedPath(outputDirectory);
  await mkdir(directory, { recursive: true });
  const oldRenders = (await readdir(directory)).filter((name) => /^Slide\d+\.(png|jpe?g)$/i.test(name));
  if (oldRenders.length && !overwrite) {
    throw new OfficeMcpError(
      `Render directory already contains slide images: ${directory}`,
      "FILE_EXISTS",
      { files: oldRenders }
    );
  }
  if (overwrite) {
    for (const name of oldRenders) await unlink(path.join(directory, name));
  }

  await runNativeBatch("powerpoint", {
    target: { path: filePath },
    operations: [{ op: "render_slides", outputDirectory: directory, width, height }]
  });

  const imagePaths = (await readdir(directory))
    .filter((name) => /^Slide\d+\.png$/i.test(name))
    .sort((left, right) =>
      Number(/\d+/.exec(left)?.[0] ?? 0) - Number(/\d+/.exec(right)?.[0] ?? 0)
    )
    .map((name) => path.join(directory, name));
  return {
    outputDirectory: directory,
    images: await Promise.all(
      imagePaths.map(async (imagePath) => ({
        path: imagePath,
        data: (await readFile(imagePath)).toString("base64")
      }))
    )
  };
}
