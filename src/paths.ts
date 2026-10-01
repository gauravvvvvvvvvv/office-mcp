import path from "node:path";
import { access, mkdir, realpath } from "node:fs/promises";
import { constants, realpathSync } from "node:fs";
import { OfficeMcpError } from "./errors.js";

const configuredRoots = (process.env.OFFICE_MCP_ROOTS ?? process.cwd())
  .split(path.delimiter)
  .filter(Boolean)
  .map((root) => {
    const resolved = path.resolve(root);
    try {
      return realpathSync.native(resolved);
    } catch {
      return resolved;
    }
  });

function isWithinRoot(candidate: string, root: string): boolean {
  const relative = path.relative(root, candidate);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function assertAllowed(candidate: string, original: string): void {
  if (!configuredRoots.some((root) => isWithinRoot(candidate, root))) {
    throw new OfficeMcpError(
      `Path is outside the configured Office MCP roots: ${original}`,
      "PATH_NOT_ALLOWED",
      { allowedRoots: configuredRoots }
    );
  }
}

export function getAllowedRoots(): string[] {
  return [...configuredRoots];
}

export function resolveAllowedPath(inputPath: string, extensions?: string[]): string {
  const resolved = path.resolve(inputPath);
  assertAllowed(resolved, inputPath);

  if (extensions && !extensions.includes(path.extname(resolved).toLowerCase())) {
    throw new OfficeMcpError(
      `Unsupported file extension: ${path.extname(resolved) || "(none)"}`,
      "UNSUPPORTED_EXTENSION",
      { supportedExtensions: extensions }
    );
  }

  return resolved;
}

export async function resolveReadablePath(inputPath: string, extensions?: string[]): Promise<string> {
  const resolved = resolveAllowedPath(inputPath, extensions);

  try {
    await access(resolved, constants.R_OK);
  } catch {
    throw new OfficeMcpError(`File cannot be read: ${inputPath}`, "FILE_NOT_FOUND");
  }
  const canonical = await realpath(resolved);
  assertAllowed(canonical, inputPath);
  return canonical;
}

export async function prepareOutputPath(
  inputPath: string,
  extensions: string[],
  overwrite = false
): Promise<string> {
  const resolved = resolveAllowedPath(inputPath, extensions);
  await mkdir(path.dirname(resolved), { recursive: true });
  const canonicalParent = await realpath(path.dirname(resolved));
  const canonicalTarget = path.join(canonicalParent, path.basename(resolved));
  assertAllowed(canonicalTarget, inputPath);

  if (!overwrite) {
    try {
      await access(canonicalTarget, constants.F_OK);
      throw new OfficeMcpError(
        `Output already exists: ${inputPath}. Set overwrite to true to replace it.`,
        "FILE_EXISTS"
      );
    } catch (error) {
      if (error instanceof OfficeMcpError) throw error;
    }
  }
  if (overwrite) {
    try {
      const existingCanonical = await realpath(canonicalTarget);
      assertAllowed(existingCanonical, inputPath);
    } catch (error) {
      if (error instanceof OfficeMcpError) throw error;
    }
  }

  return canonicalTarget;
}

export function resolveInputOutput(
  inputPath: string,
  outputPath: string | undefined,
  extensions: string[]
): { input: string; output: string } {
  return {
    input: resolveAllowedPath(inputPath, extensions),
    output: resolveAllowedPath(outputPath ?? inputPath, extensions)
  };
}
