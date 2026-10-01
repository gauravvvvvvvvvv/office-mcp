import assert from "node:assert/strict";
import { mkdir, mkdtemp, stat } from "node:fs/promises";
import path from "node:path";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

const root = process.cwd();
const benchmarkRoot = path.join(root, "work", "quality-benchmark");
await mkdir(benchmarkRoot, { recursive: true });
const runDirectory = await mkdtemp(path.join(benchmarkRoot, "run-"));
const asset = (name: string) => path.join(benchmarkRoot, "assets", name);
const deckPath = path.join(runDirectory, "world-domination-tactics.pptx");
const renderDirectory = path.join(runDirectory, "renders");

const client = new Client({ name: "office-quality-benchmark", version: "1.0.0" });
const serverPath = process.env.OFFICE_MCP_BENCHMARK_SERVER ?? path.join(root, "dist", "index.js");
const transport = new StdioClientTransport({ command: process.execPath, args: [serverPath], cwd: root });
async function call(name: string, args: Record<string, unknown>) {
  const result = await client.callTool({ name, arguments: args });
  assert.notEqual(result.isError, true, `${name}: ${JSON.stringify(result.content)}`);
  return result;
}

try {
  await client.connect(transport);
  await call("powerpoint_create_designed_presentation", {
    path: deckPath,
    style: "cinematic",
    slides: [
      {
        layout: "cover", title: "World Domination Tactics",
        subtitle: "A fictional playbook for influence through usefulness",
        imagePath: asset("globe.png"),
        imageAltText: "A fictional strategist studies an illuminated globe"
      },
      {
        layout: "imageText", title: "Legitimacy precedes reach",
        body: "The first win is local: make a service people choose to use. Expansion follows demonstrated value, not declarations.",
        imagePath: asset("coalition.png"), imageSide: "left",
        imageAltText: "Two fictional coalition leaders meet beneath a world map"
      },
      {
        layout: "comparison", title: "Two ways to scale influence",
        left: { heading: "Coercion", body: "Fast compliance, fragile support. Every new territory increases the cost of control." },
        right: { heading: "Coalition", body: "Slower to form, harder to dislodge. Partners extend reach while keeping local trust." },
        takeaway: "Durable reach comes from mutual value"
      },
      {
        layout: "panorama", title: "A network compounds faster than a command chain",
        body: "Shared standards let capable partners act without waiting for headquarters.",
        imagePath: asset("panorama.png"),
        imageAltText: "A fictional coastal metropolis linked by bridges and harbors at night"
      },
      {
        layout: "imageText", title: "Make the system resilient",
        body: "Distribute decisions, publish clear rules, and design graceful failure before growth makes coordination difficult.",
        imagePath: asset("network.png"), imageSide: "right",
        imageAltText: "Fictional coastal city connected by illuminated transport routes"
      }
    ]
  });
  const inspection = await call("powerpoint_inspect_presentation", { path: deckPath });
  const audit = await call("powerpoint_audit_presentation", { path: deckPath });
  assert.equal((inspection.structuredContent as { slideCount: number }).slideCount, 5);
  assert.equal((audit.structuredContent as { summary: { issueCount: number } }).summary.issueCount, 0);
  const render = await call("powerpoint_render_presentation", {
    path: deckPath, outputDirectory: renderDirectory, width: 1600, height: 900, maxSlides: 5
  });
  for (let slide = 1; slide <= 5; slide++) {
    assert.ok((await stat(path.join(renderDirectory, `Slide${slide}.png`))).size > 1000);
  }
  console.log(JSON.stringify({
    deckPath, renderDirectory,
    inspection: inspection.structuredContent,
    audit: audit.structuredContent,
    render: render.structuredContent
  }, null, 2));
} finally {
  await client.close();
}
