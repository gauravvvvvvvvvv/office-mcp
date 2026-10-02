---
name: office-presentation-quality
description: Create or improve PowerPoint decks through the office MCP server, including visual planning, rendering, and revision. Use for PPTX requests that explicitly use the office MCP tools; not for unrelated presentation workflows.
---

# Office presentation quality

The office MCP server writes and renders PowerPoint files. You supply the story, art direction, assets, and visual judgment. A successful tool call only proves the file was created.

## Choose the construction path

- If the user provides an existing PPTX to match, inspect it with `powerpoint_inspect_template`. Use `powerpoint_create_from_template` to reuse a layout or duplicate a styled slide into a separate output file. Use `powerpoint_native_batch` for edits the template tool cannot express.
- For a new narrative deck, start with `powerpoint_create_designed_presentation`. Its cover, image-text, statement, and comparison layouts are starting points, not a reason to force content into a template.
- Use `powerpoint_create_presentation` for custom compositions, charts, tables, and precise placement. Its coordinates are inches. Use image `fit: "cover"` or `"contain"`; avoid stretching photographs.
- Use `powerpoint_native_batch` when PowerPoint fidelity or existing objects matter. Its coordinates are points.

File-based native PowerPoint work uses hidden presentation windows by default. PowerPoint is single-instance on Windows: if background native work refuses because PowerPoint is already running, continue with portable creation or ask the user to close PowerPoint. Do not set `target.active: true` or `target.visible: true` merely to bypass the refusal without their consent.

## Before calling a creation tool

Decide the audience, purpose, and one main idea per slide. Preserve the user's requested subject and tone; do not silently replace it with a different premise. Choose a coherent visual direction. Prefer relevant images or an existing design reference over decorative circles, generic card grids, and repetitive timelines. Obtain or generate image files with available tools when imagery materially helps, and pass their local paths to the office tool. Keep data displays editable when the user needs to edit them.

Write direct, specific slide copy. Remove filler labels and vague slogans. Use comfortably readable text; do not solve overflow by shrinking body text below roughly 18 pt. Cut content or change the composition instead. Vary slide compositions when the story benefits from it.
Provide meaningful image alt text when an image conveys information; the audit flags missing descriptions.

## Motion when appropriate

For requested animations or transitions on Windows with desktop PowerPoint, use `powerpoint_native_batch`. Call `list_shapes` to obtain stable shape names, then `add_animation` or `set_transition`, save the deck, reopen it, and call `list_animations` to verify persistence. Use restrained motion that clarifies sequence or emphasis. The portable creation tools do not add animations. Static PNG renders do not show animation playback, so never claim to have visually verified motion from those images alone.

For cinematic Morph, camera-like pans or zooms, persistent objects, or a motion preview, also use the `office-morph-storytelling` skill.

## Validate the result

Call `powerpoint_audit_presentation` for structural review flags, then `powerpoint_render_presentation` and inspect every returned slide image at presentation size. Check legibility, cropping, alignment, contrast, hierarchy, content clarity, and whether the deck looks repetitive or generic. Correct visible issues and render again. The audit is heuristic, and a render call without reviewing the images is not validation. If the result remains weak, say so and ask for a reference or design direction rather than calling it polished.

Avoid replacing an existing user file unless they authorized it. Use a distinct draft path while iterating, then save the approved result to the requested location. Report the final PPTX path and any unverified visual limitations.
