---
name: office-morph-storytelling
description: Plan, build, and verify cinematic PowerPoint narratives with true Morph transitions through Office MCP. Use when a PPTX request calls for Morph, camera-like pans or zooms, persistent objects, staged movement, or a motion preview; not for ordinary static decks.
---

# Office Morph storytelling

Use Office MCP as the execution layer and treat consecutive slides as animation keyframes. A successful transition call is not enough: the shared objects, story rhythm, static frames, and exported playback must all be coherent.

## Plan before editing

Choose a conclusion-first, SCQA, or problem-solution story that fits the prompt. Keep one narrative job per slide. Before building, create a compact transition plan containing:

- the purpose of each slide pair;
- every persistent object's stable `!!` name;
- its start and end geometry;
- when it enters, leaves, or remains stationary;
- whether an intermediate keyframe is needed for sequential movement.

For detailed choreography rules and the operation pattern, read [references/morph-choreography.md](references/morph-choreography.md).

## Build with stable actors

- Create the first composed slide with `powerpoint_create_presentation`, a designed layout, or a user template.
- Prefer `duplicate_slide` for the next keyframe. Duplication preserves object identity and reduces accidental drift.
- Call `list_shapes`, then use `rename_shape` before setting Morph. Use `!!scene-*` for persistent background or framing objects and `!!actor-*` for content-bearing objects that transform.
- Names must be unique within a slide and identical only for objects intended to match across adjacent slides.
- Use `update_shape` to change position, size, rotation, text, or transparency. Use `set_z_order`, `group_shapes`, and `ungroup_shape` when composition requires them.
- If an object should animate out, keep its counterpart on the destination slide and move it fully outside the canvas. Delete it only when a fade is acceptable.
- Use an intermediate duplicated slide when objects must move at different times. Morph animates matched objects in one transition concurrently.
- Apply `set_transition` to destination slides. Use `morph` for object motion, `morphWords` for text-word changes, and `morphCharacters` only when character-level interpolation improves the result.

Keep file-based native work hidden. Do not set `target.visible=true` or attach to the active PowerPoint session merely to bypass a background-session refusal.

## Design motion, not decoration

Motion should establish hierarchy, reveal a relationship, preserve spatial context, or move the audience between sections. Avoid adding persistent ornaments solely to make the slide move. Use restrained travel distances and consistent directions unless the story intentionally changes direction.

Use real image assets and editable native text, charts, or tables where appropriate. Do not rely on Morph to rescue weak layouts, crowded copy, poor contrast, or irrelevant imagery.

## Verify the actual result

1. Save the deck, reopen it, and call `list_shapes` on each Morph pair. Confirm required `!!` names exist once on both slides.
2. Call `list_animations` on every destination slide. PowerPoint should report transition effect IDs `3954`, `3955`, or `3956` for object, word, or character Morph.
3. Run the PowerPoint audit and quality contract. Off-slide objects are acceptable only when they are deliberate animation exits and visually reviewed.
4. Render and inspect every slide at presentation size. Repair clipping, poor crops, unreadable text, unwanted remnants, and inconsistent hierarchy.
5. For motion-critical work, use `export_video`, verify completion status, and inspect frames near the middle and end of important transitions. Static slide renders do not verify motion.
6. Finalize only after every slide and prompt criterion has evidence-backed review notes.

When reconstructing a video or filmed presentation, describe the result as a faithful reconstruction unless the original PPTX and all source assets were supplied. Do not claim byte-for-byte or object-for-object identity from visible footage alone.
