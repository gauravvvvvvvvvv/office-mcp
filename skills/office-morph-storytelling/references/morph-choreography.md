# Morph choreography for Office MCP

Read this reference when a deck needs more than one simple Morph transition.

## Transition blueprint

Plan each adjacent pair before creating objects:

| Pair | Narrative beat | Stable actors | Geometry change | Exit or entry | Duration |
| --- | --- | --- | --- | --- | --- |
| 1→2 | Establish context | `!!scene-map` | centered → left, 100% → 70% | title enters | 0.8 s |
| 2→3 | Focus on evidence | `!!scene-map`, `!!actor-marker` | map zooms; marker grows | prior copy exits | 1.1 s |

Use names that communicate role rather than appearance alone. `!!actor-revenue-outlier` is more durable than `!!red-circle`.

## Operation pattern

The reliable native sequence is:

1. `list_shapes` on the composed source slide.
2. `rename_shape` for every object that must match.
3. `duplicate_slide` to create the destination keyframe.
4. `update_shape` on the duplicate.
5. Optionally use `set_z_order`, grouping, or transparency changes.
6. `set_transition` on the destination slide.
7. `save` to a draft path.
8. Reopen the draft and call `list_shapes` plus `list_animations`.

Example operation fragment:

```json
[
  { "op": "rename_shape", "slide": 1, "shape": "Picture 1", "name": "!!scene-map" },
  { "op": "duplicate_slide", "slide": 1, "toIndex": 2 },
  { "op": "update_shape", "slide": 2, "name": "!!scene-map", "left": -180, "top": -90, "width": 1320, "height": 742.5 },
  { "op": "set_transition", "slide": 2, "effect": "morph", "durationSeconds": 1.1, "advanceOnClick": true },
  { "op": "save", "outputPath": "draft-morph.pptx" }
]
```

Native coordinates are points. Portable creation coordinates are inches.

## Entering and exiting

- For a transform, keep the same stable name on both slides.
- For a fade-in, introduce an unmatched object on the destination slide.
- For a fade-out, omit the object from the destination slide.
- For a directional exit, retain the matching object and move it outside the destination canvas.
- Plan the return path if the actor will reappear later; do not accidentally leave off-slide remnants without narrative purpose.

## Staged movement

All matched actors within one Morph transition animate together. To stagger actions, add keyframes:

- Slide 2→3: actor A moves while actor B keeps identical geometry.
- Slide 3→4: actor B moves while actor A holds its new position or exits.

Prefer two intentional keyframes over combining Morph with a fragile collection of independent entrance effects.

## Camera moves

A full-slide or oversized image can act as a shared canvas:

- zoom in by enlarging it around the focal region;
- pan by changing `left` and `top` while preserving aspect ratio;
- zoom out into a card by reducing its size and adding supporting copy around it;
- move between regions of a large collage by combining scale and offset.

Keep essential subjects away from crop boundaries throughout the interpolation. Use a non-Morph cut when the spatial relationship would be misleading.

## Verification checklist

- Each destination slide has the intended Morph variant and duration.
- Every stable name occurs exactly once on each relevant slide.
- Unrelated objects do not share names.
- Persistent actors have meaningful displacement, scale, rotation, or style change.
- Intermediate frames do not expose blank edges from an oversized canvas.
- Text remains readable at the beginning and end of every transition.
- No prior-section actor remains visible unintentionally.
- Exported video timing matches the requested pace.
