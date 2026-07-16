# DESIGN.md

## Design Strategy Brief

- Context: A church worship PowerPoint production tool used repeatedly by ministry workers. The interface should feel trustworthy, reverent, calm, and practical during time-sensitive weekly preparation.
- Industry fit: Use restrained liturgical cues rather than decorative religious imagery. Trust comes from stable controls, quiet hierarchy, legible status feedback, and consistent states. The PowerPoint canvas remains visually neutral because it represents the exported artifact rather than the application brand.
- Layout pattern: Preserve the existing three-column production workspace exactly: worship flow, content editor, and slide preview. Styling must not change grid widths, panel order, field order, or canvas geometry.
- Style direction: “Soft sanctuary” with warm paper surfaces, deep indigo structure, and muted antique-gold accents. Surfaces are layered through subtle tonal changes and restrained shadows rather than high contrast.
- Palette type: Deep indigo for primary actions and focus, warm ivory for editing surfaces, cool mist for navigation and preview surroundings, antique gold for selection and sacred warmth. Error and success colors remain muted enough for long sessions.
- Typography pairing: Keep the installed Microsoft JhengHei system stack to avoid loading delays and metric changes. Use weight, color, and letter spacing—not new font families—to establish hierarchy.
- UX checklist: Preserve keyboard focus visibility; distinguish hover, active, disabled, success, error, and busy states; announce status changes with a polite live region; use a clear but compact busy icon; support `prefers-reduced-motion`; verify desktop and narrow widths without horizontal overflow.
- Anti-patterns: Avoid pure-black panels, glowing neon colors, ornamental crosses, stained-glass decoration, large gradients, excessive blur, animated backgrounds, layout restructuring, and external icon dependencies.

## Design Summary

The application uses a calm liturgical visual layer while retaining its established production layout and workflow.

## Brand Voice

Reverent, composed, dependable, and quietly warm. Feedback should be direct and reassuring rather than playful.

## Color Tokens

- `--accent: #243b63`: primary indigo.
- `--accent-deep: #192b4a`: strong structure and overlays.
- `--surface: #fffdf8`: warm editing paper.
- `--mist: #eef0f3`: navigation and preview surroundings.
- `--gold: #a88b50`: restrained focus and selection accent.
- `--success: #3f6556`; `--danger: #8a433d`.

## Typography

Retain Microsoft JhengHei and existing font sizes. Establish hierarchy through weight and muted uppercase-style section labels.

## Spacing and Layout

Do not modify workspace grid columns, panel order, panel padding, field spacing, responsive breakpoints, or slide geometry.

## Components and CTA

Primary actions use deep indigo; secondary actions use warm paper with cool borders. Status messages distinguish idle, busy, success, and error. Busy operations show a small open-Bible emoji using the same gentle sway pattern as the New Family system's sprout animation, without shifting page layout.

## Responsive Rules

Keep existing breakpoints. The busy indicator remains inside viewport bounds at narrow widths.

## Accessibility Rules

Maintain visible keyboard focus, polite live announcements, `aria-busy`, sufficient contrast, disabled-state differentiation, and reduced-motion handling.

## Implementation Notes

Load `theme.css` after all existing styles so the change remains an isolated visual layer. Use the native open-Bible emoji with a CSS sway animation; do not add external image or font dependencies.
