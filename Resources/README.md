# Document Emojis — C# resource pack

80 original glyphs at **48 × 48 px**, grouped into 8 categories. Bundled as PNG (raster) and SVG (vector source).

## Files

- `Resources/Emojis/{Section}/{name}.png` — 80 raster PNGs (48×48, transparent background)
- `Resources/Emojis/{Section}/{name}.svg` — 80 vector SVGs (matching set)
- `Resources/Emojis.resx` — drop-in WinForms / .NET resource file
- `Resources/Emoji.cs` — strongly-typed accessor (`Emoji.Status.Done`, `Emoji.Reactions.ThumbsUp`, etc.)
- `Resources/emoji-manifest.json` — programmatic index of every glyph

## Sections

| # | Section     | Emojis |
|---|-------------|--------|
| 1 | Status      | todo, in_progress, done, blocked, review, draft, on_hold, waiting, archive, new |
| 2 | Reactions   | thumbs_up, thumbs_down, heart, laugh, mind_blown, thinking, celebrate, eyes, raised_hand, comment |
| 3 | Actions     | edit, save, share, copy, delete, download, upload, search, link, print |
| 4 | Priority    | urgent, high, medium, low, pinned, flagged, bookmark, star, warning, locked |
| 5 | People      | person, team, mention, assign, invite, handshake, group_chat, handover, collaborators, request |
| 6 | Tools       | pen, pencil, clipboard, calendar, folder, document, chart, table, image, attachment |
| 7 | Pointers    | up, down, left, right, point, redo, undo, expand, refresh, callout |
| 8 | Decoration  | sparkle, fire, rocket, crown, trophy, award, idea, target, gem, confetti |

## Usage in C#

### Option A — `.resx` (WinForms / classic .NET)

1. Add `Resources/Emojis.resx` to your project.
2. Make sure the PNG files keep their relative path `Resources/Emojis/{Section}/{name}.png`.
3. Visual Studio auto-generates a `Emojis` class. Use it like:

```csharp
pictureBox1.Image = Emojis.Status_done;
pictureBox2.Image = Emojis.Reactions_thumbs_up;
```

### Option B — Embedded resources + `Emoji.cs` helper

1. Mark each PNG file's **Build Action** = `Embedded Resource` in your `.csproj`:

```xml
<ItemGroup>
  <EmbeddedResource Include="Resources\Emojis\**\*.png" />
</ItemGroup>
```

2. Set `Emoji.ResourceNamespace` to match your assembly's default namespace + folder path
   (default: `"DocumentEmojis.Resources.Emojis"`).
3. Use the strongly-typed accessor:

```csharp
using DocumentEmojis;

myButton.Image = Emoji.Actions.Save;
statusIcon.Image = Emoji.Status.InProgress;
toolbarItem.Image = Emoji.Decoration.Sparkle;
```

## Specifications

- **Size:** 48 × 48 px (PNG), scalable (SVG)
- **Background:** transparent
- **Style:** flat geometric with 1.5–2px outline accents
- **Palette:** sun `#FFC93C`, coral `#FF6B6B`, mint `#4ECDC4`, sky `#5FA8E8`, lavender `#9B8DEC`, peach `#FFB4A2`, leaf `#7BC47F`, charcoal `#2D3142`
