# Quarto PPTX Layouts (Dynamic Sizing)

A Quarto extension that provides dynamic PowerPoint slide layout sizing using standard markdown classes like `{.smaller}` and `{.smallest}`.

## The Problem: Pandoc and PowerPoint

Quarto relies on Pandoc to generate PowerPoint presentations. While Pandoc's PowerPoint generation is incredibly useful, it has an important limitation: **You cannot natively choose arbitrary custom layouts from your markdown.**

Instead, Pandoc automatically maps your slide's content structure to a fixed set of **7 pre-defined layouts**. Pandoc infers the layout for each slide based on its content (e.g., whether it has two columns, images, etc.).

The 7 layout types Pandoc looks for in your `reference-doc` are:
1. **Title Slide**: The initial slide generated from YAML metadata.
2. **Section Header**: Slides with a level 1 header (`#`).
3. **Two Content**: Two-column slides containing only text.
4. **Comparison**: Two-column slides where one column contains an image.
5. **Content with Caption**: Single-column slides containing an image.
6. **Blank**: Slides with no visible content (e.g., only speaker notes).
7. **Title and Content**: The fallback standard layout for everything else.

Because of this hardcoded mapping, it is usually impossible to say: *"Use the Two Content layout, but make the text smaller on this specific slide."* 

## The Solution

This extension solves this limitation by introducing a two-step pipeline:
1. **A Lua Filter (`size_tags.lua`)**: Intercepts slides that have the `{.smaller}` or `{.smallest}` classes applied to their header and injects a hidden tag into the slide's speaker notes.
2. **A Post-Render Script (`apply-layouts.py`)**: Automatically scans the generated `.pptx` file, reads the hidden tags in the notes, changes the slide to the corresponding smaller/smallest custom layout, and removes the hidden tag so your notes stay clean.

## Installation & Setup

It is very easy to use this in your own Quarto project.

### Step 1: Install the extension
Open your terminal (or command prompt) and run:
```bash
quarto add David-Rieger/quarto-pptx-layouts
```

### Step 2: Activate it in your project
Open your `_quarto.yml` file (this is the configuration file of your project) and add the following lines so Quarto knows it should use the extension:

```yaml
filters:
  - quarto-pptx-layouts
  
project:
  post-render:
    - python _extensions/David-Rieger/quarto-pptx-layouts/apply-layouts.py
```

## Creating Your PPTX Template

For this trick to work, your PowerPoint template needs to actually have these smaller layouts prepared. Quarto allows you to link a custom PowerPoint template using the `reference-doc` option in your document's YAML header:

```yaml
format:
  pptx:
    reference-doc: my-template.pptx
```

### The Easiest Way: Use our ready-made template
You don't need to build this template yourself! Just look in the `example/` folder of this repository. There you will find a file called `template.pptx`. 

This is the absolute base template for Quarto to PowerPoint. It already contains all 7 standard layouts, **plus** a ` smaller` and ` smallest` version for each of them. 

You can just copy this `template.pptx` into your project folder and link it as your `reference-doc`. If you want to use your own company branding (like specific colors, fonts, or logos), you can simply open this `template.pptx` in PowerPoint, go to `View` -> `Slide Master`, and customize the designs there. The layout names are already perfectly set up for you!

### Doing it yourself (Advanced)
If you already have a complex custom template and want to add this functionality yourself:
1. Open your PowerPoint template.
2. Go to `View` -> `Slide Master`.
3. Find the 7 standard layouts mentioned above.
4. Duplicate them.
5. Make the text in the duplicates smaller (e.g., 6pt smaller).
6. Rename the duplicates by adding ` smaller` or ` smallest` to their name (e.g., name it `Two Content smaller`).

## Usage in your quarto slides

Simply append the `{.smaller}` or `{.smallest}` class to the header of the slide you want to resize. Pandoc will still infer the base layout, and our script will swap it to the sized variant.

> [!NOTE]
> This extension is logically designed for the 5 standard layouts that display body content. It is not intended for the **Title Slide** (which is generated directly from YAML metadata, not Markdown headers) or the **Blank** layout (which by definition contains no text to resize).

```markdown
## Standard Slide

This slide uses the normal "Title and Content" layout.

## Smallest Slide {.smallest}

This slide automatically gets mapped to the "Title and Content smallest" layout!
```

## Testing the Example

Check the `example/` folder in this repository. It contains a complete setup demonstrating all 7 layouts with the normal, smaller, and smallest classes.

You can test it by running:
```bash
cd example
quarto render example.qmd --to pptx
python ../apply-layouts.py example.pptx
```