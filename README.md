# React template builder

A React template-authoring workflow built on `@superdoc-dev/react`. It provides a right-hand field library whose Insert buttons create inline Word content controls (SDTs) at the current editor cursor.

The demo starts with four reusable fields: User name, Email address, Company name, and Date. It also supports opening an existing DOCX and exporting the completed template.

The bundled Mutual NDA opens as the default document.

The header toggle switches inserted fields between their labels and sample values.
The Highlight SDTs toggle adds a temporary blue outline to template-field content controls without exposing hidden metadata-anchor SDTs or changing the exported document.
Each field's Edit action opens a sidebar form for changing its placeholder and value; saving also refreshes matching SDTs already in the document.
The Add field action creates additional reusable fields with their own name, placeholder, and value.

## Run

Install dependencies and start the app:

```bash
pnpm install
pnpm dev
```

Place the cursor in editable document text, then click Insert beside a field.

This project deliberately does not use `@superdoc-dev/template-builder`. Field insertion calls `editor.doc.create.contentControl(...)` directly through `@superdoc-dev/react`.

The editor dependencies are pinned to the latest v1 release line: `superdoc@1.46.3` and `@superdoc-dev/react@1.17.1`.
