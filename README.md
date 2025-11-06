# Word Mendeley Citation Linker (VBA)

A VBA module for Microsoft Word that finds Mendeley/CSL citations (content controls / ADDIN fields) and converts them into intra-document hyperlinks to the bibliography entries.

## Features
- Detects bibliography area (content-control tag `mendeley_bibliography` or "References" heading)
- Builds bookmarks for bibliography entries
- Resolves citation tokens (content controls and ADDIN fields) and adds hyperlinks to reference bookmarks
- Heuristics for surname/year matching, normalized keys, and fallback resolution

## Getting started

1. Clone repository.
2. Open Word (Windows) and press `Alt+F11` to open VBA editor.
3. Import `src/LinkMendeleyCitations.bas` (File → Import File).
4. Open your document (make a copy first).
5. Run `LinkMendeleyCitations_Try2` from the macro list, or call the parameterized entry point.

## Recommended workflow
- Always run on a copy of your document.
- Check the Immediate window (Ctrl+G) for diagnostic messages and unmatched tokens summary.
- Use the configuration constants at the top of the module to tweak matching behavior.

## Examples
See `/examples` for a sample Word document and instructions.

## Contributing
Please read CONTRIBUTING.md for how to propose changes and run tests.

## License
GNU General Public License v3.0  
> This software is free: you can redistribute it and/or modify it under the terms of the GPL-3.0.  
