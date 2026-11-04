# Third-party notices

The local Spanish spell checker uses:

- `nspell` 2.1.5, MIT License. Copyright (c) 2016 Titus Wormer. Its license text is included in `licenses/nspell-MIT.txt`.
- `dictionary-es` 4.0.0, Spanish dictionary by the LibreOffice/OpenOffice Spanish dictionary contributors, distributed under GPL-3.0-or-later, LGPL-3.0-or-later, or MPL-1.1-or-later. This project uses the dictionary under MPL 1.1; the upstream notice is included in `licenses/dictionary-es.txt`.
- `mammoth` 1.13.0, used for local `.docx` to HTML preview, BSD-2-Clause. Its license text is included in `licenses/mammoth-BSD-2-Clause.txt`.
- `@e965/xlsx` 0.20.3, used to read uploaded Excel files in the browser, Apache-2.0. Its license text is included in `licenses/xlsx-Apache-2.0.txt`.
- `cfb` 1.2.2, used to create embedded OLE package objects for browser-generated Excel files, Apache-2.0. Its license text is included in `licenses/xlsx-Apache-2.0.txt`.
- `docx` 9.15.2, used to generate per-issue Word reports in the browser, MIT.
- `jszip` 3.10.2, used to preserve the original Word report header in generated documents, MIT.

The dictionary is generated from [sbosio/rla-es](https://github.com/sbosio/rla-es) and is bundled locally so spell checking works without sending issue text to an external service.

The `esbuild` development dependency and Mammoth's bundled dependencies are distributed under their respective licenses, including MIT and BSD-2-Clause. Their license texts and notices are available with their packages in `node_modules` and at their upstream repositories.
