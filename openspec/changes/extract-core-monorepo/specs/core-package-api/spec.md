## ADDED Requirements

### Requirement: Core package exports DOCX parsing and serialization

The `@heyirisai/docx-editor-core` package SHALL export functions for parsing DOCX files into the document model and serializing the document model back to DOCX.

#### Scenario: Parse a DOCX file

- **WHEN** a consumer imports `parseDocx` from `@heyirisai/docx-editor-core`
- **AND** calls it with a DOCX file buffer
- **THEN** it SHALL return a parsed `Document` object matching the existing API

#### Scenario: Serialize a document to DOCX

- **WHEN** a consumer imports `serializeDocx` from `@heyirisai/docx-editor-core`
- **AND** calls it with a `Document` object
- **THEN** it SHALL return a DOCX file buffer matching the existing API

### Requirement: Core package exports document model types

The `@heyirisai/docx-editor-core` package SHALL export all document model types needed to work with parsed documents.

#### Scenario: Import document types

- **WHEN** a consumer imports types from `@heyirisai/docx-editor-core`
- **THEN** all types from `src/types/` SHALL be available (Paragraph, Run, Table, Style, Color, etc.)

### Requirement: Core package exports ProseMirror schema and extensions

The `@heyirisai/docx-editor-core` package SHALL export the ProseMirror schema, extension system, and all extensions for building an editor instance.

#### Scenario: Create a ProseMirror editor from core

- **WHEN** a consumer imports `createStarterKit` and `ExtensionManager` from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to build a ProseMirror schema and create an `EditorState`
- **AND** all existing extensions (bold, italic, table, list, etc.) SHALL be available

#### Scenario: Convert between document model and ProseMirror

- **WHEN** a consumer imports `toProseDoc` and `fromProseDoc` from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to convert between the document model and ProseMirror document format

### Requirement: Core package exports layout engine

The `@heyirisai/docx-editor-core` package SHALL export the layout engine for paginating documents and the layout painter for rendering pages to DOM.

#### Scenario: Use layout engine for pagination

- **WHEN** a consumer imports layout engine functions from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to paginate a ProseMirror document into pages

#### Scenario: Use layout painter for DOM rendering

- **WHEN** a consumer imports layout painter functions from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to render pages as vanilla DOM elements (no React required)

### Requirement: Core package exports headless API

The `@heyirisai/docx-editor-core` package SHALL export the `DocumentAgent` class for headless document manipulation without a browser.

#### Scenario: Use DocumentAgent headlessly

- **WHEN** a consumer imports `DocumentAgent` from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to programmatically manipulate DOCX files in Node.js without React or DOM dependencies

### Requirement: Core package exports utility functions

The `@heyirisai/docx-editor-core` package SHALL export utility functions for color resolution, unit conversion, font loading, template processing, and variable detection.

#### Scenario: Resolve theme colors

- **WHEN** a consumer imports color resolution utilities from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to resolve theme colors to hex values

#### Scenario: Process docxtemplater templates

- **WHEN** a consumer imports template processing utilities from `@heyirisai/docx-editor-core`
- **THEN** they SHALL be able to detect variables and process templates

### Requirement: React package does not re-export core

The `@heyirisai/docx-editor-react` package SHALL expose only React-owned adapter surfaces. Framework-agnostic document utilities, headless APIs, core plugins, and MCP APIs SHALL be imported from their canonical core or agents packages.

#### Scenario: React adapter imports stay adapter-scoped

- **WHEN** a user imports from `@heyirisai/docx-editor-react`
- **THEN** only React editor, React UI, React hooks, React dialogs, React plugin-api, and React style surfaces SHALL be available

#### Scenario: Canonical core imports

- **WHEN** a user needs document factories, headless processing, or core plugins
- **THEN** they SHALL import from `@heyirisai/docx-editor-core`, `@heyirisai/docx-editor-core/headless`, or `@heyirisai/docx-editor-core/core-plugins`
