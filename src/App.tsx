import { useRef, useState } from "react";
import { SuperDocEditor } from "@superdoc-dev/react";
import type { Editor, SuperDocRef } from "@superdoc-dev/react";
import { createSuperDocUI } from "superdoc/ui";
import type { SelectionTarget, SuperDocUI } from "superdoc/ui";
import { FieldController } from "./FieldController";
import type { TemplateField } from "./FieldController";
import "@superdoc-dev/react/style.css";
import "./App.css";

type FieldDisplayMode = "placeholders" | "values";

const defaultFields: TemplateField[] = [
  {
    id: "user-name",
    label: "User name",
    placeholder: "User name",
    value: "John Doe",
  },
  {
    id: "email-address",
    label: "Email address",
    placeholder: "Email address",
    value: "john@example.com",
  },
  {
    id: "company-name",
    label: "Company name",
    placeholder: "Company name",
    value: "Acme, Inc.",
  },
  {
    id: "date",
    label: "Date",
    placeholder: "Date",
    value: "September 3, 2026",
  },
];

const fieldTag = (field: TemplateField) =>
  JSON.stringify({ fieldId: field.id });

const toolbarModules = {
  toolbar: {
    groups: {
      left: ["undo", "redo"],
      center: [
        "fontFamily",
        "fontSize",
        "bold",
        "italic",
        "underline",
        "color",
        "highlight",
        "link",
        "textAlign",
        "list",
        "numberedlist",
      ],
      right: ["zoom"],
    },
    responsiveToContainer: true,
  },
};

function FieldIcon() {
  return (
    <svg aria-hidden="true" viewBox="0 0 20 20">
      <path d="M4.5 3.5h11v13h-11zM7 7h6M7 10h6M7 13h3" />
    </svg>
  );
}

function App() {
  const editorRef = useRef<SuperDocRef>(null);
  const activeEditorRef = useRef<Editor | null>(null);
  const uiRef = useRef<SuperDocUI | null>(null);
  const insertionTargetRef = useRef<SelectionTarget | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [document, setDocument] = useState<string | File>("/mutual-NDA.docx");
  const [isReady, setIsReady] = useState(false);
  const [fieldController] = useState(() => new FieldController(defaultFields));
  const [fields, setFields] = useState(() => fieldController.list());
  const [fieldDisplayMode, setFieldDisplayMode] =
    useState<FieldDisplayMode>("placeholders");

  // ===== Field highlighting =====
  const [highlightSdts, setHighlightSdts] = useState(false);
  // ===== End field highlighting =====

  // ===== Drag and drop logic =====
  const [draggedFieldId, setDraggedFieldId] = useState<string | null>(null);
  const [isDraggingOverDocument, setIsDraggingOverDocument] = useState(false);
  // ===== End drag and drop logic =====
  const [editingFieldId, setEditingFieldId] = useState<string | null>(null);
  const [isCreatingField, setIsCreatingField] = useState(false);
  const [fieldDraft, setFieldDraft] = useState({
    label: "",
    placeholder: "",
    value: "",
  });
  const [message, setMessage] = useState(
    "Place your cursor in the document, then insert a field.",
  );

  const captureInsertionTarget = () => {
    const capture = uiRef.current?.selection.capture();
    const segment = capture?.target?.segments?.[0];
    const story = capture?.target?.story;
    if (!segment) return null;

    const point = {
      kind: "text" as const,
      blockId: segment.blockId,
      offset: segment.range.start,
      ...(story ? { story } : {}),
    };
    const target: SelectionTarget = {
      kind: "selection",
      start: point,
      end: point,
      ...(story ? { story } : {}),
    };
    insertionTargetRef.current = target;
    return target;
  };

  const insertField = async (
    field: TemplateField,
    explicitTarget?: SelectionTarget,
  ) => {
    const editor =
      editorRef.current?.getInstance()?.activeEditor ?? activeEditorRef.current;
    const target =
      explicitTarget ?? insertionTargetRef.current ?? captureInsertionTarget();

    if (!editor?.doc?.create?.contentControl) {
      setMessage(
        "The Content Controls API is still loading. Try again in a moment.",
      );
      return;
    }

    if (!target) {
      setMessage("Place the cursor in the document before inserting a field.");
      return;
    }

    const result = await editor.doc.create.contentControl({
      kind: "inline",
      controlType: "text",
      at: target,
      content: fieldDisplayMode === "values" ? field.value : field.placeholder,
      alias: field.label,
      tag: fieldTag(field),
      lockMode: "unlocked",
    });
    insertionTargetRef.current = null;

    if (result.success) {
      setMessage(`${field.label} inserted at the cursor.`);
    } else {
      setMessage(result.failure?.message ?? `Could not insert ${field.label}.`);
    }
  };

  // ===== Drag and drop logic =====
  const handleFieldDragStart = (
    event: React.DragEvent<HTMLElement>,
    field: TemplateField,
  ) => {
    event.dataTransfer.effectAllowed = "copy";
    event.dataTransfer.setData("application/x-superdoc-field", field.id);
    event.dataTransfer.setData("text/plain", field.label);
    setDraggedFieldId(field.id);
    setMessage(`Drop ${field.label} where it should appear in the document.`);
  };

  const handleDocumentDragOver = (event: React.DragEvent<HTMLElement>) => {
    if (!draggedFieldId) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    setIsDraggingOverDocument(true);
  };

  const handleDocumentDrop = (event: React.DragEvent<HTMLElement>) => {
    event.preventDefault();
    setIsDraggingOverDocument(false);

    const fieldId =
      event.dataTransfer.getData("application/x-superdoc-field") ||
      draggedFieldId;
    const field = fieldId ? fieldController.get(fieldId) : undefined;
    const hit = uiRef.current?.viewport.positionAt({
      x: event.clientX,
      y: event.clientY,
    });

    setDraggedFieldId(null);
    if (!field) return;
    if (!hit) {
      setMessage("Drop the field directly onto a document page.");
      return;
    }

    void insertField(field, hit.target);
  };
  // ===== End drag and drop logic =====

  // ===== Field value/placeholder toggle =====
  const setFieldMode = (mode: FieldDisplayMode) => {
    const doc =
      editorRef.current?.getInstance()?.activeEditor?.doc ??
      activeEditorRef.current?.doc;
    if (!doc) {
      setMessage(
        "The Content Controls API is still loading. Try again in a moment.",
      );
      return;
    }

    let updated = 0;
    for (const field of fields) {
      const value = mode === "values" ? field.value : field.placeholder;
      const controls = doc.contentControls.selectByTag({
        tag: fieldTag(field),
      }).items;

      for (const control of controls) {
        const result = doc.contentControls.text.setValue({
          target: control.target,
          value,
        });
        if (result.success) updated += 1;
      }
    }

    setFieldDisplayMode(mode);
    const label = mode === "values" ? "Field values" : "Field placeholders";
    setMessage(
      `${label} shown${updated ? ` in ${updated} field${updated === 1 ? "" : "s"}` : ""}.`,
    );
  };
  // ===== End field value/placeholder toggle =====

  const beginEditingField = (field: TemplateField) => {
    setIsCreatingField(false);
    setEditingFieldId(field.id);
    setFieldDraft({
      label: field.label,
      placeholder: field.placeholder,
      value: field.value,
    });
  };

  const beginCreatingField = () => {
    setEditingFieldId(null);
    setIsCreatingField(true);
    setFieldDraft({ label: "", placeholder: "", value: "" });
  };

  const saveField = () => {
    if (isCreatingField) {
      const label = fieldDraft.label.trim();
      if (!label) return;

      const newField = fieldController.create({
        label,
        placeholder: fieldDraft.placeholder,
        value: fieldDraft.value,
      });
      setFields(fieldController.list());
      setIsCreatingField(false);
      setMessage(`${label} added.`);
      return;
    }

    if (!editingFieldId) return;
    const field = fieldController.get(editingFieldId);
    if (!field) return;

    const updatedField = fieldController.update(field.id, {
      placeholder: fieldDraft.placeholder,
      value: fieldDraft.value,
    });
    if (!updatedField) return;
    setFields(fieldController.list());

    const doc =
      editorRef.current?.getInstance()?.activeEditor?.doc ??
      activeEditorRef.current?.doc;
    const displayedText =
      fieldDisplayMode === "values"
        ? updatedField.value
        : updatedField.placeholder;
    if (doc) {
      const controls = doc.contentControls.selectByTag({
        tag: fieldTag(updatedField),
      }).items;
      for (const control of controls) {
        doc.contentControls.text.setValue({
          target: control.target,
          value: displayedText,
        });
      }
    }

    setEditingFieldId(null);
    setMessage(`${field.label} updated.`);
  };

  const deleteField = () => {
    if (!editingFieldId) return;
    const field = fieldController.get(editingFieldId);
    if (!field || !fieldController.delete(editingFieldId)) return;

    setFields(fieldController.list());
    closeFieldEditor();
    setMessage(`${field.label} removed from the field library.`);
  };

  const closeFieldEditor = () => {
    setEditingFieldId(null);
    setIsCreatingField(false);
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setDocument(file);
    setIsReady(false);
    setMessage(`Loading ${file.name}…`);
    event.target.value = "";
  };

  const handleExport = async (mode: FieldDisplayMode) => {
    setFieldMode(mode);
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => resolve()),
    );
    await editorRef.current?.getInstance()?.export({ triggerDownload: true });
    setMessage(
      `Template exported with field ${mode === "values" ? "values" : "placeholders"} displayed.`,
    );
  };

  return (
    <div className={`app-shell ${highlightSdts ? "show-sdt-outlines" : ""}`}>
      <header className="app-header">
        <div>
          <h1>Template builder</h1>
          <p>Build reusable DOCX templates with structured fields.</p>
        </div>
        <div className="header-actions">
          {/* ===== Field highlighting ===== */}
          <button
            className={`sdt-highlight-toggle ${highlightSdts ? "active" : ""}`}
            type="button"
            aria-pressed={highlightSdts}
            disabled={!isReady}
            onClick={() => setHighlightSdts((current) => !current)}
          >
            Highlight SDTs
          </button>
          {/* ===== End field highlighting ===== */}

          {/* ===== Field value/placeholder toggle ===== */}
          <div className="field-mode-toggle" aria-label="Field display mode">
            <button
              className={fieldDisplayMode === "placeholders" ? "active" : ""}
              disabled={!isReady}
              onClick={() => setFieldMode("placeholders")}
            >
              Placeholders
            </button>
            <button
              className={fieldDisplayMode === "values" ? "active" : ""}
              disabled={!isReady}
              onClick={() => setFieldMode("values")}
            >
              Values
            </button>
          </div>
          {/* ===== End field value/placeholder toggle ===== */}
          <button
            className="button"
            onClick={() => fileInputRef.current?.click()}
          >
            Open DOCX
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".docx"
            hidden
            onChange={handleFileSelect}
          />
          <button
            className="button"
            disabled={!isReady}
            onClick={() => void handleExport("placeholders")}
          >
            Export with placeholders
          </button>
          <button
            className="button button-primary"
            disabled={!isReady}
            onClick={() => void handleExport("values")}
          >
            Export with values
          </button>
        </div>
      </header>

      <main className="workspace">
        {/* ===== Drag and drop logic ===== */}
        <section
          className={`document-workspace ${isDraggingOverDocument ? "drag-over" : ""}`}
          aria-label="Document editor"
          onDragEnter={(event) => {
            if (!draggedFieldId) return;
            event.preventDefault();
            setIsDraggingOverDocument(true);
          }}
          onDragOver={handleDocumentDragOver}
          onDragLeave={(event) => {
            if (!event.currentTarget.contains(event.relatedTarget as Node)) {
              setIsDraggingOverDocument(false);
            }
          }}
          onDrop={handleDocumentDrop}
        >
          <SuperDocEditor
            ref={editorRef}
            document={document}
            documentMode="editing"
            role="editor"
            modules={toolbarModules}
            user={{ name: "Template Author", email: "author@example.com" }}
            rulers
            contained
            onEditorCreate={({ editor }) => {
              activeEditorRef.current = editor;
            }}
            onEditorDestroy={() => {
              activeEditorRef.current = null;
              uiRef.current?.destroy();
              uiRef.current = null;
              insertionTargetRef.current = null;
            }}
            onReady={({ superdoc }) => {
              uiRef.current?.destroy();
              uiRef.current = createSuperDocUI({ superdoc });
              setIsReady(true);
              setMessage(
                "Place your cursor in the document, then insert a field.",
              );
            }}
            renderLoading={() => (
              <div className="loading-state">Loading editor…</div>
            )}
            style={{ height: "100%" }}
          />
        </section>
        {/* ===== End drag and drop logic ===== */}

        <aside className="field-sidebar">
          {editingFieldId || isCreatingField ? (
            <div className="field-edit-view">
              <button
                className="back-button"
                type="button"
                onClick={closeFieldEditor}
              >
                ← Back to fields
              </button>
              <div className="sidebar-heading">
                <span className="eyebrow">
                  {isCreatingField ? "New field" : "Edit field"}
                </span>
                <h2>
                  {isCreatingField
                    ? "Add a field"
                    : fieldController.get(editingFieldId ?? "")?.label}
                </h2>
                <p>
                  {isCreatingField
                    ? "Create a reusable template field."
                    : "Change what this field shows in each display mode."}
                </p>
              </div>

              <form
                className="field-edit-form"
                onSubmit={(event) => {
                  event.preventDefault();
                  saveField();
                }}
              >
                {isCreatingField && (
                  <label>
                    <span>Name</span>
                    <input
                      value={fieldDraft.label}
                      onChange={(event) =>
                        setFieldDraft((current) => ({
                          ...current,
                          label: event.target.value,
                        }))
                      }
                      placeholder="Contract amount"
                      required
                      autoFocus
                    />
                  </label>
                )}
                <label>
                  <span>Placeholder</span>
                  <input
                    value={fieldDraft.placeholder}
                    onChange={(event) =>
                      setFieldDraft((current) => ({
                        ...current,
                        placeholder: event.target.value,
                      }))
                    }
                    placeholder="Company name"
                    autoFocus={!isCreatingField}
                  />
                  <small>Shown when Placeholders is selected.</small>
                </label>
                <label>
                  <span>Value</span>
                  <input
                    value={fieldDraft.value}
                    onChange={(event) =>
                      setFieldDraft((current) => ({
                        ...current,
                        value: event.target.value,
                      }))
                    }
                    placeholder="Acme, Inc."
                  />
                  <small>Shown when Values is selected.</small>
                </label>
                <div className="field-edit-actions">
                  {!isCreatingField && (
                    <button
                      className="button delete-field-button"
                      type="button"
                      onClick={deleteField}
                    >
                      Delete
                    </button>
                  )}
                  <button
                    className="button"
                    type="button"
                    onClick={closeFieldEditor}
                  >
                    Cancel
                  </button>
                  <button className="button button-primary" type="submit">
                    {isCreatingField ? "Add field" : "Save field"}
                  </button>
                </div>
              </form>
            </div>
          ) : (
            <>
              <div className="sidebar-heading">
                <span className="eyebrow">Template</span>
                <div className="sidebar-title-row">
                  <h2>Fields</h2>
                  <button
                    className="add-field-button"
                    type="button"
                    onClick={beginCreatingField}
                  >
                    + Add field
                  </button>
                </div>
                <p>
                  Drag a field onto the document, or place the cursor and use
                  Insert. Fields can be reused.
                </p>
              </div>

              <div className="field-list">
                {fields.map((field) => (
                  <div
                    className={`field-card ${draggedFieldId === field.id ? "dragging" : ""}`}
                    key={field.id}
                  >
                    <span
                      className="field-icon field-drag-handle"
                      draggable={isReady}
                      title="Drag into the document"
                      aria-label={`Drag ${field.label} into the document`}
                      onDragStart={(event) =>
                        handleFieldDragStart(event, field)
                      }
                      onDragEnd={() => {
                        setDraggedFieldId(null);
                        setIsDraggingOverDocument(false);
                      }}
                    >
                      <FieldIcon />
                    </span>
                    <span className="field-copy">
                      <strong>{field.label}</strong>
                      <small>{field.value}</small>
                    </span>
                    <span className="field-card-actions">
                      <button
                        className="edit-field-button"
                        type="button"
                        onClick={() => beginEditingField(field)}
                      >
                        Edit
                      </button>
                      <button
                        className="insert-field-button"
                        disabled={!isReady}
                        onMouseDown={captureInsertionTarget}
                        onClick={() => void insertField(field)}
                      >
                        Insert
                      </button>
                    </span>
                  </div>
                ))}
              </div>

              <div className="sidebar-status" role="status">
                {message}
              </div>
            </>
          )}
        </aside>
      </main>
    </div>
  );
}

export default App;
