import type { DragEvent } from "react";
import type { Editor } from "@superdoc-dev/react";
import type { SelectionTarget, SuperDocUI } from "superdoc/ui";
import type { Field } from "./types";

type FieldDragDropControllerOptions = {
  fields: Set<Field>;
  getEditor: () => Editor | null;
  getUI: () => SuperDocUI | null;
  getFieldContent: (field: Field) => string;
  getFieldTag: (field: Field) => string;
  insertField: (field: Field, target: SelectionTarget) => Promise<void>;
  onDraggedFieldChange: (fieldId: string | null) => void;
  onDocumentDragChange: (isDraggingOver: boolean) => void;
  onMessage: (message: string) => void;
};

export class FieldDragDropController {
  private draggedFieldId: string | null = null;

  constructor(private readonly options: FieldDragDropControllerOptions) {}

  handleFieldDragStart(event: DragEvent<HTMLElement>, field: Field): void {
    event.dataTransfer.effectAllowed = "copy";
    event.dataTransfer.setData("application/x-superdoc-field", field.id);
    event.dataTransfer.setData("text/plain", field.label);
    this.setDraggedField(field.id);
    this.options.onMessage(
      `Drop ${field.label} where it should appear in the document.`,
    );
  }

  handleFieldDragEnd(): void {
    this.setDraggedField(null);
    this.options.onDocumentDragChange(false);
  }

  handleDocumentDragEnter(event: DragEvent<HTMLElement>): void {
    if (!this.draggedFieldId) return;
    event.preventDefault();
    this.options.onDocumentDragChange(true);
  }

  handleDocumentDragOver(event: DragEvent<HTMLElement>): void {
    if (!this.draggedFieldId) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    this.options.onDocumentDragChange(true);
  }

  handleDocumentDragLeave(event: DragEvent<HTMLElement>): void {
    if (!event.currentTarget.contains(event.relatedTarget as Node)) {
      this.options.onDocumentDragChange(false);
    }
  }

  handleDocumentDrop(event: DragEvent<HTMLElement>): void {
    event.preventDefault();
    this.options.onDocumentDragChange(false);

    const fieldId =
      event.dataTransfer.getData("application/x-superdoc-field") ||
      this.draggedFieldId;
    const field = [...this.options.fields].find((item) => item.id === fieldId);
    const ui = this.options.getUI();
    const hit = ui?.viewport.positionAt({
      x: event.clientX,
      y: event.clientY,
    });
    const existingFieldHit = ui?.viewport
      .entityAt({ x: event.clientX, y: event.clientY })
      .find(
        (entity) =>
          entity.type === "contentControl" &&
          entity.tag?.startsWith('{"fieldId":'),
      );

    this.setDraggedField(null);
    if (!field) return;
    if (!hit) {
      this.options.onMessage("Drop the field directly onto a document page.");
      return;
    }

    if (existingFieldHit?.type === "contentControl") {
      this.replaceExistingField(existingFieldHit, field);
      return;
    }

    void this.options.insertField(field, hit.target);
  }

  private replaceExistingField(
    hit: {
      id: string;
      scope?: "block" | "inline";
    },
    field: Field,
  ): void {
    const editor = this.options.getEditor();
    const target = {
      kind: hit.scope === "block" ? ("block" as const) : ("inline" as const),
      nodeType: "sdt" as const,
      nodeId: hit.id,
    };
    const patchResult = editor?.doc.contentControls.patch({
      target,
      alias: field.label,
      tag: this.options.getFieldTag(field),
    });
    const valueResult = editor?.doc.contentControls.text.setValue({
      target,
      value: this.options.getFieldContent(field),
    });

    this.options.onMessage(
      patchResult?.success && valueResult?.success
        ? `Existing field replaced with ${field.label}.`
        : "Could not replace the existing field.",
    );
  }

  private setDraggedField(fieldId: string | null): void {
    this.draggedFieldId = fieldId;
    this.options.onDraggedFieldChange(fieldId);
  }
}
