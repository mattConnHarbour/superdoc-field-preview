import type { Editor } from "@superdoc-dev/react";
import type { SelectionTarget } from "superdoc/ui";
import type { CreateFieldInput, Field } from "./types";

type TransactionLike = {
  docChanged: boolean;
  selection: {
    $from: {
      parent: {
        textBetween(
          from: number,
          to: number,
          blockSeparator?: string,
          leafText?: string,
        ): string;
      };
      parentOffset: number;
    };
  };
};

type FieldTokenCandidate = {
  token: string;
  fieldName: string;
};

type FieldAutofillControllerOptions = {
  fields: Set<Field>;
  createField: (input: CreateFieldInput) => Field;
  getSelectionTarget: () => SelectionTarget | null;
  getFieldContent: (field: Field) => string;
  getFieldTag: (field: Field) => string;
  onMessage: (message: string) => void;
};

export class FieldAutofillController {
  private timer: number | null = null;
  private isConverting = false;

  constructor(private readonly options: FieldAutofillControllerOptions) {}

  handleTransaction(editor: Editor, transaction: TransactionLike): void {
    if (!transaction.docChanged || this.isConverting) return;

    const { $from } = transaction.selection;
    const textBeforeCaret = $from.parent.textBetween(
      0,
      $from.parentOffset,
      "",
      "\ufffc",
    );
    const tokenMatch = textBeforeCaret.match(/\{\{([^{}\r\n]+)\}\}$/);
    const fieldName = tokenMatch?.[1]?.trim();
    if (!tokenMatch || !fieldName) return;

    const candidate = { token: tokenMatch[0], fieldName };
    this.options.onMessage(`${candidate.token} detected…`);
    this.schedule(editor, candidate);
  }

  destroy(): void {
    if (this.timer !== null) {
      window.clearTimeout(this.timer);
      this.timer = null;
    }
  }

  private schedule(editor: Editor, candidate: FieldTokenCandidate): void {
    if (this.timer !== null) window.clearTimeout(this.timer);
    this.timer = window.setTimeout(() => {
      this.timer = null;
      void this.convert(editor, candidate);
    }, 25);
  }

  private async convert(
    editor: Editor,
    candidate: FieldTokenCandidate,
  ): Promise<void> {
    if (this.isConverting) return;

    const caret = this.options.getSelectionTarget();
    if (
      !caret ||
      caret.start.kind !== "text" ||
      caret.end.kind !== "text" ||
      caret.start.blockId !== caret.end.blockId ||
      caret.start.offset !== caret.end.offset ||
      caret.end.offset < candidate.token.length
    ) {
      this.options.onMessage(
        `${candidate.token} was detected, but its caret range could not be resolved.`,
      );
      return;
    }

    const tokenTarget: SelectionTarget = {
      kind: "selection",
      start: {
        ...caret.start,
        offset: caret.end.offset - candidate.token.length,
      },
      end: caret.end,
      ...(caret.story ? { story: caret.story } : {}),
    };

    const normalizedName = candidate.fieldName.toLocaleLowerCase();
    let field = [...this.options.fields].find(
      (item) => item.label.toLocaleLowerCase() === normalizedName,
    );
    if (!field) {
      field = this.options.createField({
        label: candidate.fieldName,
        placeholder: candidate.fieldName,
        value: candidate.fieldName,
      });
    }

    this.isConverting = true;
    try {
      const result = await editor.doc.create.contentControl({
        kind: "inline",
        controlType: "text",
        at: tokenTarget,
        content: this.options.getFieldContent(field),
        alias: field.label,
        tag: this.options.getFieldTag(field),
        lockMode: "unlocked",
      });
      if (result.success) {
        this.options.onMessage(
          `${candidate.token} converted to the ${field.label} field.`,
        );
      } else {
        this.options.onMessage(
          result.failure?.message ??
            `Could not convert ${candidate.token} to a field.`,
        );
      }
    } finally {
      this.isConverting = false;
    }
  }
}
