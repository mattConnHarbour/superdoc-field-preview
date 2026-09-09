import type { CreateFieldInput, Field, UpdateFieldInput } from "./types";

export class FieldController {
  private fields: Field[];

  constructor(initialFields: Field[] = []) {
    this.fields = initialFields.map((field) => ({ ...field }));
  }

  list(): Field[] {
    return this.fields.map((field) => ({ ...field }));
  }

  get(id: string): Field | undefined {
    const field = this.fields.find((candidate) => candidate.id === id);
    return field ? { ...field } : undefined;
  }

  getByName(name: string): Field | undefined {
    const normalizedName = name.trim().toLocaleLowerCase();
    const field = this.fields.find(
      (candidate) => candidate.label.toLocaleLowerCase() === normalizedName,
    );
    return field ? { ...field } : undefined;
  }

  create(input: CreateFieldInput): Field {
    const label = input.label.trim();
    if (!label) throw new Error("Field name is required.");

    const baseId =
      label
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-|-$/g, "") || "field";
    let id = baseId;
    let suffix = 2;
    while (this.fields.some((field) => field.id === id)) {
      id = `${baseId}-${suffix}`;
      suffix += 1;
    }

    const field = {
      id,
      label,
      placeholder: input.placeholder.trim() || label,
      value: input.value.trim(),
    };
    this.fields = [...this.fields, field];
    return { ...field };
  }

  update(id: string, input: UpdateFieldInput): Field | undefined {
    const current = this.fields.find((field) => field.id === id);
    if (!current) return undefined;

    const label = input.label?.trim() || current.label;
    const updated = {
      ...current,
      ...input,
      label,
      placeholder:
        input.placeholder === undefined
          ? current.placeholder
          : input.placeholder.trim() || label,
      value: input.value === undefined ? current.value : input.value.trim(),
    };
    this.fields = this.fields.map((field) =>
      field.id === id ? updated : field,
    );
    return { ...updated };
  }

  delete(id: string): boolean {
    const nextFields = this.fields.filter((field) => field.id !== id);
    if (nextFields.length === this.fields.length) return false;
    this.fields = nextFields;
    return true;
  }
}
