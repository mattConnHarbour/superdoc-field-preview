export type Field = {
  id: string;
  label: string;
  placeholder: string;
  value: string;
};

export type CreateFieldInput = Omit<Field, "id">;
export type UpdateFieldInput = Partial<Omit<Field, "id">>;
