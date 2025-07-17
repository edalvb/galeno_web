export class RolModel {
  cod_rol: number;
  descripcion: RoleType;

  constructor(data: { cod_rol: number; descripcion: RoleType }) {
    this.cod_rol = data.cod_rol;
    this.descripcion = data.descripcion;
  }
}

export enum RoleType {
  ADMIN = "ADMIN",
  DOCTOR = "DOCTOR",
  CLIENT = "CLIENT",
}

export function getRoleTypebyDescription(description: string): RoleType {
  switch (description) {
    case "ADMIN":
      return RoleType.ADMIN;
    case "DOCTOR":
      return RoleType.DOCTOR;
    case "CLIENT":
      return RoleType.CLIENT;
    default:
      throw new Error(`Unknown role description: ${description}`);
  }
}

export const RoleTypeExtensions = {
  [RoleType.ADMIN]: "Administrator",
  [RoleType.DOCTOR]: "Doctor",
  [RoleType.CLIENT]: "Client",
};
