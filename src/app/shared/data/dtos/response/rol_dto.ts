import { RolModel, getRoleTypebyDescription, RoleType } from "../../../domain/rol_model";

export class RolDto {
  cod_rol?: number;
  descripcion?: RoleType;

  constructor(data: Partial<RolDto>) {
    this.cod_rol = data.cod_rol;
    this.descripcion = data.descripcion;
  }

  fromJson(json: any): RolDto {
    return new RolDto({
      cod_rol: json.cod_rol,
      descripcion: json.descripcion,
    });
  }

  toModel(): RolModel {
    return new RolModel({
      cod_rol: this.cod_rol ?? 0,
      descripcion: getRoleTypebyDescription(this.descripcion ?? ""),
    });
  }
}
