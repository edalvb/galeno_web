import { UsuarioRolModel } from "../../../domain/usuario_rol_model";

export class UsuarioRolDto {
  usuario?: string;
  cod_rol?: number;

  constructor(data: Partial<UsuarioRolDto>) {
    this.usuario = data.usuario;
    this.cod_rol = data.cod_rol;
  }

  fromJson(json: any): UsuarioRolDto {
    return new UsuarioRolDto({
      usuario: json.usuario,
      cod_rol: json.cod_rol,
    });
  }

  toModel(): UsuarioRolModel {
    return new UsuarioRolModel({
      usuario: this.usuario ?? "",
      cod_rol: this.cod_rol ?? 0,
    });
  }
}
