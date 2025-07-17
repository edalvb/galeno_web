import { UsuarioSucursalModel } from "../../../domain/usuario_sucursal_model";

export class UsuarioSucursalDto {
  usuario?: string;
  cod_sucursal?: number;

  constructor(data: Partial<UsuarioSucursalDto>) {
    this.usuario = data.usuario;
    this.cod_sucursal = data.cod_sucursal;
  }

  fromJson(json: any): UsuarioSucursalDto {
    return new UsuarioSucursalDto({
      usuario: json.usuario,
      cod_sucursal: json.cod_sucursal,
    });
  }

  toModel(): UsuarioSucursalModel {
    return new UsuarioSucursalModel({
      usuario: this.usuario ?? "",
      cod_sucursal: this.cod_sucursal ?? 0,
    });
  }
}
