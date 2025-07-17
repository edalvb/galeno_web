import { UserSessionModel } from "@shared/domain/user_session_model";
import { SucursalDto } from "./sucursal_dto";
import { PantallaDto } from "./pantalla_dto";
import { ModuloDto } from "./modulo_dto";
import { RolDto } from "./rol_dto";
import { UsuarioDto } from "./usuario_dto";

export class ResponseJwtDto {
  user?: UsuarioDto;
  roles?: RolDto[];
  modulos?: ModuloDto[];
  pantallas?: PantallaDto[];
  sucursales?: SucursalDto[];
  iat?: number;
  exp?: number;

  constructor(data: any) {
    this.user = data.user ? new UsuarioDto(data.user) : undefined;
    this.roles = data.roles
      ? data.roles.map((role: any) => new RolDto(role))
      : [];
    this.modulos = data.modulos
      ? data.modulos.map((modulo: any) => new ModuloDto(modulo))
      : [];
    this.pantallas = data.pantallas
      ? data.pantallas.map((pantalla: any) => new PantallaDto(pantalla))
      : [];
    this.sucursales = data.sucursales
      ? data.sucursales.map((sucursal: any) => new SucursalDto(sucursal))
      : [];
    this.iat = data.iat;
    this.exp = data.exp;
  }

  static fromJson(data: any): ResponseJwtDto {
    return new ResponseJwtDto(data);
  }

  toModel(): UserSessionModel {
    return new UserSessionModel({
      user: this.user?.toModel() ?? new UsuarioDto({}).toModel(),
      roles: (this.roles ?? []).map((role) => role.toModel()),
      modulos: (this.modulos ?? []).map((modulo) => modulo.toModel()),
      pantallas: (this.pantallas ?? []).map((pantalla) => pantalla.toModel()),
      sucursales: (this.sucursales ?? []).map((sucursal) => sucursal.toModel()),
      iat: this.iat ?? 0,
      exp: this.exp ?? 0,
    });
  }
}
