import { UsuarioModel } from "./usuario_model";
import { RolModel } from "./rol_model";
import { ModuloModel } from "./modulo_model";
import { PantallaModel } from "./pantalla_model";
import { SucursalModel } from "./sucursal_model";

export class UserSessionModel {
  user: UsuarioModel;
  roles: RolModel[];
  modulos: ModuloModel[];
  pantallas: PantallaModel[];
  sucursales: SucursalModel[];
  iat: number;
  exp: number;

  constructor(data: {
    user: UsuarioModel;
    roles: RolModel[];
    modulos: ModuloModel[];
    pantallas: PantallaModel[];
    sucursales: SucursalModel[];
    iat: number;
    exp: number;
  }) {
    this.user = data.user;
    this.roles = data.roles;
    this.modulos = data.modulos;
    this.pantallas = data.pantallas;
    this.sucursales = data.sucursales;
    this.iat = data.iat;
    this.exp = data.exp;
  }
}
