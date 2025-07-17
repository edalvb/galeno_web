export class UsuarioRolModel {
  usuario: string;
  cod_rol: number;

  constructor(data: { usuario: string; cod_rol: number }) {
    this.usuario = data.usuario;
    this.cod_rol = data.cod_rol;
  }
}
