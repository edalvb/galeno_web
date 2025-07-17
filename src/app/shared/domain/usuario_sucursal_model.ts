export class UsuarioSucursalModel {
  usuario: string;
  cod_sucursal: number;

  constructor(data: { usuario: string; cod_sucursal: number }) {
    this.usuario = data.usuario;
    this.cod_sucursal = data.cod_sucursal;
  }
}
