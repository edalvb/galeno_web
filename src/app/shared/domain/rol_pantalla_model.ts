export class RolPantallaModel {
  cod_rol: number;
  cod_pantalla: string;

  constructor(data: { cod_rol: number; cod_pantalla: string }) {
    this.cod_rol = data.cod_rol;
    this.cod_pantalla = data.cod_pantalla;
  }
}
