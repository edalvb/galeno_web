export class PantallaModel {
  cod_pantalla: string;
  descripcion: string;
  cod_modulo: string;
  alias: string;
  tipo_pantalla: string;

  constructor(data: {
    cod_pantalla: string;
    descripcion: string;
    cod_modulo: string;
    alias: string;
    tipo_pantalla: string;
  }) {
    this.cod_pantalla = data.cod_pantalla;
    this.descripcion = data.descripcion;
    this.cod_modulo = data.cod_modulo;
    this.alias = data.alias;
    this.tipo_pantalla = data.tipo_pantalla;
  }
}
