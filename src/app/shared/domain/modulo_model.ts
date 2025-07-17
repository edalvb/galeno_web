export class ModuloModel {
  cod_modulo: string;
  descripcion: string;

  constructor(data: { cod_modulo: string; descripcion: string }) {
    this.cod_modulo = data.cod_modulo;
    this.descripcion = data.descripcion;
  }
}
