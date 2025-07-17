export class SucursalModel {
  cod_sucursal: number;
  descripcion: string;
  firma1: string;
  firma2: string;
  firma3: string;
  cods_sucursal: string;
  firma4: string;

  constructor(data: {
    cod_sucursal: number;
    descripcion: string;
    firma1: string;
    firma2: string;
    firma3: string;
    cods_sucursal: string;
    firma4: string;
  }) {
    this.cod_sucursal = data.cod_sucursal;
    this.descripcion = data.descripcion;
    this.firma1 = data.firma1;
    this.firma2 = data.firma2;
    this.firma3 = data.firma3;
    this.cods_sucursal = data.cods_sucursal;
    this.firma4 = data.firma4;
  }
}
