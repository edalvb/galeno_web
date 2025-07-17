import { SucursalModel } from "../../../domain/sucursal_model";

export class SucursalDto {
  cod_sucursal?: number;
  descripcion?: string;
  firma1?: string;
  firma2?: string;
  firma3?: string;
  cods_sucursal?: string;
  firma4?: string;

  constructor(data: Partial<SucursalDto>) {
    this.cod_sucursal = data.cod_sucursal;
    this.descripcion = data.descripcion;
    this.firma1 = data.firma1;
    this.firma2 = data.firma2;
    this.firma3 = data.firma3;
    this.cods_sucursal = data.cods_sucursal;
    this.firma4 = data.firma4;
  }

  fromJson(json: any): SucursalDto {
    return new SucursalDto({
      cod_sucursal: json.cod_sucursal,
      descripcion: json.descripcion,
      firma1: json.firma1,
      firma2: json.firma2,
      firma3: json.firma3,
      cods_sucursal: json.cods_sucursal,
      firma4: json.firma4,
    });
  }

  toModel(): SucursalModel {
    return new SucursalModel({
      cod_sucursal: this.cod_sucursal ?? 0,
      descripcion: this.descripcion ?? "",
      firma1: this.firma1 ?? "",
      firma2: this.firma2 ?? "",
      firma3: this.firma3 ?? "",
      cods_sucursal: this.cods_sucursal ?? "",
      firma4: this.firma4 ?? "",
    });
  }
}
