import { PantallaModel } from "../../../domain/pantalla_model";

export class PantallaDto {
  cod_pantalla?: string;
  descripcion?: string;
  cod_modulo?: string;
  alias?: string;
  tipo_pantalla?: string;

  constructor(data: Partial<PantallaDto>) {
    this.cod_pantalla = data.cod_pantalla;
    this.descripcion = data.descripcion;
    this.cod_modulo = data.cod_modulo;
    this.alias = data.alias;
    this.tipo_pantalla = data.tipo_pantalla;
  }

  fromJson(json: any): PantallaDto {
    return new PantallaDto({
      cod_pantalla: json.cod_pantalla,
      descripcion: json.descripcion,
      cod_modulo: json.cod_modulo,
      alias: json.alias,
      tipo_pantalla: json.tipo_pantalla,
    });
  }

  toModel(): PantallaModel {
    return new PantallaModel({
      cod_pantalla: this.cod_pantalla ?? "",
      descripcion: this.descripcion ?? "",
      cod_modulo: this.cod_modulo ?? "",
      alias: this.alias ?? "",
      tipo_pantalla: this.tipo_pantalla ?? "",
    });
  }
}
