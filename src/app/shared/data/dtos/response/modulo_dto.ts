import { ModuloModel } from "../../../domain/modulo_model";

export class ModuloDto {
  cod_modulo?: string;
  descripcion?: string;

  constructor(data: Partial<ModuloDto>) {
    this.cod_modulo = data.cod_modulo;
    this.descripcion = data.descripcion;
  }

  fromJson(json: any): ModuloDto {
    return new ModuloDto({
      cod_modulo: json.cod_modulo,
      descripcion: json.descripcion,
    });
  }

  toModel(): ModuloModel {
    return new ModuloModel({
      cod_modulo: this.cod_modulo ?? "",
      descripcion: this.descripcion ?? "",
    });
  }
}
