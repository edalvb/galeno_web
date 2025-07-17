import { RolPantallaModel } from "../../../domain/rol_pantalla_model";

export class RolPantallaDto {
  cod_rol?: number;
  cod_pantalla?: string;

  constructor(data: Partial<RolPantallaDto>) {
    this.cod_rol = data.cod_rol;
    this.cod_pantalla = data.cod_pantalla;
  }

  fromJson(json: any): RolPantallaDto {
    return new RolPantallaDto({
      cod_rol: json.cod_rol,
      cod_pantalla: json.cod_pantalla,
    });
  }

  toModel(): RolPantallaModel {
    return new RolPantallaModel({
      cod_rol: this.cod_rol ?? 0,
      cod_pantalla: this.cod_pantalla ?? "",
    });
  }
}
