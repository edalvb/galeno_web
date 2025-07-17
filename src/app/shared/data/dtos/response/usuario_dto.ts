import { UsuarioModel } from "../../../domain/usuario_model";

export class UsuarioDto {
  usuario?: string;
  nombre?: string;
  contrasenha?: string;
  estado?: string;
  fecha_caduca?: Date;
  cedula?: string;
  fecha_registro?: Date;
  fec_mod_estado?: Date;
  session?: string;
  tipo?: string;
  socio_id?: string;

  constructor(data: Partial<UsuarioDto>) {
    this.usuario = data.usuario;
    this.nombre = data.nombre;
    this.contrasenha = data.contrasenha;
    this.estado = data.estado;
    this.fecha_caduca = data.fecha_caduca;
    this.cedula = data.cedula;
    this.fecha_registro = data.fecha_registro;
    this.fec_mod_estado = data.fec_mod_estado;
    this.session = data.session;
    this.tipo = data.tipo;
    this.socio_id = data.socio_id;
  }

  fromJson(json: any): UsuarioDto {
    return new UsuarioDto({
      usuario: json.usuario,
      nombre: json.nombre,
      contrasenha: json.contrasenha,
      estado: json.estado,
      fecha_caduca: json.fecha_caduca ? new Date(json.fecha_caduca) : undefined,
      cedula: json.cedula,
      fecha_registro: json.fecha_registro
        ? new Date(json.fecha_registro)
        : undefined,
      fec_mod_estado: json.fec_mod_estado
        ? new Date(json.fec_mod_estado)
        : undefined,
      session: json.session,
      tipo: json.tipo,
      socio_id: json.socio_id,
    });
  }

  toModel(): UsuarioModel {
    return new UsuarioModel({
      usuario: this.usuario ?? "",
      nombre: this.nombre ?? "",
      contrasenha: this.contrasenha ?? "",
      estado: this.estado ?? "",
      fecha_caduca: this.fecha_caduca ?? new Date(),
      cedula: this.cedula ?? "",
      fecha_registro: this.fecha_registro ?? new Date(),
      fec_mod_estado: this.fec_mod_estado ?? new Date(),
      session: this.session ?? "",
      tipo: this.tipo ?? "",
      socio_id: this.socio_id ?? "",
    });
  }
}
