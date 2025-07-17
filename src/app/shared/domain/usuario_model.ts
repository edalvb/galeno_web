export class UsuarioModel {
  usuario: string;
  nombre: string;
  contrasenha: string;
  estado: string;
  fecha_caduca: Date;
  cedula: string;
  fecha_registro: Date;
  fec_mod_estado: Date;
  session: string;
  tipo: string;
  socio_id: string;

  constructor(data: {
    usuario: string;
    nombre: string;
    contrasenha: string;
    estado: string;
    fecha_caduca: Date;
    cedula: string;
    fecha_registro: Date;
    fec_mod_estado: Date;
    session: string;
    tipo: string;
    socio_id: string;
  }) {
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
}
