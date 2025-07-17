export class LoginRequestDto {
  user_cedula?: string;
  user_password?: string;

  constructor(data: Partial<LoginRequestDto>) {
    this.user_cedula = data.user_cedula;
    this.user_password = data.user_password;
  }

  toJson(): any {
    return {
      user_cedula: this.user_cedula,
      user_password: this.user_password,
    };
  }
}
