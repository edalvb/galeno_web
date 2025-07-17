export class AppRoutesNamed {
  static readonly ROOT = "/";
  static readonly AUTH = "/auth/*";
  static readonly SIGNIN = "sign-in";
  static readonly SIGNUP = "sign-up";
  static readonly FORGOT_PASSWORD_REQUEST = "forgot-password/request";
  static readonly FORGOT_PASSWORD_VALIDATE = "forgot-password/validate";
  static readonly FORGOT_PASSWORD_RESET = "forgot-password/reset";
  static readonly ONBOARDING = "onboarding";
  static readonly ROLES = "roles/";
  static readonly DASHBOARD = "/dashboard/*";

  static readonly USER = "user";
  static readonly USER_DETAIL = `${AppRoutesNamed.USER}/detail/`;

  static readonly CONFIG_GENERAL = "configuration/general";
}

export class AppRoutesNamedFullPath {
  static readonly SIGNIN = "/auth/sign-in";
  static readonly SIGNUP = "/auth/sign-up";
  static readonly FORGOT_PASSWORD_REQUEST = "/auth/forgot-password/request";
  static readonly FORGOT_PASSWORD_VALIDATE = "/auth/forgot-password/validate";
  static readonly FORGOT_PASSWORD_RESET = "/auth/forgot-password/reset";

  static readonly ONBOARDING = "/onboarding";

  static readonly ROLES = "/roles/";

  static readonly USER = `/dashboard/${AppRoutesNamed.USER}`;
  static readonly USER_DETAIL = `/dashboard/${AppRoutesNamed.USER_DETAIL}`;

  static readonly CONFIG_GENERAL = `/dashboard/${AppRoutesNamed.CONFIG_GENERAL}`;
}
