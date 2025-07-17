import { LoginRequestDto } from "@shared/data/dtos/request/request_login_dto";
import { ResponseJwtDto } from "@shared/data/dtos/response/response_jwt_dto";
import { IAuthRepository } from "@shared/data/repositories/auth_repository";
import { RoleType } from "@shared/domain";
import { ResponseMutationModel } from "@shared/domain/response_mutation_model";
import { UserSessionModel } from "@shared/domain/user_session_model";
import { sessionDataStateProvider } from "@shared/presentation/states/session_data";
import { sl } from "@utils/dependency_injection";
import { Keys } from "@utils/keys";
import { jwtDecode } from "jwt-decode";

export class AuthService {
  static async doAuth(
    request: LoginRequestDto,
  ): Promise<ResponseMutationModel<string>> {
    const repository = sl.get<IAuthRepository>(Keys.AUTH_REPOSITORY);
    const response = await repository.signIn(request);
    if (response.success === true && response.data) {
      const userSession = this.processToken(response.data);
      this.writteUserDataStorage(response.data);
      this.writteUserDataSession(userSession);
    }
    return response;
  }

  static processToken(token: string): UserSessionModel {
    const decodedToken = jwtDecode(token);
    const userSession = ResponseJwtDto.fromJson(decodedToken).toModel();
    const normalizedUserSession = this.normalizeUserData(userSession);
    return normalizedUserSession;
  }

  // Normalization can happen when checking permissions, not necessarily on login
  static normalizeUserData(userSession: UserSessionModel): UserSessionModel {
    const roles = userSession.roles.filter(
      (role) =>
        role.descripcion === RoleType.ADMIN ||
        role.descripcion === RoleType.DOCTOR,
    );

    userSession.roles = roles;
    return userSession;
  }

  static writteUserDataStorage(token: string) {
    localStorage.setItem(Keys.AUTH_STATE, "true");
    localStorage.setItem(Keys.AUTH_TOKEN, token);
  }

  static writteUserDataSession(userSession: UserSessionModel) {
    sessionDataStateProvider.getState().initialize(userSession);
  }

  static logout() {
    sessionDataStateProvider.getState().reset();
    localStorage.removeItem(Keys.AUTH_STATE);
    localStorage.removeItem(Keys.AUTH_TOKEN);
    localStorage.removeItem(Keys.ROLE_ID);
  }

  static async getAuthState(): Promise<boolean> {
    const authState = localStorage.getItem(Keys.AUTH_STATE) === "true";
    if (!authState) {
      return false;
    }
    const token = localStorage.getItem(Keys.AUTH_TOKEN);
    if (!token) {
      return false;
    }

    // Check token expiration locally first
    try {
      const decoded: { exp: number } = jwtDecode(token);
      if (decoded.exp * 1000 < Date.now()) {
        console.log("Token expired locally.");
        this.logout(); // Logout if expired
        return false;
      }
    } catch (error) {
      console.error("Failed to decode token locally:", error);
      this.logout(); // Logout if token is invalid
      return false;
    }

    // Attempt to verify and refresh with the backend
    const repository = sl.get<IAuthRepository>(Keys.AUTH_REPOSITORY);
    const response = await repository.verifyAndRefreshToken();
    if (response.success === true && response.data) {
      const userSession = this.processToken(response.data);
      this.writteUserDataStorage(response.data); // Update token in storage
      this.writteUserDataSession(userSession);
      return true;
    }

    console.log("Token verification/refresh failed on backend.");
    this.logout(); // If refresh fails, log out
    return false;
  }

  static getRoleState(): boolean {
    const roleId = localStorage.getItem(Keys.ROLE_ID);
    if (!roleId) {
      return false;
    }

    const roles = sessionDataStateProvider.getState().userSession?.roles;
    if (!roles) {
      return false;
    }

    const roleSelected = roles.find(
      (role) => role.cod_rol === parseInt(roleId),
    );

    if (!roleSelected) {
      return false;
    }

    sessionDataStateProvider.getState().setRole(roleSelected);
    return true;
  }
}
