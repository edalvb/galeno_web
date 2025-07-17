import { RolModel } from "@shared/domain";
import { UserSessionModel } from "@shared/domain/user_session_model";
import { create } from "zustand";

interface ISessionDataState {
  userSession?: UserSessionModel;
  userRole?: RolModel;
  setRole: (role: RolModel) => void;
  initialize: (user: UserSessionModel) => void;
  reset: () => void;
  companyId?: number;
}

export const sessionDataStateProvider = create<ISessionDataState>((set) => ({
  userSession: undefined,
  userRole: undefined,
  setRole: (role: RolModel) => {
    set({
      userRole: role,
    });
  },
  initialize: (user: UserSessionModel) => {
    set({
      userSession: user,
    });
  },
  reset: () => {
    set({
      userSession: undefined,
      userRole: undefined,
    });
  },
}));
