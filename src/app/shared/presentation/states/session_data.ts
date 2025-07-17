import { UserSessionModel } from "@shared/domain";
import { create } from "zustand";

interface ISessionDataState {
  userSession?: UserSessionModel;
  initialize: (user: UserSessionModel) => void;
  reset: () => void;
}

export const sessionDataStateProvider = create<ISessionDataState>((set) => ({
  userSession: undefined,
  initialize: (user: UserSessionModel) => {
    set({
      userSession: user,
    });
  },
  reset: () => {
    set({
      userSession: undefined,
    });
  },
}));
