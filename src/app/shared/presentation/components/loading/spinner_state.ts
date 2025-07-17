import { create } from "zustand";

interface ISpinnerstate {
  showSpinner: boolean;
  setShowSpinner: (showSpinner: boolean) => void;
}

export const spinnerStateProvider = create<ISpinnerstate>((set) => ({
  showSpinner: false,
  setShowSpinner: (showSpinner: boolean) => set({ showSpinner }),
}));

export class Spinner {
  static show() {
    spinnerStateProvider.setState({ showSpinner: true });
  }

  static hide() {
    spinnerStateProvider.setState({ showSpinner: false });
  }
}
