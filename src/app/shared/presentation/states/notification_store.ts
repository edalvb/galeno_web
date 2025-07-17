import { create } from "zustand";

export interface INotificationState {
  // Export the interface
  unreadCount: number;
  isConnected: boolean; // Track WebSocket connection status
  setUnreadCount: (count: number) => void;
  incrementUnreadCount: () => void;
  decrementUnreadCount: () => void;
  resetUnreadCount: () => void;
  setConnectionStatus: (isConnected: boolean) => void;
}

export const useNotificationStore = create<INotificationState>((set) => ({
  unreadCount: 0,
  isConnected: false, // Default to disconnected
  setUnreadCount: (count) => set({ unreadCount: count >= 0 ? count : 0 }), // Ensure count doesn't go below 0
  incrementUnreadCount: () =>
    set((state) => ({ unreadCount: state.unreadCount + 1 })),
  decrementUnreadCount: () =>
    set((state) => ({ unreadCount: Math.max(0, state.unreadCount - 1) })), // Prevent negative count
  resetUnreadCount: () => set({ unreadCount: 0 }),
  setConnectionStatus: (isConnected) => set({ isConnected }),
}));
