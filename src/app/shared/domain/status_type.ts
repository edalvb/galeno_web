export enum StatusType {
  ACTIVE = 1,
  INACTIVE = 2,
  DELETED = 3,
}

export const StatusTypeExtension = {
  getLabel: (status: StatusType): string => {
    switch (status) {
      case StatusType.ACTIVE:
        return "Active";
      case StatusType.INACTIVE:
        return "Inactive";
      case StatusType.DELETED:
        return "Deleted";
      default:
        return "Unknown";
    }
  },
  getColor: (status: StatusType): string => {
    switch (status) {
      case StatusType.ACTIVE:
        return "primary";
      case StatusType.INACTIVE:
        return "default";
      case StatusType.DELETED:
        return "danger";
      default:
        return "default";
    }
  },
};
