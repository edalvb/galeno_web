import { Chip } from "@heroui/react";
import React, { ReactNode } from "react";

interface CustomChipProps {
  color?: string;
  content: ReactNode;
}

export const CustomChip: React.FC<CustomChipProps> = (arg: CustomChipProps) => {
  return (
    <Chip color={arg.color as any} size="sm" className="text-xs">
      {arg.content || ""}
    </Chip>
  );
};
