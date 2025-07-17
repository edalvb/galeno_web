import React from "react";
import { IconClock } from "@tabler/icons-react";
import { TimeInput } from "@heroui/react";
import { Time } from "@internationalized/date";

interface CustomTimeInputProps {
  value?: Time;
  onChange: (time: Time) => void;
  label?: string;
  description?: string;
}

export const CustomTimeInput: React.FC<CustomTimeInputProps> = (
  arg: CustomTimeInputProps,
) => {
  return (
    <TimeInput
      variant="underlined"
      className="transition-shadow pt-4"
      description={arg.description}
      label={arg.label}
      defaultValue={arg.value}
      onChange={arg.onChange as any}
      labelPlacement="outside"
      classNames={{
        label: "text-sm font-semibold text-gray-700 mb-1.5",
        inputWrapper: "border-b-1 border-[#E6E6F1]",
      }}
      endContent={<IconClock />}
    />
  );
};
