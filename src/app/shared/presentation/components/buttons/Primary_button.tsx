import { Button } from "@heroui/react";
import { ReactNode } from "react";

interface PrimaryButtonProps {
  label: string;
  onPress?: () => void;
  type?: "button" | "submit" | "reset";
  className?: string;
  small?: boolean;
  startContent?: ReactNode;
  disabled?: boolean;
}

export default function PrimaryButton({
  label,
  onPress,
  type = "button",
  className,
  small,
  startContent,
  disabled,
}: PrimaryButtonProps) {
  return (
    <Button
      color="primary"
      variant={"shadow"}
      className={`${small ? "font-medium" : "font-semibold"} text-white ${
        small ? "py-2" : "py-7"
      } ${small ? "text-[13px]" : "text-base"} w-full
            bg-gradient-to-r from-primary to-secondary
            rounded-2xl shadow-sm
            focus:outline-none focus:ring-2 focus:ring-primary-foreground focus:ring-opacity-50 ${className}`}
      size={small ? "md" : "lg"}
      startContent={startContent}
      onPress={onPress}
      type={type}
      disabled={disabled}
      radius="md"
    >
      {label}
    </Button>
  );
}
