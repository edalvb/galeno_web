import { Button } from "@heroui/react";
import { ReactNode } from "react";

interface SecondaryButtonProps {
  label: string;
  onPress?: () => void;
  type?: "button" | "submit" | "reset";
  className?: string;
  small?: boolean;
  startContent?: ReactNode;
}

export default function SecondaryButton({
  label,
  onPress,
  type = "button",
  className,
  small,
  startContent,
}: SecondaryButtonProps) {
  return (
    <Button
      color="primary"
      variant={"bordered"}
      className={`${
        small ? "font-medium" : "font-semibold"
      } text-primary text-base ${small ? "py-2" : "py-7"} ${
        small ? "text-[13px]" : "text-base"
      } w-full
            rounded-2xl shadow-sm
            border-1
            focus:outline-none focus:ring-2 focus:ring-primary-foreground focus:ring-opacity-50 ${className}`}
      size={small ? "md" : "lg"}
      startContent={startContent}
      onPress={onPress}
      type={type}
      radius="md"
    >
      {label}
    </Button>
  );
}
