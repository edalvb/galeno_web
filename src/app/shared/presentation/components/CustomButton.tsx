import { Button, ButtonProps } from "@heroui/react";

interface CustomButtonProps {
  onPress: () => void;
  radius?: ButtonProps["radius"];
  className?: string;
  size?: ButtonProps["size"];
  children?: React.ReactNode;
  disabled?: ButtonProps["disabled"];
}

const CustomButton = ({
  onPress,
  className = "",
  size = "lg",
  children = "Continue",
  disabled = false,
}: CustomButtonProps) => {
  return (
    <Button
      onPress={onPress}
      className={`
        w-full
        h-16
        bg-gradient-to-r from-primary to-primary-foreground
        text-white font-bold rounded-2xl shadow-sm
        hover:from-primary-foreground hover:to-primary
        focus:outline-none focus:ring-2 focus:ring-primary-foreground focus:ring-opacity-50
        ${className}
      `}
      size={size}
      isDisabled={disabled}
    >
      {children}
    </Button>
  );
};

export default CustomButton;
