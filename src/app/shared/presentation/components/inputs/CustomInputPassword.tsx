import { Input } from "@heroui/react";
import { Icon, IconEye, IconEyeOff, IconProps } from "@tabler/icons-react";
import React, {
  ForwardRefExoticComponent,
  InputHTMLAttributes,
  RefAttributes,
} from "react";

type ValidationFunction = (
  value: string,
) => string | string[] | true | null | undefined;

interface CustomInputPasswordProps {
  value: InputHTMLAttributes<any>["value"];
  onChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  iconComponent?: ForwardRefExoticComponent<IconProps & RefAttributes<Icon>>;
  label?: string;
  placeholder?: string;
  validate?: ValidationFunction;
  additionalClassNames?: string;
  required?: boolean;
  disabled?: boolean;
}

export default function CustomInputPassword({
  value,
  onChange,
  label,
  placeholder,
  iconComponent: IconComponent,
  validate,
  additionalClassNames,
  required,
  disabled = false,
}: CustomInputPasswordProps) {
  const [isVisible, setIsVisible] = React.useState(false);

  const toggleVisibility = () => setIsVisible(!isVisible);

  return (
    <div className="transition-all duration-300 w-full">
      <Input
        size="md"
        variant="underlined"
        label={label}
        placeholder={placeholder}
        isRequired={required}
        value={value as any}
        disabled={disabled}
        onChange={onChange}
        startContent={
          IconComponent ? <IconComponent className="text-gray-400" /> : null
        }
        endContent={
          <button
            aria-label="toggle password visibility"
            className="focus:outline-none"
            type="button"
            onClick={toggleVisibility}
          >
            {isVisible ? (
              <IconEyeOff className="text-2xl text-default-400 pointer-events-none" />
            ) : (
              <IconEye className="text-2xl text-default-400 pointer-events-none" />
            )}
          </button>
        }
        className={`transition-shadow ${additionalClassNames}`}
        classNames={{
          label: "text-sm font-semibold text-gray-700 mb-1.5",
          inputWrapper: "border-b-1 border-[#E6E6F1]",
        }}
        type={isVisible ? "text" : "password"}
        validate={validate}
      />
    </div>
  );
}
