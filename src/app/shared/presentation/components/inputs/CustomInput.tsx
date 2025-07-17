import { Input, InputProps, Textarea } from "@heroui/react";
import { Icon, IconProps } from "@tabler/icons-react";
import React, {
  ForwardRefExoticComponent,
  InputHTMLAttributes,
  RefAttributes,
} from "react";

type ValidationFunction = (
  value: string,
) => string | string[] | true | null | undefined;

interface CustomInputProps {
  value: InputHTMLAttributes<any>["value"];
  onChange: (
    event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => void;
  onKeyDown?: React.KeyboardEventHandler<
    HTMLInputElement | HTMLTextAreaElement
  >;
  iconComponent?: ForwardRefExoticComponent<IconProps & RefAttributes<Icon>>;
  endContent?: React.ReactNode;
  label?: string;
  placeholder?: string;
  type?: InputHTMLAttributes<any>["type"];
  validate?: ValidationFunction;
  additionalClassNames?: string;
  required?: boolean;
  multiline?: boolean;
  minLength?: number;
  maxLength?: number;
  disabled?: boolean;
  onBlur?: InputProps["onBlur"];
}

export default function CustomInput({
  value,
  onChange,
  onKeyDown,
  label,
  placeholder,
  iconComponent: IconComponent,
  endContent,
  type,
  validate,
  additionalClassNames,
  required,
  multiline,
  minLength,
  maxLength,
  disabled,
  onBlur,
}: CustomInputProps) {
  if (multiline) {
    return (
      <div className="transition-all duration-300 w-full">
        {label && (
          <label className="text-xs font-semibold text-gray-700 mb-1.5">
            {label}
          </label>
        )}
        <div className="relative">
          {IconComponent && (
            <div className="absolute left-3 top-3 text-gray-400">
              <IconComponent />
            </div>
          )}
          <Textarea
            size="md"
            variant="underlined"
            className={`transition-shadow ${additionalClassNames}`}
            placeholder={placeholder}
            required={required}
            value={value as any}
            onChange={onChange}
            onKeyDown={onKeyDown}
            minLength={minLength}
            maxLength={maxLength}
            disabled={disabled}
            onBlur={onBlur}
            rows={5}
            type={type}
            classNames={{
              label: "text-sm font-semibold text-gray-700 mb-1.5",
              inputWrapper: "border-b-1 border-[#E6E6F1]",
            }}
            endContent={endContent}
          />
        </div>
      </div>
    );
  }

  return (
    <div className="transition-all duration-300 w-full">
      <Input
        size="md"
        variant="underlined"
        label={label}
        placeholder={placeholder}
        isRequired={required}
        value={value as any}
        onChange={onChange}
        onKeyDown={onKeyDown}
        minLength={minLength}
        maxLength={maxLength}
        disabled={disabled}
        onBlur={onBlur}
        startContent={
          IconComponent ? <IconComponent className="text-gray-400" /> : null
        }
        className={`transition-shadow ${additionalClassNames}`}
        classNames={{
          label: "text-sm font-semibold text-gray-700 mb-1.5",
          inputWrapper: "border-b-1 border-[#E6E6F1]",
        }}
        type={type}
        validate={validate}
        endContent={endContent}
      />
    </div>
  );
}
