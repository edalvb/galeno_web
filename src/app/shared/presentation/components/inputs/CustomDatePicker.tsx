import { DatePicker, DatePickerProps } from "@heroui/react";
import { Util } from "@utils/utils";
import React from "react";
import { CalendarDate } from "@internationalized/date"; // Import CalendarDate
import { Granularity } from "@react-types/datepicker";

interface CustomDatePickerProps {
  value?: Date; // Input/Output as Date
  onChange?: (date?: Date) => void;
  label?: string;
  granularity?: Granularity;
  validate?: DatePickerProps["validate"];
  isRequired?: boolean;
  disabled?: boolean;
  minDate?: Date;
  maxDate?: Date;
}

export const CustomDatePicker: React.FC<CustomDatePickerProps> = (
  arg: CustomDatePickerProps,
) => {
  // Convert Date to CalendarDate for the DatePicker component
  const calendarValue = arg.value ? Util.dateToCalendarDate(arg.value) : null;
  const minCalendarValue = arg.minDate
    ? Util.dateToCalendarDate(arg.minDate)
    : undefined;
  const maxCalendarValue = arg.maxDate
    ? Util.dateToCalendarDate(arg.maxDate)
    : undefined;

  return (
    <DatePicker
      label={arg.label}
      value={calendarValue as any} // Pass CalendarDate
      onChange={(calendarDate: CalendarDate | any) => {
        // Convert CalendarDate back to Date for the parent component
        arg.onChange?.(
          calendarDate ? Util.calendarToDate(calendarDate) : undefined,
        );
      }}
      labelPlacement="outside"
      variant="underlined"
      validate={arg.validate && (arg.validate as any)}
      isRequired={arg.isRequired}
      isDisabled={arg.disabled}
      minValue={minCalendarValue}
      maxValue={maxCalendarValue}
      granularity={arg.granularity}
      className="transition-shadow"
      classNames={{
        label: "text-sm font-semibold text-gray-700 mb-1.5",
        inputWrapper: "border-b-1 border-[#E6E6F1]",
      }}
    />
  );
};
