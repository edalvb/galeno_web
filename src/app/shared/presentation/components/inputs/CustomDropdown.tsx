/**
 * CustomDropdown is a generic component that renders a dropdown using the Heroui Select component.
 * It supports both single and multiple selection modes.
 *
 * @template T - The type of the items.
 *
 * Props:
 * - label: (optional) The label displayed above the dropdown.
 * - items: An array of items to be rendered as options.
 * - renderValue: A function that takes an item and returns a string to be displayed.
 * - onValueChange: Callback function triggered when the selection changes.
 * - value: The current selected value(s). It can be a single value or an array of values.
 * - itemIdGetter: A function that returns the unique identifier of an item. This identifier is used internally for selection.
 * - selectionMode: (optional) Determines whether the selection is "single" or "multiple".
 * - parseSelectedKey: (optional) A function to convert the selected key (which is a string) back to its original type.
 *
 * Implementation Details:
 * - The component normalizes item keys to strings using the keyToString function to ensure proper selection management,
 *   as the Heroui Select component internally handles keys as strings.
 * - The selectedKeys are converted back to their original type using the parseSelectedKey function if provided.
 *   If not provided, a default conversion is applied that converts numeric strings to numbers or returns the original string.
 * - Conditional styling is applied to the SelectItem elements to visually indicate the selected items.
 *
 * Usage Example:
 *
 * const [selectedServices, setSelectedServices] = useState<number[]>([]);
 *
 * <CustomDropdown
 *    label="Services"
 *    placeholder="Select Services"
 *    items={store?.allServices ?? []}
 *    renderValue={(item) => item.name}
 *    onValueChange={(vals) => setSelectedServices(vals)}
 *    value={selectedServices}
 *    itemIdGetter={(item) => item.id}
 *    selectionMode="multiple"
 * />
 */

import { Select, SelectItem, SelectProps } from "@heroui/react";

interface CustomDropdownProps<T extends object>
  extends Omit<SelectProps, "renderValue" | "children"> {
  label?: string;
  items: T[];
  renderValue: (item: T) => string;
  onValueChange: (value: any) => void;
  value: any;
  itemIdGetter: (item: T) => any;
  selectionMode?: "single" | "multiple";
  parseSelectedKey?: (key: string) => any;
  validate?: SelectProps["validate"];
}

const keyToString = (key: any): string => {
  if (typeof key === "string") return key;
  if (typeof key === "number") return key.toString();
  return JSON.stringify(key);
};

const CustomDropdown = <T extends object>({
  label,
  items,
  renderValue,
  onValueChange,
  value,
  itemIdGetter,
  selectionMode,
  parseSelectedKey,
  validate,
  ...props
}: CustomDropdownProps<T>) => {
  const isMultiple = selectionMode === "multiple";

  const parseFn = parseSelectedKey
    ? parseSelectedKey
    : (key: string) => {
        const num = Number(key);
        return isNaN(num) ? key : num;
      };

  return (
    <Select
      label={label}
      placeholder={props.placeholder}
      variant="underlined"
      isRequired={props.isRequired}
      className={`transition-shadow ${props.className || ""}`}
      selectionMode={selectionMode}
      validate={validate}
      selectedKeys={
        isMultiple
          ? new Set((value as any[]).map(keyToString))
          : new Set(value !== undefined ? [keyToString(value)] : [])
      }
      onSelectionChange={(selectedKeys) => {
        const keys = selectedKeys as unknown as Set<string>;
        if (isMultiple) {
          onValueChange(Array.from(keys, parseFn));
        } else {
          const newValue =
            keys.size > 0 ? parseFn(Array.from(keys)[0]) : undefined;
          onValueChange(newValue);
        }
      }}
      classNames={{
        label: "text-sm font-semibold text-gray-700 mb-1.5",
        trigger: "border-b-1 border-[#E6E6F1]",
      }}
      {...props}
    >
      {items.map((item) => {
        const itemKeyStr = keyToString(itemIdGetter(item));
        return (
          <SelectItem
            key={itemKeyStr}
            className={
              isMultiple
                ? (value as any[]).map(keyToString).includes(itemKeyStr)
                  ? "bg-primary text-background"
                  : ""
                : keyToString(itemIdGetter(item)) === keyToString(value)
                ? "bg-primary text-background"
                : ""
            }
          >
            {renderValue(item)}
          </SelectItem>
        );
      })}
    </Select>
  );
};

export default CustomDropdown;
