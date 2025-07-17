import { Form, Modal, ModalContent, ModalProps } from "@heroui/react";
import PrimaryButton from "../buttons/Primary_button";
import SecondaryButton from "../buttons/Secondary_button";
import { ReactNode } from "react";

interface FormDialogProps {
  isOpen: boolean;
  onOpenChange: (isOpen: boolean) => void;
  onClose: () => void;
  title: string;
  form: ReactNode;
  doubleAction?: boolean;
  labelAcept?: string;
  labelCancel?: string;
  onTapCancel?: () => void;
  size?: ModalProps["size"];
  handleSubmit?: () => void;
}
export default function FormDialog({
  isOpen,
  onOpenChange,
  onClose,
  title,
  form,
  doubleAction = false,
  labelAcept = "Acept",
  labelCancel = "Cancel",
  onTapCancel = () => {
    onClose();
  },
  size = "sm",
  handleSubmit = () => {},
}: FormDialogProps) {
  return (
    <Modal
      isOpen={isOpen}
      onOpenChange={onOpenChange}
      onClose={onClose}
      size={size}
      className="primaryMainTheme"
      isDismissable={false}
      isKeyboardDismissDisabled={true}
    >
      <ModalContent className="bg-white flex flex-col items-start justify-start p-6 max-h-screen overflow-y-auto">
        <Form
          className="w-full"
          validationBehavior="native"
          onSubmit={(e) => {
            e.preventDefault();
            handleSubmit();
          }}
        >
          <p className="text-xl font-semibold text-gray-800">{title}</p>
          <div className="pt-6 pb-10 w-full">
            {typeof form === "string" ? (
              <p className="text-sm text-foreground mb-3">{form}</p>
            ) : (
              form
            )}
          </div>
          {doubleAction ? (
            <div className="mt-3 w-full flex flex-row space-x-4 justify-end">
              <div>
                <SecondaryButton
                  label={labelCancel}
                  onPress={onTapCancel}
                  small
                />
              </div>
              <div>
                <PrimaryButton type="submit" label={labelAcept} small />
              </div>
            </div>
          ) : (
            <div className="mt-3 w-full flex flex-row justify-end">
              <div>
                <PrimaryButton label={labelAcept} type="submit" small />
              </div>
            </div>
          )}
        </Form>
      </ModalContent>
    </Modal>
  );
}
