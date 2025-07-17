import { Modal, ModalContent } from "@heroui/react";
import PrimaryButton from "../buttons/Primary_button";
import SecondaryButton from "../buttons/Secondary_button";
import { ReactNode } from "react";

interface ActionDialogProps {
  isOpen: boolean;
  onOpenChange: (isOpen: boolean) => void;
  onClose: () => void;
  title: string;
  subtitle: ReactNode;
  doubleAction?: boolean;
  labelAcept?: string;
  labelCancel?: string;
  onTapAcept?: () => void;
  onTapCancel?: () => void;
  size?: "sm" | "md" | "lg";
  primaryButtonType?: "submit" | "button";
}

export default function ActionDialog({
  isOpen,
  onOpenChange,
  onClose,
  title,
  subtitle = "",
  doubleAction = false,
  labelAcept = "Acept",
  labelCancel = "Cancel",
  onTapAcept = () => {
    onClose();
  },
  onTapCancel = () => {
    onClose();
  },
  size = "sm",
  primaryButtonType = "button",
}: ActionDialogProps) {
  return (
    <Modal
      isOpen={isOpen}
      onOpenChange={onOpenChange}
      onClose={onClose}
      size={size}
      className="greenLight"
    >
      <ModalContent className="bg-white flex flex-col items-start justify-start p-6">
        <p className="text-xl font-semibold text-gray-800">{title}</p>
        {typeof subtitle === "string" ? (
          <p className=" text-sm text-foreground mb-3">{subtitle}</p>
        ) : (
          subtitle
        )}
        {doubleAction ? (
          <div className="mt-3 w-full flex flex-row space-x-4">
            <div className="text-transparent w-[30%]">.</div>
            <SecondaryButton
              label={labelCancel}
              onPress={onTapCancel}
              small
            ></SecondaryButton>
            <PrimaryButton
              type={primaryButtonType}
              label={labelAcept}
              onPress={primaryButtonType === "submit" ? undefined : onTapAcept}
              small
            ></PrimaryButton>
          </div>
        ) : (
          <div className="mt-3 w-full flex flex-row justify-end">
            <div>
              <PrimaryButton label={labelAcept} onPress={onTapAcept} />
            </div>
          </div>
        )}
      </ModalContent>
    </Modal>
  );
}
