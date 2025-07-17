import {
  Table,
  TableBody,
  TableCell,
  TableCellProps,
  TableColumn,
  TableColumnProps,
  TableHeader,
  TableRow,
  useDisclosure,
} from "@heroui/react";
import CustomInput from "@shared/presentation/components/inputs/CustomInput";
import FormDialog from "@shared/presentation/components/dialogs/Form_dialog";
import PrimaryButton from "@shared/presentation/components/buttons/Primary_button";
import { IconSearch } from "@tabler/icons-react";
import { ReactNode, useState } from "react";
import LoadingAnimation from "@shared/presentation/components/loading/Loading_animation";
import EmptyTableState from "@shared/presentation/components/EmptyTableState";
import { TABLE_STYLES } from "@shared/presentation/styles/table_styles";
import toast from "react-hot-toast";
import ActionDialog from "@shared/presentation/components/dialogs/Action_dialog";

export interface ItemTableType<T> {
  key?: keyof T;
  tableProps?: TableColumnProps<T>;
  tableContent: ReactNode;
  cellProps?: TableCellProps;
  cellContent: (value: T) => ReactNode;
}

export interface SearchModalProps<T> {
  keyId?: keyof T;
  title?: string;
  labelInput?: string;
  placeHolderInput?: string;
  msgNoResults?: string;
  initialResults?: T[];
  initialSearchTerm?: string;
  initialSelectedItemId?: T[keyof T];
  itemsTableCells: ItemTableType<T>[];
  onClose: () => void;
  handleController: (value: string) => Promise<T[]>;
  handleDialogConfirm?: (value: T) => void;
  handleSearch?: (value: string) => void;
  msgErrorDefault?: string;
}

export function SearchModal<T>(prop: SearchModalProps<T>) {
  const [searchTerm, setSearchTerm] = useState(prop.initialSearchTerm || "");
  const [searchResults, setSearchResults] = useState<T[]>(
    prop.initialResults || [],
  );
  const [selectedItemId, setSelectedItemId] = useState<T[keyof T] | null>(
    prop.initialSelectedItemId || null,
  );
  const [isLoadingSearch, setIsLoadingSearch] = useState(false);
  const [messageError, setMessageError] = useState<string | null>(null);

  const {
    isOpen: isOpenActionDialogError,
    onOpen: onOpenActionDialogError,
    onOpenChange: onOpenChangeActionDialogError,
    onClose: onCloseActionDialogError,
  } = useDisclosure();

  const handleUserSelection = (value: T) => {
    if (value == null || prop.keyId == null) {
      return;
    }

    setSelectedItemId(value[prop.keyId]);
  };

  const handleDialogConfirm = () => {
    if (selectedItemId == null || !prop.keyId) {
      return;
    }

    if (prop.keyId == null) {
      return;
    }

    if (prop.handleDialogConfirm != null) {
      const key = prop.keyId;

      const fullItem = searchResults.find(
        (item) => item[key] === selectedItemId,
      );

      if (fullItem) {
        prop.handleDialogConfirm(fullItem);
      }
    }
  };

  const handleSearch = async () => {
    if (isLoadingSearch == true) {
      toast.error("Please wait until the search is completed.");

      return;
    }

    if (!searchTerm) {
      setSearchResults([]);
      return;
    }

    try {
      setIsLoadingSearch(true);
      const response = await prop.handleController(searchTerm);
      setIsLoadingSearch(false);

      setSearchResults(response);

      prop.handleSearch && prop.handleSearch(searchTerm);
    } catch (error) {
      console.error(error);

      const msgDefault =
        prop.msgErrorDefault || "Error searching. Please try again.";

      setMessageError(msgDefault);

      onOpenActionDialogError();
    }
  };

  return (
    <FormDialog
      isOpen={true}
      onOpenChange={prop.onClose}
      onClose={prop.onClose}
      title={prop.title || "Modal Search"}
      size="3xl"
      onTapCancel={prop.onClose}
      handleSubmit={selectedItemId == null ? prop.onClose : handleDialogConfirm}
      doubleAction={selectedItemId != null}
      labelCancel="Close"
      labelAcept={selectedItemId == null ? "Close" : "Accept"}
      form={
        <div className="w-full">
          <ActionDialog
            isOpen={isOpenActionDialogError}
            onOpenChange={onOpenChangeActionDialogError}
            onClose={onCloseActionDialogError}
            title="Error"
            subtitle={messageError}
          />
          <Table
            aria-label="Search results"
            classNames={TABLE_STYLES.classNames}
            topContent={
              <div className="mb-4 flex gap-3">
                <div className="flex-grow col-span-2">
                  <CustomInput
                    label={prop.labelInput || "Search"}
                    placeholder={prop.placeHolderInput || "Enter search term"}
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();

                        if (searchTerm) {
                          handleSearch();

                          return;
                        }
                      }
                    }}
                  />
                </div>
                <div className="flex items-end">
                  <PrimaryButton
                    label="Search"
                    small
                    onPress={handleSearch}
                    startContent={<IconSearch size={18} />}
                  />
                </div>
              </div>
            }
          >
            <TableHeader>
              {[
                <TableColumn>Selection</TableColumn>,
                ...prop.itemsTableCells.map((item) => (
                  <TableColumn key={item.tableProps?.key}>
                    {item.tableContent}
                  </TableColumn>
                )),
              ]}
            </TableHeader>
            <TableBody
              loadingContent={<LoadingAnimation />}
              loadingState={isLoadingSearch ? "loading" : "idle"}
              emptyContent={<EmptyTableState message={prop.msgNoResults} />}
            >
              {searchResults.map((element) => {
                const key = prop.keyId == null ? null : element[prop.keyId];
                const isSeleted = selectedItemId === key;

                return (
                  <TableRow
                    key={String(key)}
                    className={`cursor-pointer ${
                      isSeleted
                        ? "bg-primary bg-opacity-15"
                        : "hover:bg-gray-100 hover:bg-opacity-50"
                    }`}
                    onClick={() => handleUserSelection(element)}
                  >
                    {[
                      <TableCell className="w-20">
                        <div
                          className={`w-5 h-5 rounded-full border-2 flex-shrink-0 transition-colors mx-auto ${
                            isSeleted
                              ? "border-secondary bg-secondary"
                              : "border-light_box"
                          }`}
                        >
                          {isSeleted && (
                            <div className="w-full h-full flex items-center justify-center text-white">
                              <span className="text-xs">✓</span>
                            </div>
                          )}
                        </div>
                      </TableCell>,
                      ...prop.itemsTableCells.map((item) => (
                        <TableCell key={item.cellProps?.key}>
                          {item.cellContent(element)}
                        </TableCell>
                      )),
                    ]}
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
        </div>
      }
    />
  );
}
