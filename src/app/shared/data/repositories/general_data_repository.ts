import { injectable } from "inversify";
import { AxiosClient } from "@network/axios_client";
import { ResponseMutationModel } from "@shared/domain/response_mutation_model";
import { FolderNames } from "@shared/domain/folder_names";
import { AxiosError } from "axios";

export abstract class IGeneralDataRepository {
  abstract uploadFiles(
    folderName: FolderNames,
    files: File[],
  ): Promise<ResponseMutationModel<string[]>>;
}

@injectable()
export class GeneralDataRepository extends IGeneralDataRepository {
  private axiosClient: AxiosClient = AxiosClient.instance;

  async uploadFiles(
    folderName: FolderNames,
    files: File[],
  ): Promise<ResponseMutationModel<string[]>> {
    try {
      const response = await this.axiosClient.post({
        url: `/api/v1/images/upload-multiple/${folderName}`,
        isFormData: true,
        body: { files: Array.from(files) },
      });

      if (response.status === 201) {
        return new ResponseMutationModel({
          success: true,
          message: "Files uploaded successfully",
          data: response.data.data.map(
            (fileInfo: any) => fileInfo.fileName,
          ) as string[],
        });
      } else {
        throw new Error(response.data.error.message);
      }
    } catch (error) {
      return new ResponseMutationModel({
        success: false,
        message:
          error instanceof AxiosError
            ? error.response?.data?.error?.message
            : "Error uploading files. Please try again.",
      });
    }
  }
}
