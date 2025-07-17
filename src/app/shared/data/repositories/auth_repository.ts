import { injectable } from "inversify";
import { AxiosClient } from "@network/axios_client";
import { ResponseMutationModel } from "@shared/domain/response_mutation_model";
import { AxiosError } from "axios";
import { ResponseWrapperDto } from "@shared/data/dtos/response/response_wrapper_dto";
import { LoginRequestDto } from "../dtos/request/request_login_dto";

export abstract class IAuthRepository {
  abstract signIn(
    request: LoginRequestDto,
  ): Promise<ResponseMutationModel<string>>;
}

@injectable()
export class AuthRepository extends IAuthRepository {
  private axiosClient: AxiosClient = AxiosClient.instance;

  async signIn(
    request: LoginRequestDto,
  ): Promise<ResponseMutationModel<string>> {
    try {
      const response = await this.axiosClient.post({
        url: `/api/v1/auth`,
        body: request.toJson(),
      });

      if (response.status === 201) {
        const responseDto = new ResponseWrapperDto<string>(response.data);

        return new ResponseMutationModel({
          success: true,
          message: responseDto.message ?? "Login successful",
          data: responseDto.data,
        });
      } else {
        throw new Error(response.data.error.message);
      }
    } catch (error) {
      return new ResponseMutationModel({
        success: false,
        message:
          error instanceof AxiosError
            ? error.response?.data.error.message
            : "Error logging in. Please try again.",
      });
    }
  }
}
