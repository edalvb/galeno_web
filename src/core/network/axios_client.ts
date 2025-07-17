import { Environment } from "../environment/environment";
import axios, { AxiosInstance, type AxiosRequestConfig, type AxiosResponse } from "axios";
import { Keys } from "../utils/keys";


export class AxiosClient {

    static instance: AxiosClient;

    private axiosClient: AxiosInstance | undefined;

    static initialize() {
        this.instance = new AxiosClient();
        this.instance.axiosClient = axios.create({
            baseURL: Environment.appConfig.baseUrl,
            timeout: Environment.appConfig.connectTimeout,
            responseType: "json",
            headers: {
                'Content-Type': 'application/json',
            }
        });
    }

    private getCommonConfig(params: AxiosBaseParams): AxiosRequestConfig {
        const token = localStorage.getItem(Keys.AUTH_TOKEN);
        return {
            headers: {
                Authorization: token ? `Bearer ${token}` : undefined,
                ...(params.isFormData && { 'Content-Type': 'multipart/form-data' }),
                ...params.headers,
            },
            params: params.queryParameters,
            responseType: params.responseType,
        };
    }

    async get(params: AxiosGetParams): Promise<AxiosResponse> {
        const config = this.getCommonConfig(params);
        return await this.axiosClient!.get(params.url, config);
    }

    async post(params: AxiosPostParams): Promise<AxiosResponse> {
        const config = this.getCommonConfig(params);
        const data = params.isFormData ? this.createFormData(params.body) : params.body;
        return await this.axiosClient!.post(params.url, data, config);
    }

    async put(params: AxiosPutParams): Promise<AxiosResponse> {
        const config = this.getCommonConfig(params);
        const data = params.isFormData ? this.createFormData(params.body) : params.body;
        return await this.axiosClient!.put(params.url, data, config);
    }

    async delete(params: AxiosDeleteParams): Promise<AxiosResponse> {
        const config = this.getCommonConfig(params);
        return await this.axiosClient!.delete(params.url, config);
    }

    private createFormData(data: any): FormData {
        const formData = new FormData();

        if (data) {
            Object.entries(data).forEach(([key, value]) => {
                if (Array.isArray(value)) {
                    value.forEach((file: any) => {
                        if (file instanceof File) {
                            formData.append(key, file);
                        }
                    });
                } else if (value instanceof File) {
                    formData.append(key, value);
                } else {
                    formData.append(key, value as string);
                }
            });
        }

        return formData;
    }

}

interface AxiosBaseParams {
    url: string;
    queryParameters?: any;
    headers?: Record<string, string>;
    isFormData?: boolean;
    responseType?: 'arraybuffer' | 'blob' | 'document' | 'json' | 'text' | 'stream';
}

export interface AxiosGetParams extends AxiosBaseParams { }

export interface AxiosPostParams extends AxiosBaseParams {
    body?: any;
}

export interface AxiosPutParams extends AxiosBaseParams {
    body?: any;
}

export interface AxiosDeleteParams extends AxiosBaseParams { }