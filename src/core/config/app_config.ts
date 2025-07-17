import { ENV, Environment } from "../environment/environment";
import { AxiosClient } from "../network/axios_client";

export class AppConfig {
    static async initialize(): Promise<void> {
        Environment.env = ENV.DEV;
        AxiosClient.initialize();
    }
}