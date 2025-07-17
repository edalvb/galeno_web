export enum ENV {
    DEV,
    QA,
    PROD,
}

export class Environment {

    static env: ENV;

    static get appConfig(): EnvConfig {
        switch (this.env) {
            case ENV.DEV:
                return {
                    env: ENV.DEV,
                    baseUrl: 'http://192.168.3.31:4565',
                    connectTimeout: 600000,
                    receiveTimeout: 600000
                }
            case ENV.QA:
                return {
                    env: ENV.QA,
                    baseUrl: 'http://192.168.0.38:4565',
                    connectTimeout: 600000,
                    receiveTimeout: 600000
                }
            case ENV.PROD:
                return {
                    env: ENV.PROD,
                    baseUrl: 'http://localhost:4565',
                    connectTimeout: 600000,
                    receiveTimeout: 600000
                }

        }
    }
}

export interface EnvConfig {
    env: ENV,
    baseUrl: string,
    connectTimeout: number,
    receiveTimeout: number,
}