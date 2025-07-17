import { PaginationModel } from "../../../domain/pagination_model";

export class ResponseWrapperDto<T> {
    status?: "success" | "error" | null;
    message?: string;
    meta?: MetaDto;
    data?: T;
    pagination?: PaginationDto;
    error?: ErrorDto;

    constructor(
        json: any,
        fromJsonData?: (data: any) => T
    ) {
        this.status = json.status;
        this.message = json.message;
        this.meta = json.meta ? MetaDto.fromJson(json.meta) : undefined;
        this.data = json.data && fromJsonData ? fromJsonData(json.data) : json.data;
        this.pagination = json.pagination ? PaginationDto.fromJson(json.pagination) : undefined;
        this.error = json.error ? ErrorDto.fromJson(json.error) : undefined;
    }

    static fromJson<T>(json: any, fromJsonData: (data: any) => T): ResponseWrapperDto<T> {
        return new ResponseWrapperDto<T>(json, fromJsonData);
    }
}

export class MetaDto {
    requestId?: string;
    version?: string;
    timestamp?: Date;

    constructor(data: {
        requestId?: string;
        version?: string;
        timestamp?: Date;
    }) {
        this.requestId = data.requestId;
        this.version = data.version;
        this.timestamp = data.timestamp;
    }

    static fromJson(json: any): MetaDto {
        return new MetaDto({
            requestId: json.request_id,
            version: json.version,
            timestamp: json.timestamp ? new Date(json.timestamp) : undefined
        });
    }
}

export class PaginationDto {
    total?: number;
    perPage?: number;
    lastPage?: number;
    currentPage?: number;

    constructor(data: {
        total?: number;
        perPage?: number;
        lastPage?: number;
        currentPage?: number;
    }) {
        this.total = data.total;
        this.perPage = data.perPage;
        this.lastPage = data.lastPage;
        this.currentPage = data.currentPage;
    }

    static fromJson(json: any): PaginationDto {
        return new PaginationDto({
            total: json.total,
            perPage: json.per_page,
            lastPage: json.last_page,
            currentPage: json.current_page
        });
    }

    toJson(): any {
        return {
            total: this.total,
            per_page: this.perPage,
            last_page: this.lastPage,
            current_page: this.currentPage
        };
    }

    toModel(): PaginationModel {
        return new PaginationModel({
            total: this.total ?? 0,
            perPage: this.perPage ?? 0,
            lastPage: this.lastPage ?? 0,
            currentPage: this.currentPage ?? 0
        });
    }
}

export class ErrorDto {
    type?: string;
    title?: string;
    code?: number;
    message?: string;
    instance?: string;

    constructor(data: {
        type?: string;
        title?: string;
        code?: number;
        message?: string;
        instance?: string;
    }) {
        this.type = data.type;
        this.title = data.title;
        this.code = data.code;
        this.message = data.message;
        this.instance = data.instance;
    }

    static fromJson(json: any): ErrorDto {
        return new ErrorDto({
            type: json.type,
            title: json.title,
            code: json.code,
            message: json.message,
            instance: json.instance
        });
    }
}