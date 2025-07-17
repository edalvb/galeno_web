import { kRowsPerPage } from "@utils/constants";

export abstract class FiltersModel {
    abstract toQueryParams(): Record<string, any>;
}

export class FilterDefaultModel extends FiltersModel {
    page: number;
    per_page: number;
    no_pagination: boolean;

    constructor(page: number) {
        super();
        this.page = page;
        this.per_page = kRowsPerPage;
        this.no_pagination = false;
    }

    toQueryParams(): Record<string, any> {
        return {
            page: this.page,
            per_page: this.per_page,
            no_pagination: this.no_pagination,
        };
    }
}

export function formatQueryParams(filters: FiltersModel | null, page: number): Record<string, any> {
    const paginationParams = {
        per_page: kRowsPerPage,
        page: page,
    };

    if (!filters) {
        return {
            ...paginationParams,
        };
    }

    const queryParams = filters.toQueryParams();

    Object.keys(queryParams).forEach(key => {
        if (queryParams[key] === null || queryParams[key] === '') {
            delete queryParams[key];
        }
    });

    return {
        ...paginationParams,
        ...queryParams,
    };
}