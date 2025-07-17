import { PaginationModel } from "@shared/domain/pagination_model";

export class ResponseWithList<T> {
  list: T[];
  pagination: PaginationModel;

  constructor(data: { list: T[]; pagination: PaginationModel }) {
    this.list = data.list;
    this.pagination = data.pagination;
  }
}
