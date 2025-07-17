import { PaginationModel } from "./pagination_model";

export class ResponseInbox<T> {
  list: T[];
  pagination: PaginationModel;

  constructor(data: { list: T[]; pagination: PaginationModel }) {
    this.list = data.list;
    this.pagination = data.pagination;
  }
}
