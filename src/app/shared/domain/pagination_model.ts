export class PaginationModel {
  total: number;
  perPage: number;
  lastPage: number;
  currentPage: number;

  constructor(data: {
    total: number;
    perPage: number;
    lastPage: number;
    currentPage: number;
  }) {
    this.total = data.total;
    this.perPage = data.perPage;
    this.lastPage = data.lastPage;
    this.currentPage = data.currentPage;
  }

  static invalid(): PaginationModel {
    return new PaginationModel({
      total: 0,
      perPage: 0,
      lastPage: 0,
      currentPage: 0,
    });
  }
}
